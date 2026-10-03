using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace SageTypes
{
    // Acesso à duckdb.dll (API em C do DuckDB, banco analítico embutido): um banco de
    // sessão por processo do Excel, usado só pela thread de interface (onde o VBA roda).
    // O banco fica num arquivo temporário (apagado ao sair): as tabelas ficam comprimidas
    // e o DuckDB mantém na memória só o que está em uso, o que permite dezenas de milhões
    // de linhas mesmo com pouca RAM.
    static class Engine
    {
        static IntPtr database, connection;
        static string databaseFile;
        static readonly ConcurrentQueue<string> pendingDrops = new ConcurrentQueue<string>(); // tabelas de DataFrames já liberados
        static int tableCounter;

        // ------------------------------------------------------------------
        // API em C
        // ------------------------------------------------------------------

        [StructLayout(LayoutKind.Sequential)]
        internal struct Result
        {
            public ulong ColumnCount, RowCount, RowsChanged;
            public IntPtr Columns, ErrorMessage, InternalData;
        }

        const string Dll = "duckdb.dll";
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_open(byte[] path, out IntPtr database);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_connect(IntPtr database, out IntPtr connection);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern void duckdb_disconnect(ref IntPtr connection);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern void duckdb_close(ref IntPtr database);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_query(IntPtr connection, byte[] query, IntPtr result);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern void duckdb_destroy_result(IntPtr result);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr duckdb_result_error(IntPtr result);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern ulong duckdb_column_count(IntPtr result);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr duckdb_column_name(IntPtr result, ulong col);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_column_type(IntPtr result, ulong col);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr duckdb_fetch_chunk(Result result);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern ulong duckdb_data_chunk_get_size(IntPtr chunk);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr duckdb_data_chunk_get_vector(IntPtr chunk, ulong col);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr duckdb_vector_get_data(IntPtr vector);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr duckdb_vector_get_validity(IntPtr vector);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern void duckdb_destroy_data_chunk(ref IntPtr chunk);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_appender_create(IntPtr connection, byte[] schema, byte[] table, out IntPtr appender);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr duckdb_appender_error(IntPtr appender);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_appender_end_row(IntPtr appender);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_appender_close(IntPtr appender);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_appender_destroy(ref IntPtr appender);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_append_bool(IntPtr appender, [MarshalAs(UnmanagedType.I1)] bool value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_append_int64(IntPtr appender, long value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_append_double(IntPtr appender, double value);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_append_timestamp(IntPtr appender, long micros);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_append_date(IntPtr appender, int days);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_append_varchar_length(IntPtr appender, byte[] value, ulong length);
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)] static extern int duckdb_append_null(IntPtr appender);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern IntPtr LoadLibrary(string path);

        // Tipos (duckdb_type) lidos diretamente; os demais são convertidos no SQL (ver Readable)
        internal const int TBoolean = 1, TTinyInt = 2, TSmallInt = 3, TInteger = 4, TBigInt = 5, TUTinyInt = 6, TUSmallInt = 7,
            TUInteger = 8, TUBigInt = 9, TFloat = 10, TDouble = 11, TTimestamp = 12, TDate = 13, TVarchar = 17;

        // ------------------------------------------------------------------
        // Conexão
        // ------------------------------------------------------------------

        static void Open()
        {
            if (connection != IntPtr.Zero) return;
            // A duckdb.dll fica ao lado do SageTypes.dll; o Windows não a procura ali sozinho
            string folder = Path.GetDirectoryName(typeof(Engine).Assembly.Location);
            if (LoadLibrary(Path.Combine(folder, Dll)) == IntPtr.Zero)
                throw Interop.Error(53, "duckdb.dll não encontrada em " + folder + " (reinstale o SageTypes).");
            string temp = Path.Combine(Path.GetTempPath(), "SageDuckDB");
            Directory.CreateDirectory(temp);
            RemoveOrphans(temp);
            int pid = System.Diagnostics.Process.GetCurrentProcess().Id;
            databaseFile = Path.Combine(temp, "sage_" + pid.ToString(CultureInfo.InvariantCulture) + ".duckdb");
            DeleteDatabase(databaseFile); // de um processo anterior com o mesmo número
            if (duckdb_open(Utf8(databaseFile), out database) != 0) throw Interop.Error(5, "Não foi possível abrir o DuckDB em " + databaseFile + ".");
            if (duckdb_connect(database, out connection) != 0) throw Interop.Error(5, "Não foi possível conectar ao DuckDB.");
            AppDomain.CurrentDomain.ProcessExit += delegate { Close(); };
            AppDomain.CurrentDomain.DomainUnload += delegate { Close(); };

            // 40% da memória do computador (o resto fica para o Excel e o Windows); o que
            // passar disso fica só no disco
            ulong total = new Microsoft.VisualBasic.Devices.ComputerInfo().TotalPhysicalMemory;
            Run("SET memory_limit = '" + Math.Max(256, (long)(total * 2 / 5 / 1048576)) + "MB'");
            Run("SET temp_directory = " + Literal(temp));
            Run("SET preserve_insertion_order = true");
        }

        static void Close()
        {
            if (database == IntPtr.Zero) return;
            try
            {
                duckdb_disconnect(ref connection);
                duckdb_close(ref database);
            }
            catch (Exception) { }
            DeleteDatabase(databaseFile);
        }

        // Bancos de processos do Excel que terminaram sem fechar (travou, foi encerrado)
        static void RemoveOrphans(string folder)
        {
            foreach (string file in Directory.GetFiles(folder, "sage_*.duckdb"))
            {
                int pid;
                string id = Path.GetFileNameWithoutExtension(file).Substring(5);
                if (!int.TryParse(id, NumberStyles.None, CultureInfo.InvariantCulture, out pid)) continue;
                bool running;
                try { running = !System.Diagnostics.Process.GetProcessById(pid).HasExited; }
                catch (ArgumentException) { running = false; }
                catch (Exception) { running = true; } // sem acesso: pode estar em uso
                if (!running) DeleteDatabase(file);
            }
        }

        static void DeleteDatabase(string file)
        {
            foreach (string f in new[] { file, file + ".wal" })
            {
                try { if (File.Exists(f)) File.Delete(f); }
                catch (Exception) { } // em uso por outro processo
            }
        }

        // Nome novo para uma tabela de DataFrame
        public static string NewTable()
        {
            tableCounter++;
            return "sage_df_" + tableCounter.ToString(CultureInfo.InvariantCulture);
        }

        // Chamado pelo finalizador (outra thread): só agenda
        public static void ScheduleDrop(string table) { pendingDrops.Enqueue(table); }

        static void DropPending()
        {
            string table;
            while (pendingDrops.TryDequeue(out table))
            {
                using (QueryResult r = RawQuery("DROP TABLE IF EXISTS " + Quote(table))) { }
            }
        }

        // ------------------------------------------------------------------
        // Consultas
        // ------------------------------------------------------------------

        public static void Run(string sql)
        {
            using (Query(sql)) { }
        }

        public static QueryResult Query(string sql)
        {
            Open();
            DropPending();
            return RawQuery(sql);
        }

        static QueryResult RawQuery(string sql)
        {
            QueryResult r = new QueryResult();
            if (duckdb_query(connection, Utf8(sql), r.Handle) != 0)
            {
                string message = Text(duckdb_result_error(r.Handle));
                r.Dispose();
                throw Interop.Error(5, "DuckDB: " + message);
            }
            r.Load();
            return r;
        }

        // Um único valor (count(*), sum(...))
        public static object Scalar(string sql)
        {
            using (QueryResult r = Query(sql))
            {
                List<object[]> rows = r.ReadAll(1);
                return rows.Count == 0 || rows[0].Length == 0 ? null : rows[0][0];
            }
        }

        // Resultado de uma consulta: colunas e leitura em blocos
        public sealed class QueryResult : IDisposable
        {
            public readonly IntPtr Handle = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Result)));
            public string[] Names = new string[0];
            public int[] Types = new int[0];
            bool disposed;

            public QueryResult()
            {
                for (int i = 0; i < Marshal.SizeOf(typeof(Result)); i++) Marshal.WriteByte(Handle, i, 0);
            }

            internal void Load()
            {
                int n = (int)duckdb_column_count(Handle);
                Names = new string[n];
                Types = new int[n];
                for (int c = 0; c < n; c++)
                {
                    Names[c] = Text(duckdb_column_name(Handle, (ulong)c));
                    Types[c] = duckdb_column_type(Handle, (ulong)c);
                }
            }

            // Próximo bloco de linhas (até 2048), por coluna; null no fim
            public object[][] NextChunk(out int rows)
            {
                rows = 0;
                Result value = (Result)Marshal.PtrToStructure(Handle, typeof(Result));
                IntPtr chunk = duckdb_fetch_chunk(value);
                if (chunk == IntPtr.Zero) return null;
                try
                {
                    rows = (int)duckdb_data_chunk_get_size(chunk);
                    object[][] columns = new object[Types.Length][];
                    for (int c = 0; c < Types.Length; c++)
                        columns[c] = ReadVector(duckdb_data_chunk_get_vector(chunk, (ulong)c), Types[c], rows);
                    return columns;
                }
                finally { duckdb_destroy_data_chunk(ref chunk); }
            }

            // Todas as linhas (até max), cada uma um array de valores
            public List<object[]> ReadAll(long max)
            {
                List<object[]> result = new List<object[]>();
                int rows;
                object[][] columns;
                while (result.Count < max && (columns = NextChunk(out rows)) != null)
                {
                    for (int r = 0; r < rows && result.Count < max; r++)
                    {
                        object[] row = new object[columns.Length];
                        for (int c = 0; c < columns.Length; c++) row[c] = columns[c][r];
                        result.Add(row);
                    }
                }
                return result;
            }

            public void Dispose()
            {
                if (disposed) return;
                disposed = true;
                duckdb_destroy_result(Handle);
                Marshal.FreeHGlobal(Handle);
            }
        }

        static readonly DateTime Epoch = new DateTime(1970, 1, 1);

        static object[] ReadVector(IntPtr vector, int type, int rows)
        {
            object[] values = new object[rows];
            IntPtr data = duckdb_vector_get_data(vector);
            IntPtr validity = duckdb_vector_get_validity(vector);
            for (int r = 0; r < rows; r++)
            {
                if (validity != IntPtr.Zero && ((Marshal.ReadInt64(validity, (r / 64) * 8) >> (r % 64)) & 1) == 0) continue; // NULL
                switch (type)
                {
                    case TBoolean: values[r] = Marshal.ReadByte(data, r) != 0; break;
                    case TTinyInt: values[r] = (int)(sbyte)Marshal.ReadByte(data, r); break;
                    case TSmallInt: values[r] = (int)Marshal.ReadInt16(data, r * 2); break;
                    case TInteger: values[r] = Marshal.ReadInt32(data, r * 4); break;
                    case TBigInt: values[r] = Whole(Marshal.ReadInt64(data, r * 8)); break;
                    case TUTinyInt: values[r] = (int)Marshal.ReadByte(data, r); break;
                    case TUSmallInt: values[r] = (int)(ushort)Marshal.ReadInt16(data, r * 2); break;
                    case TUInteger: values[r] = Whole((uint)Marshal.ReadInt32(data, r * 4)); break;
                    case TUBigInt: values[r] = (double)(ulong)Marshal.ReadInt64(data, r * 8); break;
                    case TFloat: values[r] = (double)BitConverter.ToSingle(BitConverter.GetBytes(Marshal.ReadInt32(data, r * 4)), 0); break;
                    case TDouble: values[r] = BitConverter.Int64BitsToDouble(Marshal.ReadInt64(data, r * 8)); break;
                    case TDate: values[r] = Epoch.AddDays(Marshal.ReadInt32(data, r * 4)); break;
                    case TTimestamp: values[r] = Epoch.AddTicks(Marshal.ReadInt64(data, r * 8) * 10); break;
                    case TVarchar: values[r] = ReadString(data + r * 16); break;
                }
            }
            return values;
        }

        // Inteiro: Long do VBA quando cabe; senão LongLong
        static object Whole(long v) { return v >= int.MinValue && v <= int.MaxValue ? (object)(int)v : v; }

        // duckdb_string_t: até 12 bytes dentro da própria estrutura; acima disso, um ponteiro
        static string ReadString(IntPtr s)
        {
            int length = Marshal.ReadInt32(s);
            if (length == 0) return "";
            byte[] bytes = new byte[length];
            Marshal.Copy(length <= 12 ? s + 4 : Marshal.ReadIntPtr(s, 8), bytes, 0, length);
            return Encoding.UTF8.GetString(bytes);
        }

        // ------------------------------------------------------------------
        // Inserção em massa
        // ------------------------------------------------------------------

        public sealed class Appender : IDisposable
        {
            IntPtr handle;

            public Appender(string table)
            {
                Open();
                if (duckdb_appender_create(connection, null, Utf8(table), out handle) != 0)
                    throw Interop.Error(5, "DuckDB: " + Text(duckdb_appender_error(handle)));
            }

            // Valor convertido para o tipo da coluna (ver Kind)
            public void Add(object value, int kind)
            {
                int state;
                if (value == null || value is DBNull || Interop.IsMissing(value)) state = duckdb_append_null(handle);
                else
                {
                    switch (kind)
                    {
                        case TBoolean: state = duckdb_append_bool(handle, Convert.ToBoolean(value, CultureInfo.InvariantCulture)); break;
                        case TBigInt: state = duckdb_append_int64(handle, Convert.ToInt64(value, CultureInfo.InvariantCulture)); break;
                        case TDouble: state = duckdb_append_double(handle, Convert.ToDouble(value, CultureInfo.InvariantCulture)); break;
                        case TTimestamp:
                            state = duckdb_append_timestamp(handle, (Convert.ToDateTime(value, CultureInfo.InvariantCulture) - Epoch).Ticks / 10);
                            break;
                        case TDate:
                            state = duckdb_append_date(handle, (int)(Convert.ToDateTime(value, CultureInfo.InvariantCulture).Date - Epoch).TotalDays);
                            break;
                        default:
                            byte[] bytes = Encoding.UTF8.GetBytes(Interop.Text(value));
                            state = duckdb_append_varchar_length(handle, bytes, (ulong)bytes.Length);
                            break;
                    }
                }
                if (state != 0) Fail();
            }

            public void EndRow() { if (duckdb_appender_end_row(handle) != 0) Fail(); }

            void Fail() { throw Interop.Error(5, "DuckDB: " + Text(duckdb_appender_error(handle))); }

            public void Dispose()
            {
                if (handle == IntPtr.Zero) return;
                int state = duckdb_appender_close(handle);
                string error = state != 0 ? Text(duckdb_appender_error(handle)) : null;
                duckdb_appender_destroy(ref handle);
                handle = IntPtr.Zero;
                if (error != null) throw Interop.Error(5, "DuckDB: " + error);
            }
        }

        // ------------------------------------------------------------------
        // SQL
        // ------------------------------------------------------------------

        public static string Quote(string identifier) { return "\"" + identifier.Replace("\"", "\"\"") + "\""; }

        // Valor do VBA como literal SQL
        public static string Literal(object value)
        {
            value = Interop.Unwrap(value);
            if (value == null || value is DBNull || Interop.IsMissing(value)) return "NULL";
            string s = value as string;
            if (s != null) return "'" + s.Replace("'", "''") + "'";
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            if (value is DateTime)
            {
                DateTime d = (DateTime)value;
                return d.TimeOfDay == TimeSpan.Zero ? "DATE '" + d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + "'"
                    : "TIMESTAMP '" + d.ToString("yyyy-MM-dd HH:mm:ss.ffffff", CultureInfo.InvariantCulture) + "'";
            }
            if (value is double) return ((double)value).ToString("R", CultureInfo.InvariantCulture) + "::DOUBLE";
            if (value is float) return ((float)value).ToString("R", CultureInfo.InvariantCulture) + "::DOUBLE";
            if (value is decimal) return ((decimal)value).ToString(CultureInfo.InvariantCulture) + "::DOUBLE";
            if (PyOrder.IsNumber(value)) return Convert.ToString(value, CultureInfo.InvariantCulture);
            return "'" + Interop.Text(value).Replace("'", "''") + "'";
        }

        static byte[] Utf8(string s)
        {
            if (s == null) return null;
            byte[] bytes = new byte[Encoding.UTF8.GetByteCount(s) + 1];
            Encoding.UTF8.GetBytes(s, 0, s.Length, bytes, 0);
            return bytes;
        }

        static string Text(IntPtr utf8)
        {
            if (utf8 == IntPtr.Zero) return "";
            int length = 0;
            while (Marshal.ReadByte(utf8, length) != 0) length++;
            byte[] bytes = new byte[length];
            Marshal.Copy(utf8, bytes, 0, length);
            return Encoding.UTF8.GetString(bytes);
        }
    }
}
