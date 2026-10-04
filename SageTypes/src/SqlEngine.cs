using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace SageTypes
{
    // SqlEngine: conexão a um banco de dados, como o Engine do SQLAlchemy usado pelo pandas
    // (pd.read_sql(consulta, engine), df.to_sql("tabela", engine)).
    //
    //   Dim db As New Sage.SqlEngine
    //   db.Connect "mssql://servidor/Vendas"                      ' ou postgresql://, mysql://, arquivo .db/.accdb...
    //   Set df = db.ReadSql("SELECT * FROM Pedidos WHERE Ano = ?", Array(2024))
    //   db.Execute "UPDATE Pedidos SET Status = 'OK' WHERE Id = ?", Array(10)
    //   df.ToSql "PedidosCopia", db, "replace"
    //
    // PostgreSQL, MySQL, SQLite e arquivos do DuckDB passam pelo próprio DuckDB do DataFrame
    // (extensões que vêm no pacote): os dados chegam direto no formato do DataFrame. SQL Server
    // usa o cliente do .NET (sem driver ODBC; gravação por bulk copy); Access, ODBC e OLE DB, os
    // drivers do Windows. As consultas vão no SQL do próprio banco; parâmetros com "?".
    //
    // Interface dual: acrescente membros só no fim, com o próximo DispId (ver StringS).
    [ComVisible(true), Guid("9FEA3852-E1D7-4FB1-B27B-84749D9C2DB8"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _SqlEngine
    {
        [DispId(1)] SqlEngine Connect(string ConnectionString);
        [DispId(2)] DataFrame ReadSql(string Query, [Optional] object Params);
        [DispId(3)] long Execute(string Sql, [Optional] object Params);
        [DispId(4)] ListS Tables();
        [DispId(5)] void Close();
        [DispId(6)] StringS Kind { get; }
        [DispId(7)] bool Connected { get; }
        [DispId(8)] string ToString();
    }

    [ComVisible(true), Guid("6795C777-D195-4CBE-9E32-3DADA65AE733"), ProgId("Sage.SqlEngine")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_SqlEngine))]
    public sealed class SqlEngine : _SqlEngine
    {
        enum Route { None, Remote, Attached, Ado }
        enum Dialect { Generic, SqlServer, Access }

        static int counter;

        Route route = Route.None;
        string kind = "";
        string alias;                    // banco anexado no DuckDB (Remote e Attached)
        Engine.Connection attached;      // conexão com USE alias (SQLite, arquivo do DuckDB)
        DbConnection ado;
        Dialect dialect;

        public SqlEngine() { }

        public StringS Kind { get { return (StringS)Interop.Wrap(kind); } }
        public bool Connected { get { return route != Route.None; } }
        public override string ToString() { return route == Route.None ? "<SqlEngine (desconectado)>" : "<SqlEngine " + kind + ">"; }

        // ------------------------------------------------------------------
        // Conectar
        // ------------------------------------------------------------------

        // Formatos aceitos (detectados pelo texto):
        //   postgresql://usuário:senha@servidor:5432/banco   ou   host=... dbname=... (libpq)
        //   mysql://usuário:senha@servidor:3306/banco
        //   C:\dados\base.db | .sqlite | .sqlite3   ou   sqlite:C:\dados\base.db
        //   C:\dados\base.duckdb                      ou   duckdb:C:\dados\base.duckdb
        //   mssql://servidor/banco | mssql://usuário:senha@servidor/banco   ou   Server=...;Database=...
        //   C:\dados\base.accdb | .mdb (Access)
        //   Driver={...};... ou DSN=... (ODBC)          Provider=...;... (OLE DB)
        public SqlEngine Connect(string ConnectionString)
        {
            Close();
            string text = (ConnectionString ?? "").Trim();
            string lower = text.ToLowerInvariant();
            try
            {
                if (lower.StartsWith("postgresql://") || lower.StartsWith("postgres://") || Regex.IsMatch(lower, @"(^|\s)dbname\s*="))
                    AttachRemote("postgres", "postgres_scanner", text);
                else if (lower.StartsWith("mysql://"))
                    AttachRemote("mysql", "mysql_scanner", MySqlKeyValues(text));
                else if (lower.StartsWith("sqlite:") || Regex.IsMatch(lower, @"\.(db|sqlite|sqlite3)$"))
                    AttachFile("sqlite", "sqlite_scanner", FilePath(text, "sqlite:"));
                else if (lower.StartsWith("duckdb:") || lower.EndsWith(".duckdb"))
                    AttachFile("duckdb", null, FilePath(text, "duckdb:"));
                else if (lower.StartsWith("mssql://") || lower.StartsWith("sqlserver://"))
                    OpenAdo("sqlserver", new SqlConnection(SqlServerUrl(text)), Dialect.SqlServer);
                else if (Regex.IsMatch(lower, @"\.(accdb|mdb)$"))
                    OpenAdo("access", AccessConnection(text), Dialect.Access);
                else if (Regex.IsMatch(lower, @"(^|;)\s*(driver|dsn)\s*="))
                    OpenAdo("odbc", new OdbcConnection(text), DialectOf(lower));
                else if (Regex.IsMatch(lower, @"(^|;)\s*provider\s*="))
                    OpenAdo(lower.Contains("microsoft.ace") || lower.Contains("microsoft.jet") ? "access" : "oledb", new OleDbConnection(text), DialectOf(lower));
                else if (Regex.IsMatch(lower, @"(^|;)\s*(server|data source)\s*="))
                    OpenAdo("sqlserver", new SqlConnection(text), Dialect.SqlServer);
                else
                    throw Interop.Error(5, "ValueError: conexão não reconhecida. Use, por exemplo, postgresql://..., mysql://..., " +
                        "mssql://servidor/banco, um arquivo .db/.duckdb/.accdb, Driver={...} (ODBC) ou Provider=... (OLE DB).");
            }
            catch (COMException ex) { Close(); throw Interop.Error(5, Mask(ex.Message, text)); }
            catch (Exception ex) { Close(); throw Interop.Error(5, "ConnectionError: " + Mask(ex.Message, text)); }
            return this;
        }

        // As mensagens de erro dos drivers repetem a conexão, com a senha: troca a senha por ***
        static string Mask(string message, string connection)
        {
            List<string> secrets = new List<string>();
            Match url = Regex.Match(connection, @"://[^:/@]*:([^@]*)@");
            if (url.Success) secrets.Add(url.Groups[1].Value);
            foreach (Match m in Regex.Matches(connection, @"(?:^|[;\s])(?:password|pwd|passwd)\s*=\s*([^;\s]*)", RegexOptions.IgnoreCase))
                secrets.Add(m.Groups[1].Value);
            foreach (string secret in secrets)
            {
                if (secret.Length == 0) continue;
                message = message.Replace(secret, "***");
                string unescaped = Uri.UnescapeDataString(secret);
                if (unescaped.Length > 0) message = message.Replace(unescaped, "***");
            }
            return message;
        }

        void AttachRemote(string name, string extension, string target)
        {
            Engine.LoadExtension(extension);
            alias = NewAlias();
            Engine.Run("ATTACH " + Engine.Literal(target) + " AS " + Engine.Quote(alias) + " (TYPE " + name + ")");
            kind = name;
            route = Route.Remote;
        }

        void AttachFile(string name, string extension, string path)
        {
            if (extension != null) Engine.LoadExtension(extension);
            alias = NewAlias();
            Engine.Run("ATTACH " + Engine.Literal(Path.GetFullPath(path)) + " AS " + Engine.Quote(alias) +
                (extension != null ? " (TYPE " + name + ")" : ""));
            attached = new Engine.Connection();
            attached.Run("USE " + Engine.Quote(alias));
            kind = name;
            route = Route.Attached;
        }

        void OpenAdo(string name, DbConnection connection, Dialect sqlDialect)
        {
            connection.Open();
            ado = connection;
            dialect = sqlDialect;
            kind = name;
            route = Route.Ado;
        }

        static string NewAlias()
        {
            counter++;
            return "sage_db_" + counter.ToString(CultureInfo.InvariantCulture);
        }

        static string FilePath(string text, string prefix)
        {
            string path = text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? text.Substring(prefix.Length) : text;
            if (path.StartsWith("///")) path = path.Substring(3);   // sqlite:///C:/... (SQLAlchemy)
            return path.Replace('/', '\\');
        }

        // mysql://usuário:senha@servidor:porta/banco -> host=... user=... (formato da extensão do DuckDB)
        static string MySqlKeyValues(string url)
        {
            Uri uri = new Uri(url);
            List<string> parts = new List<string> { "host=" + uri.Host };
            if (uri.Port > 0) parts.Add("port=" + uri.Port.ToString(CultureInfo.InvariantCulture));
            if (uri.UserInfo.Length > 0)
            {
                string[] user = uri.UserInfo.Split(new[] { ':' }, 2);
                parts.Add("user=" + Uri.UnescapeDataString(user[0]));
                if (user.Length > 1) parts.Add("password=" + Uri.UnescapeDataString(user[1]));
            }
            string database = uri.AbsolutePath.Trim('/');
            if (database.Length > 0) parts.Add("database=" + Uri.UnescapeDataString(database));
            return string.Join(" ", parts);
        }

        // mssql://[usuário:senha@]servidor[/banco]; servidor pode ter instância: (localdb)\MSSQLLocalDB
        static string SqlServerUrl(string url)
        {
            string rest = url.Substring(url.IndexOf("://", StringComparison.Ordinal) + 3);
            string user = null, password = null;
            int at = rest.LastIndexOf('@');
            if (at >= 0)
            {
                string[] u = rest.Substring(0, at).Split(new[] { ':' }, 2);
                user = Uri.UnescapeDataString(u[0]);
                if (u.Length > 1) password = Uri.UnescapeDataString(u[1]);
                rest = rest.Substring(at + 1);
            }
            string server = rest, database = null;
            int slash = rest.IndexOf('/');
            if (slash >= 0) { server = rest.Substring(0, slash); database = rest.Substring(slash + 1).Trim('/'); }
            SqlConnectionStringBuilder b = new SqlConnectionStringBuilder();
            b.DataSource = server;
            if (!string.IsNullOrEmpty(database)) b.InitialCatalog = database;
            if (user != null) { b.UserID = user; b.Password = password ?? ""; }
            else b.IntegratedSecurity = true;   // usuário do Windows
            b.TrustServerCertificate = true;
            return b.ConnectionString;
        }

        // O provedor do Office mais novo instalado (16.0, depois 12.0)
        static DbConnection AccessConnection(string path)
        {
            string full = Path.GetFullPath(FilePath(path, "access:"));
            Exception last = null;
            foreach (string provider in new[] { "Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0" })
            {
                OleDbConnection c = new OleDbConnection("Provider=" + provider + ";Data Source=" + full + ";");
                try { c.Open(); c.Close(); return c; }
                catch (Exception ex) { last = ex; c.Dispose(); }
            }
            throw new InvalidOperationException("Access: " + (last != null ? last.Message : "provedor ACE não encontrado") +
                " (instale o Microsoft Access Database Engine de 64 bits).");
        }

        static Dialect DialectOf(string lower)
        {
            if (lower.Contains("access") || lower.Contains("microsoft.ace") || lower.Contains("microsoft.jet")) return Dialect.Access;
            if (lower.Contains("sql server") || lower.Contains("sqloledb") || lower.Contains("msoledbsql") || lower.Contains("sqlncli")) return Dialect.SqlServer;
            return Dialect.Generic;
        }

        public void Close()
        {
            if (attached != null) { attached.Dispose(); attached = null; }
            if (alias != null)
            {
                try { Engine.Run("DETACH " + Engine.Quote(alias)); }
                catch (Exception) { }
                alias = null;
            }
            if (ado != null) { ado.Dispose(); ado = null; }
            route = Route.None;
            kind = "";
        }

        void Require()
        {
            if (route == Route.None) throw Interop.Error(91, "SqlEngine desconectado: chame Connect antes.");
        }

        // ------------------------------------------------------------------
        // Ler e executar
        // ------------------------------------------------------------------

        // pd.read_sql: consulta no SQL do próprio banco, ou só o nome de uma tabela
        public DataFrame ReadSql(string Query, object Params)
        {
            Require();
            string query = Query.Trim();
            if (Regex.IsMatch(query, @"^[\w\.\[\]""`]+$")) query = "SELECT * FROM " + query;
            List<object> values = ParamList(Params);
            switch (route)
            {
                case Route.Remote:
                    string sql = "SELECT * FROM " + kind + "_query(" + Engine.Literal(alias) + ", " + Engine.Literal(Inline(query, values, NativeLiteral)) + ")";
                    return DataFrame.FromTable(name => Engine.Run("CREATE TABLE " + Engine.Quote(name) + " AS " + sql));
                case Route.Attached:
                    string local = Inline(query, values, Engine.Literal);
                    return DataFrame.FromTable(name => attached.Run("CREATE TABLE " + Engine.Quote(Engine.MainCatalog) + ".main." + Engine.Quote(name) + " AS " + local));
                default:
                    return ReadAdo(query, values);
            }
        }

        // Comandos sem resultado (INSERT, UPDATE, DELETE, CREATE...): devolve as linhas afetadas
        // (-1 quando o banco não informa)
        public long Execute(string Sql, object Params)
        {
            Require();
            List<object> values = ParamList(Params);
            try
            {
                switch (route)
                {
                    case Route.Remote:
                        Engine.Run("CALL " + kind + "_execute(" + Engine.Literal(alias) + ", " + Engine.Literal(Inline(Sql, values, NativeLiteral)) + ")");
                        return -1;
                    case Route.Attached:
                        using (Engine.QueryResult r = attached.Query(Inline(Sql, values, Engine.Literal)))
                        {
                            // O DuckDB devolve uma coluna "Count" em INSERT/UPDATE/DELETE
                            if (r.Names.Length != 1 || !string.Equals(r.Names[0], "Count", StringComparison.OrdinalIgnoreCase)) return -1;
                            List<object[]> rows = r.ReadAll(1);
                            return rows.Count == 0 || rows[0][0] == null ? -1 : Convert.ToInt64(rows[0][0], CultureInfo.InvariantCulture);
                        }
                    default:
                        using (DbCommand cmd = Command(Sql, values)) return cmd.ExecuteNonQuery();
                }
            }
            catch (DbException ex) { throw Interop.Error(5, "DatabaseError: " + ex.Message); }
        }

        public ListS Tables()
        {
            Require();
            List<object> names = new List<object>();
            if (route == Route.Ado)
            {
                DataTable t = ado.GetSchema("Tables");
                foreach (DataRow row in t.Rows)
                {
                    string type = Column(row, "TABLE_TYPE");
                    if (!string.Equals(type, "TABLE", StringComparison.OrdinalIgnoreCase) && !string.Equals(type, "BASE TABLE", StringComparison.OrdinalIgnoreCase)) continue;
                    string schema = Column(row, "TABLE_SCHEMA");
                    string name = Column(row, "TABLE_NAME");
                    names.Add(string.IsNullOrEmpty(schema) || schema == "dbo" ? name : schema + "." + name);
                }
                names.Sort((a, b) => string.Compare((string)a, (string)b, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                using (Engine.QueryResult r = Engine.Query("SELECT schema_name, table_name FROM duckdb_tables() WHERE database_name = " +
                    Engine.Literal(alias) + " ORDER BY 1, 2"))
                    foreach (object[] row in r.ReadAll(long.MaxValue))
                    {
                        string schema = Interop.Text(row[0]);
                        names.Add(schema == "main" || schema == "public" ? Interop.Text(row[1]) : schema + "." + Interop.Text(row[1]));
                    }
            }
            return ListS.From(names);
        }

        static string Column(DataRow row, string name)
        {
            return row.Table.Columns.Contains(name) && row[name] != DBNull.Value ? Convert.ToString(row[name], CultureInfo.InvariantCulture) : "";
        }

        // ------------------------------------------------------------------
        // Parâmetros ("?")
        // ------------------------------------------------------------------

        // Array(...) / ListS com os valores, na ordem dos "?"; um valor só também vale
        static List<object> ParamList(object values)
        {
            if (Interop.IsMissing(values) || values == null) return new List<object>();
            object raw = Interop.Unwrap(values);
            if (raw is Array || values is ListS) return Interop.Items(values).Select(v => Interop.Unwrap(v)).ToList();
            if (values is DictionaryS) throw Interop.Error(13, "TypeError: parâmetros por posição: use Array(valor1, valor2...), na ordem dos \"?\".");
            return new List<object> { raw };
        }

        // Troca cada "?" fora de textos e nomes entre aspas pelo literal do valor
        static string Inline(string sql, List<object> values, Func<object, string> literal)
        {
            int next = 0;
            string result = Placeholders(sql, delegate(int i)
            {
                if (i >= values.Count) throw Interop.Error(5, "ValueError: a consulta tem mais \"?\" do que parâmetros (" + values.Count + ").");
                next = i + 1;
                return literal(values[i]);
            });
            if (next != values.Count) throw Interop.Error(5, "ValueError: " + values.Count + " parâmetros para " + next + " \"?\" na consulta.");
            return result;
        }

        static string Placeholders(string sql, Func<int, string> replace)
        {
            StringBuilder sb = new StringBuilder();
            char quote = '\0';
            int n = 0;
            foreach (char c in sql)
            {
                if (quote != '\0')
                {
                    sb.Append(c);
                    if (c == quote || (quote == '[' && c == ']')) quote = '\0';
                    continue;
                }
                if (c == '\'' || c == '"' || c == '`' || c == '[') { quote = c; sb.Append(c); continue; }
                if (c == '?') { sb.Append(replace(n++)); continue; }
                sb.Append(c);
            }
            return sb.ToString();
        }

        // Literal no SQL do PostgreSQL/MySQL (sem os "::DOUBLE" do DuckDB)
        string NativeLiteral(object value)
        {
            if (value == null || value is DBNull) return "NULL";
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            if (value is DateTime)
            {
                DateTime d = (DateTime)value;
                return "'" + d.ToString(d.TimeOfDay == TimeSpan.Zero ? "yyyy-MM-dd" : "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + "'";
            }
            if (PyOrder.IsNumber(value)) return Convert.ToString(value, CultureInfo.InvariantCulture);
            string s = Interop.Text(value).Replace("'", "''");
            if (kind == "mysql") s = s.Replace("\\", "\\\\"); // no MySQL, a barra invertida escapa
            return "'" + s + "'";
        }

        DbCommand Command(string sql, List<object> values)
        {
            DbCommand cmd = ado.CreateCommand();
            cmd.CommandTimeout = 0; // consultas longas de análise: sem limite
            // SqlClient usa parâmetros com nome (@p0); ODBC e OLE DB, "?"
            cmd.CommandText = dialect == Dialect.SqlServer && ado is SqlConnection ? Placeholders(sql, i => "@p" + i.ToString(CultureInfo.InvariantCulture)) : sql;
            for (int i = 0; i < values.Count; i++)
            {
                DbParameter p = cmd.CreateParameter();
                p.ParameterName = "@p" + i.ToString(CultureInfo.InvariantCulture);
                p.Value = values[i] ?? DBNull.Value;
                // O Access recusa datas como DBTimeStamp (o padrão do OLE DB para DateTime)
                OleDbParameter ole = p as OleDbParameter;
                if (ole != null && values[i] is DateTime) ole.OleDbType = OleDbType.Date;
                cmd.Parameters.Add(p);
            }
            return cmd;
        }

        // ------------------------------------------------------------------
        // Leitura pelo ADO.NET (SQL Server, Access, ODBC, OLE DB) para um DataFrame
        // ------------------------------------------------------------------

        DataFrame ReadAdo(string query, List<object> values)
        {
            try
            {
                using (DbCommand cmd = Command(query, values))
                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.FieldCount == 0) throw Interop.Error(5, "A consulta não devolveu linhas; para comandos, use Execute.");
                    int n = reader.FieldCount;
                    List<string> names = new List<string>();
                    string[] types = new string[n];
                    int[] kinds = new int[n];
                    for (int i = 0; i < n; i++)
                    {
                        string name = reader.GetName(i);
                        if (string.IsNullOrEmpty(name)) name = "Column" + (i + 1).ToString(CultureInfo.InvariantCulture);
                        string unique = name;
                        for (int k = 1; names.Contains(unique, StringComparer.OrdinalIgnoreCase); k++) unique = name + "." + k.ToString(CultureInfo.InvariantCulture);
                        names.Add(unique);
                        types[i] = DuckType(reader.GetFieldType(i), reader.GetDataTypeName(i), out kinds[i]);
                    }
                    return DataFrame.FromTable(delegate(string table)
                    {
                        Engine.Run("CREATE TABLE " + Engine.Quote(table) + " (" +
                            string.Join(", ", Enumerable.Range(0, n).Select(i => Engine.Quote(names[i]) + " " + types[i])) + ")");
                        object[] row = new object[n];
                        using (Engine.Appender appender = new Engine.Appender(table))
                            while (reader.Read())
                            {
                                reader.GetValues(row);
                                for (int i = 0; i < n; i++) appender.Add(Plain(row[i]), kinds[i]);
                                appender.EndRow();
                            }
                    });
                }
            }
            catch (DbException ex) { throw Interop.Error(5, "DatabaseError: " + ex.Message); }
        }

        static string DuckType(Type type, string dbType, out int kind)
        {
            if (type == typeof(bool)) { kind = Engine.TBoolean; return "BOOLEAN"; }
            if (type == typeof(byte) || type == typeof(sbyte) || type == typeof(short) || type == typeof(ushort) ||
                type == typeof(int) || type == typeof(uint) || type == typeof(long)) { kind = Engine.TBigInt; return "BIGINT"; }
            if (type == typeof(float) || type == typeof(double) || type == typeof(decimal) || type == typeof(ulong)) { kind = Engine.TDouble; return "DOUBLE"; }
            if (type == typeof(DateTime) || type == typeof(DateTimeOffset))
            {
                if (string.Equals(dbType, "date", StringComparison.OrdinalIgnoreCase)) { kind = Engine.TDate; return "DATE"; }
                kind = Engine.TTimestamp; return "TIMESTAMP";
            }
            kind = Engine.TVarchar;
            return "VARCHAR";
        }

        // Valores que o Appender não conhece viram texto
        static object Plain(object value)
        {
            if (value == null || value is DBNull) return null;
            if (value is DateTimeOffset) return ((DateTimeOffset)value).DateTime;
            byte[] bytes = value as byte[];
            if (bytes != null) return Convert.ToBase64String(bytes);
            if (value is Guid || value is TimeSpan) return value.ToString();
            return value;
        }

        // ------------------------------------------------------------------
        // Gravação (df.ToSql)
        // ------------------------------------------------------------------

        internal void WriteTable(DataFrame df, string name, string mode)
        {
            Require();
            if (string.IsNullOrEmpty(name)) throw Interop.Error(5, "ValueError: informe o nome da tabela.");
            string[] parts = name.Split('.');
            string table = parts[parts.Length - 1], schema = parts.Length > 1 ? parts[parts.Length - 2] : null;
            try
            {
                if (route == Route.Ado) WriteAdo(df, name, table, schema, mode);
                else WriteDuck(df, parts, table, schema, mode);
            }
            catch (DbException ex) { throw Interop.Error(5, "DatabaseError: " + ex.Message); }
        }

        void WriteDuck(DataFrame df, string[] parts, string table, string schema, string mode)
        {
            string target = Engine.Quote(alias) + "." + string.Join(".", parts.Select(Engine.Quote));
            bool exists = Convert.ToInt64(Engine.Scalar("SELECT count(*) FROM duckdb_tables() WHERE database_name = " + Engine.Literal(alias) +
                " AND lower(table_name) = lower(" + Engine.Literal(table) + ")" +
                (schema != null ? " AND lower(schema_name) = lower(" + Engine.Literal(schema) + ")" : "")), CultureInfo.InvariantCulture) > 0;
            string source = df.Source;
            if (exists)
            {
                if (mode == "fail") throw Interop.Error(5, "ValueError: a tabela '" + string.Join(".", parts) + "' já existe (use IfExists:=\"replace\" ou \"append\").");
                if (mode == "append") { Engine.Run("INSERT INTO " + target + " BY NAME SELECT * FROM (" + source + ")"); return; }
                Engine.Run("DROP TABLE " + target);
            }
            Engine.Run("CREATE TABLE " + target + " AS SELECT * FROM (" + source + ")");
        }

        void WriteAdo(DataFrame df, string fullName, string table, string schema, string mode)
        {
            List<string> columns = df.ColumnNames;
            List<string> types = df.ColumnTypes;
            string target = (schema != null ? QuoteName(schema) + "." : "") + QuoteName(table);
            bool exists = TableExists(table, schema);
            if (exists && mode == "fail") throw Interop.Error(5, "ValueError: a tabela '" + fullName + "' já existe (use IfExists:=\"replace\" ou \"append\").");
            if (exists && mode == "replace") { Exec("DROP TABLE " + target); exists = false; }
            if (!exists) Exec("CREATE TABLE " + target + " (" + string.Join(", ", columns.Select((c, i) => QuoteName(c) + " " + SqlType(df, c, types[i]))) + ")");

            string readable = df.ReadableSource;
            if (ado is SqlConnection) { BulkCopy((SqlConnection)ado, target, columns, readable); return; }

            // Outros bancos: INSERT com parâmetros, numa transação
            using (DbTransaction tx = ado.BeginTransaction())
            using (DbCommand cmd = ado.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "INSERT INTO " + target + " (" + string.Join(", ", columns.Select(QuoteName)) + ") VALUES (" +
                    string.Join(", ", columns.Select(c => "?")) + ")";
                DbParameter[] ps = new DbParameter[columns.Count];
                for (int i = 0; i < columns.Count; i++)
                {
                    ps[i] = cmd.CreateParameter();
                    OleDbParameter ole = ps[i] as OleDbParameter;
                    if (ole != null && (types[i] == "DATE" || types[i].StartsWith("TIMESTAMP"))) ole.OleDbType = OleDbType.Date;
                    cmd.Parameters.Add(ps[i]);
                }
                using (Engine.QueryResult r = Engine.Query(readable))
                {
                    int rows;
                    object[][] chunk;
                    while ((chunk = r.NextChunk(out rows)) != null)
                        for (int row = 0; row < rows; row++)
                        {
                            for (int i = 0; i < ps.Length; i++) ps[i].Value = chunk[i][row] ?? DBNull.Value;
                            cmd.ExecuteNonQuery();
                        }
                }
                tx.Commit();
            }
        }

        // SQL Server: bulk copy em lotes de 50 mil linhas
        static void BulkCopy(SqlConnection connection, string target, List<string> columns, string readable)
        {
            using (SqlBulkCopy bulk = new SqlBulkCopy(connection))
            using (Engine.QueryResult r = Engine.Query(readable))
            {
                bulk.DestinationTableName = target;
                bulk.BulkCopyTimeout = 0;
                foreach (string c in columns) bulk.ColumnMappings.Add(c, c);
                DataTable batch = new DataTable();
                foreach (string c in columns) batch.Columns.Add(c, typeof(object));
                int rows;
                object[][] chunk;
                while ((chunk = r.NextChunk(out rows)) != null)
                {
                    for (int row = 0; row < rows; row++)
                    {
                        object[] values = new object[columns.Count];
                        for (int i = 0; i < values.Length; i++) values[i] = chunk[i][row] ?? DBNull.Value;
                        batch.Rows.Add(values);
                    }
                    if (batch.Rows.Count >= 50000) { bulk.WriteToServer(batch); batch.Clear(); }
                }
                if (batch.Rows.Count > 0) bulk.WriteToServer(batch);
            }
        }

        bool TableExists(string table, string schema)
        {
            if (ado is SqlConnection)
            {
                using (DbCommand cmd = ado.CreateCommand())
                {
                    cmd.CommandText = "SELECT OBJECT_ID(@n, 'U')";
                    DbParameter p = cmd.CreateParameter();
                    p.ParameterName = "@n";
                    p.Value = (schema != null ? schema + "." : "") + table;
                    cmd.Parameters.Add(p);
                    object id = cmd.ExecuteScalar();
                    return id != null && id != DBNull.Value;
                }
            }
            foreach (DataRow row in ado.GetSchema("Tables").Rows)
                if (string.Equals(Column(row, "TABLE_NAME"), table, StringComparison.OrdinalIgnoreCase) &&
                    (schema == null || string.Equals(Column(row, "TABLE_SCHEMA"), schema, StringComparison.OrdinalIgnoreCase)))
                    return true;
            return false;
        }

        void Exec(string sql)
        {
            using (DbCommand cmd = ado.CreateCommand())
            {
                cmd.CommandText = sql;
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();
            }
        }

        string QuoteName(string name)
        {
            if (dialect == Dialect.Generic) return "\"" + name.Replace("\"", "\"\"") + "\"";
            return "[" + name.Replace("]", "]]") + "]";
        }

        // Tipo da coluna no banco de destino, a partir do tipo no DuckDB
        string SqlType(DataFrame df, string column, string duckType)
        {
            string t = duckType.ToUpperInvariant();
            bool integer = t.EndsWith("INT") || t == "INTEGER";
            bool number = integer || t == "DOUBLE" || t == "FLOAT" || t.StartsWith("DECIMAL") || t == "REAL";
            string c = Engine.Quote(column);
            switch (dialect)
            {
                case Dialect.SqlServer:
                    if (t == "BOOLEAN") return "BIT";
                    if (integer) return "BIGINT";
                    if (number) return "FLOAT";
                    if (t == "DATE") return "DATE";
                    if (t.StartsWith("TIMESTAMP")) return "DATETIME2";
                    return "NVARCHAR(MAX)";
                case Dialect.Access:
                    if (t == "BOOLEAN") return "YESNO";
                    if (integer) return FitsInt32(df, c) ? "LONG" : "DOUBLE";
                    if (number) return "DOUBLE";
                    if (t == "DATE" || t.StartsWith("TIMESTAMP")) return "DATETIME";
                    return MaxLength(df, c) <= 255 ? "TEXT(255)" : "LONGTEXT";
                default:
                    if (t == "BOOLEAN") return "SMALLINT";
                    if (integer) return "BIGINT";
                    if (number) return "FLOAT";
                    if (t == "DATE") return "DATE";
                    if (t.StartsWith("TIMESTAMP")) return "TIMESTAMP";
                    long length = MaxLength(df, c);
                    return "VARCHAR(" + Math.Max(1, Math.Min(4000, length)).ToString(CultureInfo.InvariantCulture) + ")";
            }
        }

        static bool FitsInt32(DataFrame df, string column)
        {
            object min = Engine.Scalar("SELECT min(" + column + ") FROM (" + df.Source + ")");
            object max = Engine.Scalar("SELECT max(" + column + ") FROM (" + df.Source + ")");
            return (min == null || Convert.ToInt64(min, CultureInfo.InvariantCulture) >= int.MinValue) &&
                   (max == null || Convert.ToInt64(max, CultureInfo.InvariantCulture) <= int.MaxValue);
        }

        static long MaxLength(DataFrame df, string column)
        {
            object v = Engine.Scalar("SELECT max(length(CAST(" + column + " AS VARCHAR))) FROM (" + df.Source + ")");
            return v == null ? 1 : Convert.ToInt64(v, CultureInfo.InvariantCulture);
        }
    }
}
