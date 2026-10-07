using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace SageTypes
{
    // DataFrame: tabela como a do pandas, executada pelo DuckDB (banco analítico
    // embutido, colunar e multithread), para dezenas de milhões de linhas.
    //
    //   Dim df As New Sage.DataFrame
    //   df.ReadCsv "C:\dados\vendas.csv"
    //   df.Eval "Total = Qtd * Preco"
    //   Debug.Print df.Query("Cidade = ""SP"" And Total > 100").GroupBy("Produto", "Soma = sum(Total)").Head(10)
    //
    // Cada DataFrame é uma consulta "preguiçosa": Query, Eval, GroupBy, Merge... só montam
    // o SQL, e o DuckDB executa quando um resultado é pedido (Count, Head, ToRange, ToCsv,
    // Show, Debug.Print). Arquivos são carregados numa tabela interna uma vez só.
    //
    // Expressões (Query, Eval, GroupBy...) usam o SQL do DuckDB, com dois ajustes de VBA:
    // texto entre aspas duplas é texto ("SP"), e nomes com espaço vão entre colchetes
    // ([Valor Total]). =, <>, And, Or, Not, Like (com %), Is Null funcionam.
    //
    // Cursor, como no Sage.xlam: MoveFirst/EOF/MoveNext e df("Coluna") lendo e gravando a
    // linha atual (as linhas são lidas em blocos e as alterações gravadas em lote). Fora
    // de uma linha (antes do MoveFirst ou depois do EOF), df("Coluna") é a coluna inteira:
    // lê como ListS e grava um valor (ou um array/ListS do mesmo tamanho) em todas as linhas.
    //
    // Interface dual: acrescente membros só no fim, com o próximo DispId (ver StringS).
    // [PropertyGet("X")]/[PropertyLet("X")] viram a propriedade X com parâmetros na
    // geração do .tlb (install.ps1).
    [ComVisible(true), Guid("B5E1D8A3-6C47-4F29-9A0E-3D7C2B8F1E56"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _DataFrame
    {
        [DispId(0), IndexerName("Value")] object this[[Optional] object Column] { get; set; }

        // Carregar
        [DispId(1)] DataFrame ReadCsv(string Path, [Optional] object Header, [Optional] object Delimiter, [Optional] object DecimalSeparator, [Optional] object Encoding);
        [DispId(2)] DataFrame ReadParquet(string Path);
        [DispId(3)] DataFrame ReadExcel(string Path, [Optional] object Sheet, [Optional] object Header);
        [DispId(4)] DataFrame FromRange(object Range, [Optional] object Header);
        [DispId(5)] DataFrame FromArray(object Data, [Optional] object Header);
        [DispId(6)] DataFrame FromRecords(object Records);
        [DispId(7)] DataFrame Sql(string Query, [Optional] object Df1, [Optional] object Df2, [Optional] object Df3, [Optional] object Df4);

        // Salvar e mostrar
        [DispId(10)] void ToCsv(string Path, [Optional] object Delimiter, [Optional] object Header);
        [DispId(11)] void ToParquet(string Path);
        [DispId(12)] void ToExcel(string Path, [Optional] object SheetName);
        [DispId(13)] void ToRange(object Range, [Optional] object Header);
        [DispId(14)] object ToArray([Optional] object Header);
        [DispId(15)] string ToString();
        [DispId(16)] void Show([Optional] object Title);

        // Informações
        [DispId(20)] long Count { get; }
        [DispId(21)] ListS Columns { get; }
        [DispId(22)] DictionaryS DTypes { get; }
        [DispId(23)] ListS Shape { get; }
        [DispId(24)] string Info();
        [DispId(25)] DataFrame Describe();
        [DispId(26)] DataFrame Head([Optional] object Rows);
        [DispId(27)] DataFrame Tail([Optional] object Rows);

        // Selecionar e filtrar
        [DispId(30)] DataFrame Query(string Condition);
        [DispId(31)] DataFrame Select(object Columns);
        [DispId(32)] DataFrame Drop(object Columns);
        [DispId(33)] DataFrame Slice([Optional] object Start, [Optional] object Stop);
        [DispId(34)] DataFrame Sample(long Rows);
        [DispId(35)] ListS Col(object Column);
        [DispId(36)] ListS Unique(object Column);

        // Transformar
        [DispId(40)] object Eval(string Expression);
        [DispId(41)] DataFrame Rename(object Mapping);
        [DispId(42)] DataFrame SortValues(object By, [Optional] object Ascending);
        [DispId(43)] DataFrame DropDuplicates([Optional] object Subset);
        [DispId(44)] DataFrame DropNA([Optional] object Subset);
        [DispId(45)] DataFrame FillNA(object Value, [Optional] object Columns);
        [DispId(46)] DataFrame Copy();

        // Combinar e resumir
        [DispId(50)] DataFrame GroupBy(object Keys, [Optional] object Aggregations);
        [DispId(51)] DataFrame Pivot(object Index, string Columns, string Values, [Optional] object Function);
        [DispId(52)] DataFrame Merge(DataFrame Other, [Optional] object On, [Optional] object How, [Optional] object LeftOn, [Optional] object RightOn);
        [DispId(53)] DataFrame Concat(DataFrame Other);
        [DispId(54)] DataFrame ValueCounts(object Column);
        [DispId(55)] object Sum(object Column);
        [DispId(56)] object Mean(object Column);
        [DispId(57)] object Min(object Column);
        [DispId(58)] object Max(object Column);
        [DispId(59)] object Median(object Column);
        [DispId(60)] object Std(object Column);
        [DispId(61)] long NUnique(object Column);

        // Cursor
        [DispId(70)] void MoveFirst();
        [DispId(71)] void MoveNext();
        [DispId(72)] void MovePrevious();
        [DispId(73)] void MoveLast();
        [DispId(74)] void Move(long Rows);
        [DispId(75)] bool EOF { get; }
        [DispId(76)] bool BOF { get; }
        [DispId(77)] long Index { get; set; }

        // At(linha, coluna) e Loc(linha): propriedades com parâmetros
        [DispId(80), PropertyGet("At")] object GetAt(long Row, object Column);
        [DispId(81), PropertyGet("Loc")] DictionaryS GetLoc(long Row);
        [DispId(-4)] IEnumerator GetEnumerator();

        // Viram Property Let na geração do .tlb (install.ps1). Ficam sempre por último.
        [DispId(1000), PropertyLet("Value")] void LetValue([Optional] object Column, object Value);
        [DispId(1001), PropertyLet("At")] void LetAt(long Row, object Column, object Value);
        [DispId(1002), PropertyLet("Loc")] void LetLoc(long Row, object Values);

        // Banco de dados (df.to_sql do pandas): IfExists = "fail" (padrão), "replace" ou "append"
        [DispId(82)] void ToSql(string Name, SqlEngine Engine, [Optional] object IfExists);
    }

    [ComVisible(true), Guid("2F9C6E14-8B3A-4D71-A5E2-7C0D9B4F3A68"), ProgId("Sage.DataFrame")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_DataFrame))]
    public sealed class DataFrame : _DataFrame, IEnumerable
    {
        // Tabela interna do DuckDB; apagada quando nenhum DataFrame a usa mais
        sealed class Table
        {
            public readonly string Name = Engine.NewTable();
            public bool Shared;   // outro DataFrame foi derivado dela: alterações criam uma cópia
            ~Table() { Engine.ScheduleDrop(Name); }
        }

        sealed class Column
        {
            public string Name, Type;
        }

        string sql;                                    // consulta do DataFrame; null = vazio
        List<Table> deps = new List<Table>();          // tabelas usadas pela consulta
        Table own;                                     // tabela própria (cursor e alterações em lote)
        List<Column> schema;
        string schemaFor;
        long count = -1;
        string countFor;

        // Cursor
        const int BlockSize = 65536;
        long index = -1;
        bool cursor;                                   // MoveFirst/Move... já chamados
        long blockStart = -1;
        int blockRows;
        object[][] block;                              // por coluna
        readonly Dictionary<string, Dictionary<long, object>> pending =
            new Dictionary<string, Dictionary<long, object>>(StringComparer.OrdinalIgnoreCase);
        long pendingBlock = -1;                        // bloco das alterações pendentes

        public DataFrame() { }

        DataFrame(string sql, IEnumerable<Table> deps)
        {
            this.sql = sql;
            this.deps = deps.Distinct().ToList();
        }

        // ------------------------------------------------------------------
        // Valor (membro padrão)
        // ------------------------------------------------------------------

        [IndexerName("Value")]
        public object this[object Column]
        {
            get
            {
                if (Interop.IsMissing(Column)) return ToString();     // Debug.Print df
                if (InRow) return Interop.Wrap(Cell(index, Resolve(Column)));
                return Col(Column);
            }
            set { Assign(Column, value); }
        }

        public void LetValue(object Column, object Value) { Assign(Column, Value); }

        void Assign(object column, object value)
        {
            if (Interop.IsMissing(column)) { Load(value, true); return; } // df = array
            string name = Interop.Text(column).Trim();
            if (InRow) { SetCell(index, name, value); return; }
            SetColumn(name, value);
        }

        bool InRow { get { return cursor && index >= 0 && index < Count; } }

        // ------------------------------------------------------------------
        // Carregar
        // ------------------------------------------------------------------

        // Os Read aceitam caminho relativo: parte da pasta da pasta de trabalho ativa (Excel.ResolvePath)
        public DataFrame ReadCsv(string Path, object Header, object Delimiter, object DecimalSeparator, object Encoding)
        {
            Path = Excel.ResolvePath(Path);
            List<string> options = new List<string> { Engine.Literal(Path), "header = " + (Flag(Header, true) ? "true" : "false") };
            if (!Interop.IsMissing(Delimiter)) options.Add("delim = " + Engine.Literal(Interop.Text(Delimiter)));
            if (!Interop.IsMissing(DecimalSeparator)) options.Add("decimal_separator = " + Engine.Literal(Interop.Text(DecimalSeparator)));
            if (!Interop.IsMissing(Encoding)) options.Add("encoding = " + Engine.Literal(Interop.Text(Encoding)));
            return Materialize("SELECT * FROM read_csv(" + string.Join(", ", options) + ")");
        }

        public DataFrame ReadParquet(string Path)
        {
            return Materialize("SELECT * FROM read_parquet(" + Engine.Literal(Excel.ResolvePath(Path)) + ")");
        }

        // Pelo Excel: abre o arquivo só para leitura, lê a área usada da aba e fecha
        public DataFrame ReadExcel(string Path, object Sheet, object Header)
        {
            Path = Excel.ResolvePath(Path); // antes de abrir: abrir muda a pasta de trabalho ativa
            dynamic app = Excel.Application(null);
            dynamic book = app.Workbooks.Open(Path, 0, true);
            try
            {
                dynamic sheet = Interop.IsMissing(Sheet) ? book.Worksheets[1] : book.Worksheets[Interop.Unwrap(Sheet)];
                object values = sheet.UsedRange.Value;
                return FromArray(values is Array ? values : new object[,] { { values } }, Header);
            }
            finally { book.Close(false); }
        }

        public DataFrame FromRange(object Range, object Header)
        {
            dynamic range = Range;
            object values = range.Value;
            return FromArray(values is Array ? values : new object[,] { { values } }, Header);
        }

        // Array 2D (Range.Value, ou montado no VBA); a primeira linha é o cabeçalho
        public DataFrame FromArray(object Data, object Header)
        {
            Load(Data, Flag(Header, true));
            return this;
        }

        // Lista de dicionários (DictionaryS ou Scripting.Dictionary), um por linha
        public DataFrame FromRecords(object Records)
        {
            List<List<KeyValuePair<object, object>>> rows = Interop.Items(Records).Select(r => Interop.PairsOf(r).ToList()).ToList();
            List<string> names = new List<string>();
            foreach (var row in rows)
                foreach (var pair in row)
                {
                    string name = Interop.Text(pair.Key);
                    if (!names.Contains(name, StringComparer.OrdinalIgnoreCase)) names.Add(name);
                }
            object[,] grid = new object[rows.Count, names.Count];
            for (int r = 0; r < rows.Count; r++)
                foreach (var pair in rows[r])
                    grid[r, names.FindIndex(n => string.Equals(n, Interop.Text(pair.Key), StringComparison.OrdinalIgnoreCase))] = pair.Value;
            Build(names, grid, 0);
            return this;
        }

        // SQL do DuckDB; {0} é este DataFrame e {1}..{4} os demais
        public DataFrame Sql(string Query, object Df1, object Df2, object Df3, object Df4)
        {
            List<DataFrame> frames = new List<DataFrame> { this };
            foreach (object o in new[] { Df1, Df2, Df3, Df4 })
                if (!Interop.IsMissing(o) && o is DataFrame) frames.Add((DataFrame)o);
            string text = Query;
            List<Table> used = new List<Table>();
            for (int i = 0; i < frames.Count; i++)
            {
                string token = "{" + i.ToString(CultureInfo.InvariantCulture) + "}";
                if (text.IndexOf(token, StringComparison.Ordinal) < 0) continue;
                if (frames[i].sql == null) throw Empty();
                frames[i].Sync();
                text = text.Replace(token, "(" + frames[i].sql + ")");
                used.AddRange(frames[i].Share());
            }
            return new DataFrame("SELECT * FROM (" + text + ")", used);
        }

        // Array 2D (ou ListS de linhas) -> tabela, com o tipo de cada coluna deduzido
        void Load(object data, bool header)
        {
            object[,] grid = Grid(data);
            int rows = grid.GetLength(0), cols = grid.GetLength(1);
            List<string> names = new List<string>();
            for (int c = 0; c < cols; c++)
            {
                string name = header && rows > 0 ? Interop.Text(grid[0, c]).Trim() : "";
                if (name.Length == 0) name = "Column" + (c + 1).ToString(CultureInfo.InvariantCulture);
                string unique = name;
                for (int n = 1; names.Contains(unique, StringComparer.OrdinalIgnoreCase); n++) unique = name + "." + n.ToString(CultureInfo.InvariantCulture);
                names.Add(unique);
            }
            Build(names, grid, header && rows > 0 ? 1 : 0);
        }

        // Valores de erro do Excel (#N/D, #VALOR!...): o Range.Value os entrega como Int32
        // (números de célula vêm como Double), e eles viram nulos
        static readonly HashSet<int> ExcelErrors = new HashSet<int>
            { -2146826281, -2146826246, -2146826259, -2146826288, -2146826252, -2146826265, -2146826273 };

        static object Clean(object value)
        {
            return value is int && ExcelErrors.Contains((int)value) ? null : value;
        }

        // Base 0, 2 dimensões; aceita array 1D (uma coluna) e ListS de linhas
        static object[,] Grid(object data)
        {
            Array array = Interop.Unwrap(data) as Array;
            ListS list = data as ListS;
            if (list != null) array = (Array)list.ToArray();
            if (array == null) throw Interop.Error(13, "Esperado um array 2D, um Range.Value ou uma ListS de linhas.");
            if (array.Rank == 2)
            {
                int r0 = array.GetLowerBound(0), c0 = array.GetLowerBound(1);
                object[,] grid = new object[array.GetLength(0), array.GetLength(1)];
                for (int r = 0; r < grid.GetLength(0); r++)
                    for (int c = 0; c < grid.GetLength(1); c++) grid[r, c] = Clean(array.GetValue(r0 + r, c0 + c));
                return grid;
            }
            // 1D: array de linhas (arrays/ListS) ou uma coluna só
            List<object[]> rows = array.Cast<object>().Select(r => (r is Array || r is ListS) ? Interop.Items(r).ToArray() : new[] { r }).ToList();
            int width = rows.Count == 0 ? 0 : rows.Max(r => r.Length);
            object[,] result = new object[rows.Count, width];
            for (int r = 0; r < rows.Count; r++)
                for (int c = 0; c < rows[r].Length; c++) result[r, c] = rows[r][c];
            return result;
        }

        void Build(List<string> names, object[,] grid, int firstRow)
        {
            int rows = grid.GetLength(0), cols = names.Count;
            string[] types = new string[cols];
            for (int c = 0; c < cols; c++)
            {
                string type = null;
                bool whole = true; // o Range.Value entrega todo número como Double
                for (int r = firstRow; r < rows; r++)
                {
                    type = Widen(type, TypeOf(grid[r, c]));
                    object v = grid[r, c];
                    if (v is double && (Math.Floor((double)v) != (double)v || Math.Abs((double)v) > 9e15)) whole = false;
                }
                // Como o pandas ao ler o Excel: números sem casas decimais viram inteiros
                if (type == "DOUBLE" && whole && !Enumerable.Range(firstRow, rows - firstRow).Any(r => grid[r, c] is float || grid[r, c] is decimal))
                    type = "BIGINT";
                types[c] = type ?? "VARCHAR";
            }
            Table table = new Table();
            Engine.Run("CREATE TABLE " + Engine.Quote(table.Name) + " (" +
                string.Join(", ", Enumerable.Range(0, cols).Select(c => Engine.Quote(names[c]) + " " + types[c])) + ")");
            using (Engine.Appender appender = new Engine.Appender(table.Name))
            {
                int[] kinds = types.Select(Kind).ToArray();
                for (int r = firstRow; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++) appender.Add(Cell(grid[r, c], kinds[c]), kinds[c]);
                    appender.EndRow();
                }
            }
            Reset("SELECT * FROM " + Engine.Quote(table.Name), new List<Table> { table }, table);
        }

        // Valor pronto para a coluna: texto vazio numa coluna que não é de texto vira nulo
        static object Cell(object value, int kind)
        {
            value = Interop.Unwrap(value);
            if (value is string && ((string)value).Length == 0 && kind != Engine.TVarchar) return null;
            return value;
        }

        // ------------------------------------------------------------------
        // Salvar e mostrar
        // ------------------------------------------------------------------

        // Separador padrão: o de listas do Windows (";" em português), como o Excel espera
        public void ToCsv(string Path, object Delimiter, object Header)
        {
            string delimiter = Interop.IsMissing(Delimiter) ? CultureInfo.CurrentCulture.TextInfo.ListSeparator : Interop.Text(Delimiter);
            Engine.Run("COPY (" + Need() + ") TO " + Engine.Literal(Path) + " (FORMAT csv, DELIMITER " + Engine.Literal(delimiter) +
                ", HEADER " + (Flag(Header, true) ? "true" : "false") + ")");
        }

        public void ToParquet(string Path)
        {
            Engine.Run("COPY (" + Need() + ") TO " + Engine.Literal(Path) + " (FORMAT parquet)");
        }

        // Pelo Excel; o formato vem da extensão (.xlsx, .xlsm, .xlsb, .xls)
        public void ToExcel(string Path, object SheetName)
        {
            if (Count > 1048575) throw Interop.Error(5, "Uma planilha do Excel comporta até 1.048.575 linhas com o cabeçalho; use ToCsv ou ToParquet.");
            dynamic app = Excel.Application(null);
            bool alerts = app.DisplayAlerts;
            dynamic book = app.Workbooks.Add();
            try
            {
                dynamic sheet = book.Worksheets[1];
                if (!Interop.IsMissing(SheetName)) sheet.Name = Interop.Text(SheetName);
                ToRange(sheet.Range["A1"], true);
                string ext = System.IO.Path.GetExtension(Path).ToLowerInvariant();
                int format = ext == ".xls" ? 56 : ext == ".xlsb" ? 50 : ext == ".xlsm" ? 52 : 51;
                app.DisplayAlerts = false;
                book.SaveAs(Path, format);
            }
            finally
            {
                book.Close(false);
                app.DisplayAlerts = alerts;
            }
        }

        // Grava a partir da célula indicada, em blocos (planilhas grandes)
        public void ToRange(object Range, object Header)
        {
            dynamic start = ((dynamic)Range).Cells[1, 1];
            List<Column> cols = Schema();
            if (Count + (Flag(Header, true) ? 1 : 0) > 1048576) throw Interop.Error(5, "Mais linhas do que uma planilha comporta; use ToCsv ou ToParquet.");
            int offset = 0;
            if (Flag(Header, true))
            {
                object[,] head = new object[1, cols.Count];
                for (int c = 0; c < cols.Count; c++) head[0, c] = cols[c].Name;
                start.Resize[1, cols.Count].Value = head;
                offset = 1;
            }
            using (Engine.QueryResult r = Engine.Query(Readable()))
            {
                const int Batch = 50000;
                List<object[]> rows;
                while ((rows = r.ReadAll(Batch)).Count > 0)
                {
                    object[,] grid = new object[rows.Count, cols.Count];
                    for (int i = 0; i < rows.Count; i++)
                        for (int c = 0; c < cols.Count; c++) grid[i, c] = ForExcel(rows[i][c]);
                    start.Offset[offset, 0].Resize[rows.Count, cols.Count].Value = grid;
                    offset += rows.Count;
                    if (rows.Count < Batch) break;
                }
            }
        }

        static object ForExcel(object value)
        {
            if (value is long || value is decimal) return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            return value;
        }

        // Array 2D do VBA (base 0), opcionalmente com o cabeçalho na primeira linha
        public object ToArray(object Header)
        {
            bool header = Flag(Header, false);
            List<Column> cols = Schema();
            List<object[]> rows;
            using (Engine.QueryResult r = Engine.Query(Readable())) rows = r.ReadAll(long.MaxValue);
            object[,] grid = new object[rows.Count + (header ? 1 : 0), cols.Count];
            if (header) for (int c = 0; c < cols.Count; c++) grid[0, c] = cols[c].Name;
            for (int i = 0; i < rows.Count; i++)
                for (int c = 0; c < cols.Count; c++) grid[i + (header ? 1 : 0), c] = rows[i][c];
            return grid;
        }

        // Como o pandas: até 60 linhas inteiras; acima disso, as 5 primeiras e as 5 últimas
        public override string ToString()
        {
            if (sql == null) return "Empty DataFrame";
            long n = Count;
            List<Column> cols = Schema();
            List<object[]> rows;
            List<long> labels = new List<long>();
            bool cut = n > 60;
            string source = Readable();
            using (Engine.QueryResult r = Engine.Query(cut ? "(SELECT * FROM (" + source + ") LIMIT 5) UNION ALL (SELECT * FROM (" + source + ") OFFSET " +
                (n - 5).ToString(CultureInfo.InvariantCulture) + ")" : source))
                rows = r.ReadAll(long.MaxValue);
            for (int i = 0; i < rows.Count; i++) labels.Add(cut && i >= 5 ? n - 10 + i : i);
            return FormatTable(cols, rows, labels, cut ? 5 : -1) + Environment.NewLine + Environment.NewLine +
                "[" + n.ToString("N0", CultureInfo.CurrentCulture) + " linhas x " + cols.Count + " colunas]";
        }

        static string FormatTable(List<Column> cols, List<object[]> rows, List<long> labels, int ellipsisAfter)
        {
            int n = cols.Count;
            string[][] cells = rows.Select(r => Enumerable.Range(0, n).Select(c => Display(r[c], cols[c].Type)).ToArray()).ToArray();
            string[] index = labels.Select(l => l.ToString(CultureInfo.InvariantCulture)).ToArray();
            int indexWidth = index.Length == 0 ? 0 : index.Max(s => s.Length);
            int[] widths = new int[n];
            for (int c = 0; c < n; c++)
                widths[c] = Math.Min(30, Math.Max(cols[c].Name.Length, cells.Length == 0 ? 0 : cells.Max(r => r[c].Length)));
            StringBuilder sb = new StringBuilder();
            sb.Append(new string(' ', indexWidth));
            for (int c = 0; c < n; c++) sb.Append("  ").Append(Fit(cols[c].Name, widths[c], true));
            for (int r = 0; r < cells.Length; r++)
            {
                if (r == ellipsisAfter)
                {
                    sb.AppendLine().Append(Fit("...", indexWidth, false));
                    for (int c = 0; c < n; c++) sb.Append("  ").Append(Fit("...", widths[c], true));
                }
                sb.AppendLine().Append(index[r].PadRight(indexWidth));
                for (int c = 0; c < n; c++) sb.Append("  ").Append(Fit(cells[r][c], widths[c], IsNumeric(cols[c].Type)));
            }
            return sb.ToString();
        }

        static string Fit(string s, int width, bool right)
        {
            if (s.Length > width) s = s.Substring(0, Math.Max(0, width - 1)) + "…";
            return right ? s.PadLeft(width) : s.PadRight(width);
        }

        internal static string Display(object value, string type)
        {
            if (value == null) return "Empty"; // o vazio do VBA, de qualquer tipo de coluna
            if (value is DateTime)
            {
                DateTime d = (DateTime)value;
                return d.ToString(d.TimeOfDay == TimeSpan.Zero && type == "DATE" ? "yyyy-MM-dd" : "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }
            if (value is double)
            {
                double v = (double)value;
                return Math.Abs(v) >= 1e15 || (v != 0 && Math.Abs(v) < 1e-6) ? v.ToString("0.######E+0", CultureInfo.CurrentCulture) : v.ToString("0.######", CultureInfo.CurrentCulture);
            }
            if (value is bool) return (bool)value ? "True" : "False";
            return Convert.ToString(value, CultureInfo.CurrentCulture);
        }

        public void Show(object Title)
        {
            if (sql == null) throw Empty();
            EnsureTable();
            string title = Interop.IsMissing(Title) ? "" : Interop.Text(Title);
            // Na aba "DataFrame Results" da janela Terminal do editor do VBA (sem título, a
            // barra de baixo fica só com colunas e linhas); sem o SageEditor carregado, numa
            // janela própria
            if (DataFrameView.ShowInEditor(this, title)) return;
            DataFrameView.ShowDialog(this, title.Length == 0 ? "DataFrame" : title);
        }

        // Para a janela: as colunas e um intervalo de linhas (pela posição na tabela própria)
        internal List<string> ColumnNames { get { return Schema().Select(c => c.Name).ToList(); } }
        internal List<string> ColumnTypes { get { return Schema().Select(c => c.Type).ToList(); } }

        internal List<object[]> Rows(long start, int length)
        {
            EnsureTable();
            using (Engine.QueryResult r = Engine.Query(ReadableOf("SELECT * FROM " + Engine.Quote(own.Name) + " WHERE rowid >= " +
                start.ToString(CultureInfo.InvariantCulture) + " AND rowid < " + (start + length).ToString(CultureInfo.InvariantCulture) + " ORDER BY rowid")))
                return r.ReadAll(length);
        }

        // ------------------------------------------------------------------
        // Informações
        // ------------------------------------------------------------------

        public long Count
        {
            get
            {
                if (sql == null) return 0;
                // Sem Sync: alterações pendentes do cursor não mudam o número de linhas
                if (countFor != sql)
                {
                    count = Convert.ToInt64(Engine.Scalar("SELECT count(*) FROM (" + sql + ")"), CultureInfo.InvariantCulture);
                    countFor = sql;
                }
                return count;
            }
        }

        public ListS Columns { get { return ListS.From(Schema().Select(c => (object)c.Name)); } }

        public DictionaryS DTypes
        {
            get
            {
                DictionaryS d = new DictionaryS();
                foreach (Column c in Schema()) d.LetValue(c.Name, c.Type);
                return d;
            }
        }

        // (linhas, colunas), como df.shape
        public ListS Shape
        {
            get
            {
                ListS shape = ListS.From(new object[] { Count, (long)Schema().Count });
                shape.ArrayType = sgArrayTypes.sgTuple;
                return shape;
            }
        }

        public string Info()
        {
            List<Column> cols = Schema();
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<Sage.DataFrame>");
            sb.AppendLine(Count.ToString("N0", CultureInfo.CurrentCulture) + " linhas, " + cols.Count + " colunas");
            if (cols.Count > 0)
            {
                string counts = string.Join(", ", cols.Select(c => "count(" + Engine.Quote(c.Name) + ")"));
                object[] nonNull;
                using (Engine.QueryResult r = Engine.Query("SELECT " + counts + " FROM (" + sql + ")")) nonNull = r.ReadAll(1)[0];
                int width = Math.Max(6, cols.Max(c => c.Name.Length));
                sb.AppendLine(" #  " + "Coluna".PadRight(width) + "  Não nulos     Tipo");
                for (int i = 0; i < cols.Count; i++)
                    sb.AppendLine(i.ToString(CultureInfo.InvariantCulture).PadLeft(2) + "  " + cols[i].Name.PadRight(width) + "  " +
                        Convert.ToInt64(nonNull[i], CultureInfo.InvariantCulture).ToString("N0", CultureInfo.CurrentCulture).PadLeft(12) + "  " + cols[i].Type);
            }
            return sb.ToString().TrimEnd();
        }

        // Estatísticas por coluna (SUMMARIZE do DuckDB): mínimo, máximo, média, desvio,
        // quartis, distintos (aproximado), contagem e % de nulos
        public DataFrame Describe()
        {
            return Derive("SELECT * FROM (SUMMARIZE " + Need() + ")");
        }

        public DataFrame Head(object Rows)
        {
            return Derive("SELECT * FROM (" + Need() + ") LIMIT " + Number(Rows, 5));
        }

        public DataFrame Tail(object Rows)
        {
            long n = Count, rows = long.Parse(Number(Rows, 5), CultureInfo.InvariantCulture);
            return Derive("SELECT * FROM (" + Need() + ") OFFSET " + Math.Max(0, n - rows).ToString(CultureInfo.InvariantCulture));
        }

        // ------------------------------------------------------------------
        // Selecionar e filtrar
        // ------------------------------------------------------------------

        public DataFrame Query(string Condition)
        {
            return Derive("SELECT * FROM (" + Need() + ") WHERE " + Expressions.Translate(Condition));
        }

        // Colunas ("A, B" ou Array/ListS); também aceita "Novo = expressão"
        public DataFrame Select(object Columns)
        {
            List<string> items = new List<string>();
            foreach (string part in Names(Columns))
            {
                string name, expression;
                if (Expressions.Assignment(part, out name, out expression))
                    items.Add("(" + Expressions.Translate(expression) + ") AS " + Engine.Quote(name));
                else items.Add(Identifier(part));
            }
            return Derive("SELECT " + string.Join(", ", items) + " FROM (" + Need() + ")");
        }

        public DataFrame Drop(object Columns)
        {
            return Derive("SELECT * EXCLUDE (" + string.Join(", ", Names(Columns).Select(n => Engine.Quote(Resolve(n)))) + ") FROM (" + Need() + ")");
        }

        // Como iloc[start:stop] (negativos contam do fim)
        public DataFrame Slice(object Start, object Stop)
        {
            long n = Count;
            long start = Interop.IsMissing(Start) ? 0 : Interop.Integer(Start), stop = Interop.IsMissing(Stop) ? n : Interop.Integer(Stop);
            if (start < 0) start = Math.Max(0, n + start);
            if (stop < 0) stop = Math.Max(0, n + stop);
            stop = Math.Min(stop, n);
            return Derive("SELECT * FROM (" + Need() + ") LIMIT " + Math.Max(0, stop - start).ToString(CultureInfo.InvariantCulture) +
                " OFFSET " + start.ToString(CultureInfo.InvariantCulture));
        }

        public DataFrame Sample(long Rows)
        {
            return Derive("SELECT * FROM (" + Need() + ") USING SAMPLE " + Rows.ToString(CultureInfo.InvariantCulture) + " ROWS");
        }

        // Coluna inteira, como df["col"]
        public ListS Col(object Column)
        {
            string name = Resolve(Column);
            using (Engine.QueryResult r = Engine.Query(ReadableOf("SELECT " + Engine.Quote(name) + " FROM (" + Need() + ")")))
                return ListS.From(r.ReadAll(long.MaxValue).Select(row => row[0]));
        }

        // Valores distintos, na ordem em que aparecem
        public ListS Unique(object Column)
        {
            string c = Engine.Quote(Resolve(Column));
            string inner = "SELECT " + c + ", row_number() OVER () AS __n FROM (" + Need() + ")";
            using (Engine.QueryResult r = Engine.Query(ReadableOf("SELECT " + c + " FROM (" + inner + ") GROUP BY " + c + " ORDER BY min(__n)")))
                return ListS.From(r.ReadAll(long.MaxValue).Select(row => row[0]));
        }

        // ------------------------------------------------------------------
        // Transformar
        // ------------------------------------------------------------------

        // "Col = expressão" cria ou substitui a coluna (várias separadas por ";" ou linhas);
        // sem "=", devolve o resultado da expressão como ListS, como df.eval("a + b")
        public object Eval(string Expression)
        {
            Need();
            List<string> statements = Expressions.Split(Expression, ';', true);
            if (statements.Count == 1)
            {
                string single, expr;
                if (!Expressions.Assignment(statements[0], out single, out expr))
                {
                    using (Engine.QueryResult r = Engine.Query(ReadableOf("SELECT (" + Expressions.Translate(statements[0]) + ") AS value FROM (" + sql + ")")))
                        return ListS.From(r.ReadAll(long.MaxValue).Select(row => row[0]));
                }
            }
            foreach (string statement in statements)
            {
                string name, expression;
                if (!Expressions.Assignment(statement, out name, out expression))
                    throw Interop.Error(5, "Esperado \"Coluna = expressão\": " + statement);
                Replace(name, "(" + Expressions.Translate(expression) + ")");
            }
            return this;
        }

        public DataFrame Rename(object Mapping)
        {
            Dictionary<string, string> map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<object, object> pair in Interop.PairsOf(Mapping)) map[Resolve(pair.Key)] = Interop.Text(pair.Value);
            string list = string.Join(", ", Schema().Select(c => Engine.Quote(c.Name) + (map.ContainsKey(c.Name) ? " AS " + Engine.Quote(map[c.Name]) : "")));
            return Derive("SELECT " + list + " FROM (" + Need() + ")");
        }

        // By: "A, B" ou Array; Ascending: True/False ou um por coluna
        public DataFrame SortValues(object By, object Ascending)
        {
            List<string> keys = Names(By);
            List<bool> ascending = Interop.IsMissing(Ascending) ? keys.Select(k => true).ToList()
                : Interop.Unwrap(Ascending) is Array || Ascending is ListS ? Interop.Items(Ascending).Select(a => Convert.ToBoolean(a, CultureInfo.InvariantCulture)).ToList()
                : keys.Select(k => Flag(Ascending, true)).ToList();
            string order = string.Join(", ", keys.Select((k, i) => Identifier(k) + (i < ascending.Count && !ascending[i] ? " DESC" : " ASC") + " NULLS LAST"));
            return Derive("SELECT * FROM (" + Need() + ") ORDER BY " + order);
        }

        public DataFrame DropDuplicates(object Subset)
        {
            if (Interop.IsMissing(Subset)) return Derive("SELECT DISTINCT * FROM (" + Need() + ")");
            return Derive("SELECT DISTINCT ON (" + string.Join(", ", Names(Subset).Select(Identifier)) + ") * FROM (" + Need() + ")");
        }

        public DataFrame DropNA(object Subset)
        {
            IEnumerable<string> cols = Interop.IsMissing(Subset) ? Schema().Select(c => c.Name) : Names(Subset).Select(Resolve);
            return Derive("SELECT * FROM (" + Need() + ") WHERE " + string.Join(" AND ", cols.Select(c => Engine.Quote(c) + " IS NOT NULL")));
        }

        public DataFrame FillNA(object Value, object Columns)
        {
            IEnumerable<string> cols = Interop.IsMissing(Columns) ? Schema().Select(c => c.Name) : Names(Columns).Select(Resolve);
            string literal = Engine.Literal(Value);
            return Derive("SELECT * REPLACE (" + string.Join(", ", cols.Select(c => "coalesce(" + Engine.Quote(c) + ", " + literal + ") AS " + Engine.Quote(c))) +
                ") FROM (" + Need() + ")");
        }

        // Cópia independente (tabela própria), como df.copy()
        public void ToSql(string Name, SqlEngine Engine, object IfExists)
        {
            if (Engine == null) throw Interop.Error(91, "ToSql: informe o SqlEngine (Set db = New Sage.SqlEngine: db.Connect ...).");
            string mode = Interop.IsMissing(IfExists) ? "fail" : Interop.Text(IfExists).Trim().ToLowerInvariant();
            if (mode != "fail" && mode != "replace" && mode != "append")
                throw Interop.Error(5, "ValueError: IfExists deve ser \"fail\", \"replace\" ou \"append\".");
            Engine.WriteTable(this, Name, mode);
        }

        // Consulta do DataFrame (com as alterações pendentes gravadas), para ler as linhas em outro lugar
        internal string Source { get { return Need(); } }
        // A mesma, com os tipos que o leitor de blocos entende (decimais como Double...)
        internal string ReadableSource { get { return Readable(); } }

        // Tabela nova preenchida por fora (resultado de um banco de dados): fill recebe o nome dela
        internal static DataFrame FromTable(Action<string> fill)
        {
            Table table = new Table();
            fill(table.Name);
            DataFrame df = new DataFrame();
            df.Reset("SELECT * FROM " + Engine.Quote(table.Name), new List<Table> { table }, table);
            return df;
        }

        public DataFrame Copy()
        {
            DataFrame copy = new DataFrame();
            copy.MaterializeFrom(Need(), Share());
            return copy;
        }

        // ------------------------------------------------------------------
        // Combinar e resumir
        // ------------------------------------------------------------------

        // Keys: "A, B"; Aggregations: "Total = sum(Valor), N = count(*)", ou um DictionaryS
        // {coluna: função}; sem agregações, conta as linhas de cada grupo (como size())
        public DataFrame GroupBy(object Keys, object Aggregations)
        {
            List<string> keys = Names(Keys).Select(Identifier).ToList();
            List<string> aggs = new List<string>();
            object raw = Interop.Unwrap(Aggregations);
            if (Interop.IsMissing(Aggregations)) aggs.Add("count(*) AS " + Engine.Quote("count"));
            else if (raw is string)
            {
                foreach (string part in Expressions.Split((string)raw, ',', false))
                {
                    string name, expression;
                    if (Expressions.Assignment(part, out name, out expression)) aggs.Add(Expressions.Translate(expression) + " AS " + Engine.Quote(name));
                    else aggs.Add(Expressions.Translate(part));
                }
            }
            else
                foreach (KeyValuePair<object, object> pair in Interop.PairsOf(Aggregations))
                {
                    string col = Resolve(pair.Key);
                    aggs.Add(Interop.Text(pair.Value) + "(" + Engine.Quote(col) + ") AS " + Engine.Quote(col));
                }
            string groups = string.Join(", ", keys);
            return Derive("SELECT " + groups + ", " + string.Join(", ", aggs) + " FROM (" + Need() + ") GROUP BY " + groups + " ORDER BY " + groups);
        }

        // Tabela dinâmica: linhas de Index, uma coluna por valor de Columns, Function(Values)
        public DataFrame Pivot(object Index, string Columns, string Values, object Function)
        {
            string function = Interop.IsMissing(Function) ? "sum" : Interop.Text(Function);
            string groups = string.Join(", ", Names(Index).Select(Identifier));
            return Derive("SELECT * FROM (PIVOT (" + Need() + ") ON " + Identifier(Columns) + " USING " + function + "(" + Identifier(Values) +
                ") GROUP BY " + groups + " ORDER BY " + groups + ")");
        }

        // Como pd.merge: On (mesmo nome dos dois lados) ou LeftOn/RightOn; How: inner, left,
        // right, outer, cross. Colunas repetidas ganham "_x" e "_y".
        public DataFrame Merge(DataFrame Other, object On, object How, object LeftOn, object RightOn)
        {
            if (Other == null) throw Interop.Error(91, "Merge: informe o outro DataFrame.");
            string how = Interop.IsMissing(How) ? "inner" : Interop.Text(How).Trim().ToLowerInvariant();
            string join = how == "left" ? "LEFT JOIN" : how == "right" ? "RIGHT JOIN" : how == "outer" || how == "full" ? "FULL OUTER JOIN"
                : how == "cross" ? "CROSS JOIN" : "INNER JOIN";
            List<Column> left = Schema(), right = Other.Schema();
            List<string> lk, rk;
            if (how == "cross") lk = rk = new List<string>();
            else if (!Interop.IsMissing(On)) { lk = Names(On).Select(Resolve).ToList(); rk = Names(On).Select(Other.Resolve).ToList(); }
            else if (!Interop.IsMissing(LeftOn) && !Interop.IsMissing(RightOn)) { lk = Names(LeftOn).Select(Resolve).ToList(); rk = Names(RightOn).Select(Other.Resolve).ToList(); }
            else
            {
                // Sem chave: as colunas com o mesmo nome nos dois lados
                lk = left.Select(c => c.Name).Where(n => right.Any(r => string.Equals(r.Name, n, StringComparison.OrdinalIgnoreCase))).ToList();
                rk = lk.Select(Other.Resolve).ToList();
            }
            bool sameKeys = Interop.IsMissing(LeftOn) && how != "cross";
            List<string> select = new List<string>();
            if (sameKeys)
                for (int i = 0; i < lk.Count; i++)
                    select.Add((how == "right" ? "r." + Engine.Quote(rk[i]) : how == "outer" || how == "full"
                        ? "coalesce(l." + Engine.Quote(lk[i]) + ", r." + Engine.Quote(rk[i]) + ")" : "l." + Engine.Quote(lk[i])) + " AS " + Engine.Quote(lk[i]));
            Func<string, List<string>, bool> isKey = (n, keys) => sameKeys && keys.Contains(n, StringComparer.OrdinalIgnoreCase);
            foreach (Column c in left.Where(c => !isKey(c.Name, lk)))
                select.Add("l." + Engine.Quote(c.Name) + " AS " + Engine.Quote(right.Any(r => !isKey(r.Name, rk) && string.Equals(r.Name, c.Name, StringComparison.OrdinalIgnoreCase)) ? c.Name + "_x" : c.Name));
            foreach (Column c in right.Where(c => !isKey(c.Name, rk)))
                select.Add("r." + Engine.Quote(c.Name) + " AS " + Engine.Quote(left.Any(l => !isKey(l.Name, lk) && string.Equals(l.Name, c.Name, StringComparison.OrdinalIgnoreCase)) ? c.Name + "_y" : c.Name));
            string on = how == "cross" ? "" : " ON " + string.Join(" AND ", lk.Select((k, i) => "l." + Engine.Quote(k) + " = r." + Engine.Quote(rk[i])));
            Other.Sync();
            List<Table> used = Share().Concat(Other.Share()).ToList();
            return new DataFrame("SELECT " + string.Join(", ", select) + " FROM (" + Need() + ") l " + join + " (" + Other.Need() + ") r" + on, used);
        }

        // Linhas dos dois, combinando as colunas pelo nome (as que faltam ficam nulas)
        public DataFrame Concat(DataFrame Other)
        {
            if (Other == null) throw Interop.Error(91, "Concat: informe o outro DataFrame.");
            Other.Sync();
            List<Table> used = Share().Concat(Other.Share()).ToList();
            return new DataFrame("SELECT * FROM ((" + Need() + ") UNION ALL BY NAME (" + Other.Need() + "))", used);
        }

        public DataFrame ValueCounts(object Column)
        {
            string c = Engine.Quote(Resolve(Column));
            return Derive("SELECT " + c + ", count(*) AS " + Engine.Quote("count") + " FROM (" + Need() + ") GROUP BY " + c + " ORDER BY 2 DESC, 1");
        }

        public object Sum(object Column) { return Aggregate("sum", Column); }
        public object Mean(object Column) { return Aggregate("avg", Column); }
        public object Min(object Column) { return Aggregate("min", Column); }
        public object Max(object Column) { return Aggregate("max", Column); }
        public object Median(object Column) { return Aggregate("median", Column); }
        public object Std(object Column) { return Aggregate("stddev_samp", Column); }

        public long NUnique(object Column)
        {
            return Convert.ToInt64(Engine.Scalar("SELECT count(DISTINCT " + Engine.Quote(Resolve(Column)) + ") FROM (" + Need() + ")"), CultureInfo.InvariantCulture);
        }

        object Aggregate(string function, object column)
        {
            string c = Engine.Quote(Resolve(column));
            object value = Engine.Scalar(ReadableOf("SELECT " + function + "(" + c + ") AS " + c + " FROM (" + Need() + ")"));
            return Interop.Wrap(value);
        }

        // ------------------------------------------------------------------
        // Cursor
        // ------------------------------------------------------------------

        public void MoveFirst() { Position(0); }
        public void MoveLast() { Position(Count - 1); }

        public void MoveNext()
        {
            if (!cursor) { Position(0); return; }
            if (EOF) throw Interop.Error(3021, "EOF: não há próxima linha.");
            index++;
        }

        public void MovePrevious()
        {
            if (!cursor || BOF) throw Interop.Error(3021, "BOF: não há linha anterior.");
            index--;
        }

        public void Move(long Rows) { Position((cursor ? index : -1) + Rows); }

        public bool EOF { get { return sql == null || (cursor && index >= Count); } }
        public bool BOF { get { return sql == null || !cursor || index < 0; } }

        public long Index
        {
            get { return cursor ? index : -1; }
            set { Position(value); }
        }

        void Position(long row)
        {
            if (sql == null) throw Empty();
            EnsureTable();
            cursor = true;
            long n = Count;
            index = row < 0 ? -1 : row >= n ? n : row;
        }

        // ------------------------------------------------------------------
        // At e Loc
        // ------------------------------------------------------------------

        public object GetAt(long Row, object Column)
        {
            EnsureTable();
            CheckRow(Row, false);
            return Interop.Wrap(Cell(Row, Resolve(Column)));
        }

        public void LetAt(long Row, object Column, object Value)
        {
            EnsureTable();
            CheckRow(Row, false);
            SetCell(Row, Interop.Text(Column).Trim(), Value);
        }

        // Linha como DictionaryS {coluna: valor}
        public DictionaryS GetLoc(long Row)
        {
            EnsureTable();
            CheckRow(Row, false);
            DictionaryS d = new DictionaryS();
            foreach (Column c in Schema()) d.LetValue(c.Name, Cell(Row, c.Name));
            return d;
        }

        // Array de valores na ordem das colunas; Row = Count acrescenta uma linha, como
        // df.loc[len(df)] = [...]
        public void LetLoc(long Row, object Values)
        {
            EnsureTable();
            CheckRow(Row, true);
            List<object> values = Interop.Items(Values).ToList();
            List<Column> cols = Schema();
            if (values.Count > cols.Count) throw Interop.Error(9, "Mais valores do que colunas.");
            if (Row == Count)
            {
                Flush();
                Engine.Run("INSERT INTO " + Engine.Quote(own.Name) + " DEFAULT VALUES");
                countFor = null;
                InvalidateBlock();
            }
            for (int c = 0; c < values.Count; c++) SetCell(Row, cols[c].Name, values[c]);
        }

        void CheckRow(long row, bool allowAppend)
        {
            if (row < 0 || row > Count || (row == Count && !allowAppend))
                throw Interop.Error(9, "IndexError: linha " + row.ToString(CultureInfo.InvariantCulture) + " fora do intervalo.");
        }

        // For Each linha In df: cada linha como DictionaryS
        public IEnumerator GetEnumerator()
        {
            List<Column> cols = Schema();
            List<object> rows = new List<object>();
            using (Engine.QueryResult r = Engine.Query(Readable()))
                foreach (object[] row in r.ReadAll(long.MaxValue))
                {
                    DictionaryS d = new DictionaryS();
                    for (int c = 0; c < cols.Count; c++) d.LetValue(cols[c].Name, row[c]);
                    rows.Add(d);
                }
            return rows.GetEnumerator();
        }

        // ------------------------------------------------------------------
        // Células (tabela própria, em blocos)
        // ------------------------------------------------------------------

        object Cell(long row, string column)
        {
            int c = Schema().FindIndex(x => string.Equals(x.Name, column, StringComparison.OrdinalIgnoreCase));
            Dictionary<long, object> changes;
            if (pending.TryGetValue(column, out changes) && changes.ContainsKey(row)) return changes[row];
            if (row < blockStart || row >= blockStart + blockRows || block == null)
            {
                blockStart = row - row % BlockSize;
                using (Engine.QueryResult r = Engine.Query(ReadableOf("SELECT * FROM " + Engine.Quote(own.Name) + " WHERE rowid >= " +
                    blockStart.ToString(CultureInfo.InvariantCulture) + " AND rowid < " + (blockStart + BlockSize).ToString(CultureInfo.InvariantCulture) + " ORDER BY rowid")))
                {
                    block = new object[Schema().Count][];
                    List<object[][]> chunks = new List<object[][]>();
                    List<int> sizes = new List<int>();
                    int rows;
                    object[][] chunk;
                    while ((chunk = r.NextChunk(out rows)) != null) { chunks.Add(chunk); sizes.Add(rows); }
                    blockRows = sizes.Sum();
                    for (int col = 0; col < block.Length; col++)
                    {
                        block[col] = new object[blockRows];
                        int at = 0;
                        for (int k = 0; k < chunks.Count; k++) { Array.Copy(chunks[k][col], 0, block[col], at, sizes[k]); at += sizes[k]; }
                    }
                }
            }
            return block[c][row - blockStart];
        }

        void SetCell(long row, string column, object value)
        {
            value = Interop.Unwrap(value);
            if (value is ListS || value is DictionaryS || value is Array) throw Interop.Error(13, "Uma célula guarda um valor só (texto, número, data...).");
            EnsureTable();
            if (!Schema().Any(c => string.Equals(c.Name, column, StringComparison.OrdinalIgnoreCase)))
            {
                // Coluna nova: criada com o tipo do valor; as outras linhas ficam nulas
                Flush();
                Engine.Run("ALTER TABLE " + Engine.Quote(own.Name) + " ADD COLUMN " + Engine.Quote(column) + " " + (TypeOf(value) ?? "VARCHAR"));
                schemaFor = null;
                InvalidateBlock();
            }
            // Gravado em lote: quando o cursor passa para outro bloco, ou antes de qualquer
            // outra operação (Need/Sync)
            long rowBlock = row / BlockSize;
            if (pendingBlock >= 0 && rowBlock != pendingBlock) Flush();
            string name = Resolve(column);
            Dictionary<long, object> changes;
            if (!pending.TryGetValue(name, out changes)) pending[name] = changes = new Dictionary<long, object>();
            changes[row] = value;
            pendingBlock = rowBlock;
        }

        // Grava as alterações pendentes: alarga o tipo da coluna se preciso (inteiro que
        // recebe decimal vira DOUBLE, número que recebe texto vira VARCHAR...) e atualiza
        // tudo de uma vez por uma tabela auxiliar
        void Flush()
        {
            if (pending.Count == 0 || own == null) return;
            List<string> columns = pending.Keys.ToList();
            foreach (string name in columns)
            {
                Dictionary<long, object> changes = pending[name];
                Column col = Schema().First(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
                string type = col.Type;
                foreach (object v in changes.Values) type = Widen(type, TypeOf(v));
                if (type != col.Type)
                {
                    Engine.Run("ALTER TABLE " + Engine.Quote(own.Name) + " ALTER " + Engine.Quote(col.Name) + " TYPE " + type);
                    schemaFor = null;
                }
                string helper = Engine.NewTable();
                Engine.Run("CREATE TABLE " + Engine.Quote(helper) + " (rid BIGINT, v " + type + ")");
                try
                {
                    int kind = Kind(type);
                    using (Engine.Appender a = new Engine.Appender(helper))
                        foreach (KeyValuePair<long, object> change in changes)
                        {
                            a.Add(change.Key, Engine.TBigInt);
                            a.Add(change.Value, kind);
                            a.EndRow();
                        }
                    Engine.Run("UPDATE " + Engine.Quote(own.Name) + " SET " + Engine.Quote(col.Name) + " = h.v FROM " + Engine.Quote(helper) +
                        " h WHERE " + Engine.Quote(own.Name) + ".rowid = h.rid");
                }
                finally { Engine.Run("DROP TABLE IF EXISTS " + Engine.Quote(helper)); }
            }
            pending.Clear();
            pendingBlock = -1;
            InvalidateBlock();
        }

        void InvalidateBlock() { block = null; blockStart = -1; blockRows = 0; }

        // Fora do cursor: a coluna inteira recebe um valor, ou um array/ListS com um valor por linha
        void SetColumn(string name, object value)
        {
            Need();
            object raw = Interop.Unwrap(value);
            if (raw is Array || value is ListS)
            {
                List<object> values = (raw is Array && ((Array)raw).Rank == 2) ? ((Array)raw).Cast<object>().ToList() : Interop.Items(value).ToList();
                if (values.Count != Count)
                    throw Interop.Error(5, "ValueError: " + values.Count + " valores para " + Count + " linhas.");
                EnsureTable();
                string type = null;
                foreach (object v in values) type = Widen(type, TypeOf(v));
                if (!Schema().Any(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase)))
                {
                    Engine.Run("ALTER TABLE " + Engine.Quote(own.Name) + " ADD COLUMN " + Engine.Quote(name) + " " + (type ?? "VARCHAR"));
                    schemaFor = null;
                }
                string resolved = Resolve(name);
                Dictionary<long, object> changes = new Dictionary<long, object>(values.Count);
                for (int i = 0; i < values.Count; i++) changes[i] = values[i];
                pending[resolved] = changes;
                Flush();
                return;
            }
            Replace(name, Engine.Literal(value));
        }

        // Cria ou substitui uma coluna por uma expressão SQL (preguiçoso)
        void Replace(string name, string expression)
        {
            Sync();
            Column existing = Schema().FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
            string next = existing != null
                ? "SELECT * REPLACE (" + expression + " AS " + Engine.Quote(existing.Name) + ") FROM (" + sql + ")"
                : "SELECT *, " + expression + " AS " + Engine.Quote(name) + " FROM (" + sql + ")";
            Reset(next, deps, null, true); // o cursor continua na mesma linha (Eval dentro de um laço)
        }

        // ------------------------------------------------------------------
        // Plano e tabelas
        // ------------------------------------------------------------------

        string Need()
        {
            if (sql == null) throw Empty();
            Sync();
            return sql;
        }

        static Exception Empty() { return Interop.Error(5, "DataFrame vazio: carregue dados antes (ReadCsv, FromRange, FromArray...)."); }

        void Sync() { Flush(); }

        // Tabelas usadas por um DataFrame derivado deste: a própria passa a ser compartilhada
        IEnumerable<Table> Share()
        {
            if (own != null) own.Shared = true;
            return deps;
        }

        DataFrame Derive(string next)
        {
            return new DataFrame(next, Share());
        }

        // Carrega numa tabela própria (arquivos: lidos uma vez só)
        DataFrame Materialize(string source)
        {
            MaterializeFrom(source, new Table[0]);
            return this;
        }

        void MaterializeFrom(string source, IEnumerable<Table> keepAlive)
        {
            Table table = new Table();
            Engine.Run("CREATE TABLE " + Engine.Quote(table.Name) + " AS " + source);
            GC.KeepAlive(keepAlive);
            Reset("SELECT * FROM " + Engine.Quote(table.Name), new List<Table> { table }, table);
        }

        // Tabela própria e exclusiva para o cursor e as alterações em lote
        void EnsureTable()
        {
            if (sql == null) throw Empty();
            Flush();
            if (own != null && !own.Shared && sql == "SELECT * FROM " + Engine.Quote(own.Name)) return;
            Table table = new Table();
            Engine.Run("CREATE TABLE " + Engine.Quote(table.Name) + " AS " + sql);
            Reset("SELECT * FROM " + Engine.Quote(table.Name), new List<Table> { table }, table, true);
        }

        // keepCursor: a mesma tabela lógica (Eval, cópia para alterar); senão, dados novos
        void Reset(string next, List<Table> tables, Table table, bool keepCursor = false)
        {
            sql = next;
            deps = tables.ToList();
            own = table;
            schemaFor = null;
            if (!keepCursor) countFor = null;
            else countFor = countFor == null ? null : next; // linhas iguais
            InvalidateBlock();
            if (keepCursor) return;
            cursor = false;
            index = -1;
        }

        List<Column> Schema()
        {
            if (sql == null) return new List<Column>();
            // Sem Sync: as alterações pendentes não mudam as colunas (colunas novas e tipos
            // alargados passam por Flush antes)
            if (schemaFor != sql)
            {
                List<Column> cols = new List<Column>();
                using (Engine.QueryResult r = Engine.Query("DESCRIBE SELECT * FROM (" + sql + ")"))
                    foreach (object[] row in r.ReadAll(long.MaxValue))
                        cols.Add(new Column { Name = Interop.Text(row[0]), Type = Interop.Text(row[1]) });
                schema = cols;
                schemaFor = sql;
            }
            return schema;
        }

        // Nome real da coluna (sem diferenciar maiúsculas); número = posição
        string Resolve(object column)
        {
            object raw = Interop.Unwrap(column);
            List<Column> cols = Schema();
            if (PyOrder.IsNumber(raw))
            {
                int i = Interop.Integer(raw);
                if (i < 0) i += cols.Count;
                if (i < 0 || i >= cols.Count) throw Interop.Error(9, "IndexError: coluna " + raw + " fora do intervalo.");
                return cols[i].Name;
            }
            string name = Interop.Text(raw).Trim();
            if (name.StartsWith("[") && name.EndsWith("]")) name = name.Substring(1, name.Length - 2);
            Column found = cols.FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
            if (found == null) throw Interop.Error(9, "KeyError: coluna '" + name + "' não existe. Colunas: " + string.Join(", ", cols.Select(c => c.Name)));
            return found.Name;
        }

        // Coluna existente (entre aspas) ou expressão
        string Identifier(string text)
        {
            string name = text.Trim();
            if (name.StartsWith("[") && name.EndsWith("]")) name = name.Substring(1, name.Length - 2);
            Column found = Schema().FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
            return found != null ? Engine.Quote(found.Name) : Expressions.Translate(text);
        }

        // "A, B", Array ou ListS -> nomes
        static List<string> Names(object value)
        {
            object raw = Interop.Unwrap(value);
            if (raw is string) return Expressions.Split((string)raw, ',', false).Select(s => s.Trim()).Where(s => s.Length > 0).ToList();
            return Interop.Items(value).Select(v => Interop.Text(v).Trim()).ToList();
        }

        // Consulta com os tipos que o VBA entende (decimais e inteiros grandes como Double,
        // datas com fuso ou precisão diferente como Date, o resto como texto)
        string Readable() { return ReadableOf(Need()); }

        static string ReadableOf(string source)
        {
            List<string> items = new List<string>();
            bool changed = false;
            using (Engine.QueryResult r = Engine.Query("DESCRIBE SELECT * FROM (" + source + ")"))
                foreach (object[] row in r.ReadAll(long.MaxValue))
                {
                    string name = Engine.Quote(Interop.Text(row[0])), type = Interop.Text(row[1]);
                    string target = ReadableType(type);
                    if (target == null) items.Add(name);
                    else { items.Add("CAST(" + name + " AS " + target + ") AS " + name); changed = true; }
                }
            return changed ? "SELECT " + string.Join(", ", items) + " FROM (" + source + ")" : source;
        }

        static readonly HashSet<string> Direct = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "BOOLEAN", "TINYINT", "SMALLINT", "INTEGER", "BIGINT", "UTINYINT", "USMALLINT", "UINTEGER", "UBIGINT",
            "FLOAT", "DOUBLE", "DATE", "TIMESTAMP", "VARCHAR",
        };

        static string ReadableType(string type)
        {
            if (Direct.Contains(type)) return null;
            if (type.StartsWith("DECIMAL", StringComparison.OrdinalIgnoreCase) || type == "HUGEINT" || type == "UHUGEINT") return "DOUBLE";
            if (type.StartsWith("TIMESTAMP", StringComparison.OrdinalIgnoreCase)) return "TIMESTAMP";
            return "VARCHAR";
        }

        static bool IsNumeric(string type)
        {
            return type.EndsWith("INT") || type == "INTEGER" || type == "DOUBLE" || type == "FLOAT" || type.StartsWith("DECIMAL") || type.EndsWith("HUGEINT");
        }

        // Tipo do DuckDB para um valor do VBA (null: vazio, não decide)
        static string TypeOf(object value)
        {
            value = Interop.Unwrap(value);
            if (value == null || value is DBNull || Interop.IsMissing(value)) return null;
            if (value is bool) return "BOOLEAN";
            if (value is byte || value is short || value is int || value is long || value is sbyte || value is ushort || value is uint) return "BIGINT";
            if (value is double || value is float || value is decimal || value is ulong) return "DOUBLE";
            if (value is DateTime) return ((DateTime)value).TimeOfDay == TimeSpan.Zero ? "DATE" : "TIMESTAMP";
            return "VARCHAR";
        }

        // Tipo que comporta os dois (como o pandas ao misturar valores numa coluna)
        static string Widen(string current, string next)
        {
            if (next == null) return current;
            if (current == null || current == next) return next;
            bool currentNumber = IsNumeric(current) || current == "BOOLEAN", nextNumber = IsNumeric(next) || next == "BOOLEAN";
            if (currentNumber && nextNumber)
            {
                if (current == "DOUBLE" || next == "DOUBLE" || current == "FLOAT" || current.StartsWith("DECIMAL")) return "DOUBLE";
                if (current == "BOOLEAN" && next == "BOOLEAN") return "BOOLEAN";
                return "BIGINT";
            }
            bool currentDate = current == "DATE" || current.StartsWith("TIMESTAMP"), nextDate = next == "DATE" || next.StartsWith("TIMESTAMP");
            if (currentDate && nextDate) return "TIMESTAMP";
            return "VARCHAR";
        }

        static int Kind(string type)
        {
            if (type == "BOOLEAN") return Engine.TBoolean;
            if (type.EndsWith("INT") || type == "INTEGER") return Engine.TBigInt;
            if (IsNumeric(type)) return Engine.TDouble;
            if (type == "DATE") return Engine.TDate;
            if (type.StartsWith("TIMESTAMP")) return Engine.TTimestamp;
            return Engine.TVarchar;
        }

        static bool Flag(object value, bool fallback)
        {
            if (Interop.IsMissing(value) || value == null) return fallback;
            return Convert.ToBoolean(Interop.Unwrap(value), CultureInfo.InvariantCulture);
        }

        static string Number(object value, long fallback)
        {
            long n = Interop.IsMissing(value) ? fallback : Interop.Integer(value);
            return Math.Max(0, n).ToString(CultureInfo.InvariantCulture);
        }
    }

    // ------------------------------------------------------------------
    // Expressões: SQL do DuckDB com hábitos de VBA
    // ------------------------------------------------------------------
    static class Expressions
    {
        // "texto" vira 'texto' (literal), [Nome] vira "Nome" (coluna); 'texto' fica como está
        public static string Translate(string expression)
        {
            StringBuilder sb = new StringBuilder();
            int i = 0;
            while (i < expression.Length)
            {
                char c = expression[i];
                if (c == '"' || c == '\'')
                {
                    StringBuilder literal = new StringBuilder();
                    int j = i + 1;
                    while (j < expression.Length)
                    {
                        if (expression[j] == c)
                        {
                            if (j + 1 < expression.Length && expression[j + 1] == c) { literal.Append(c); j += 2; continue; }
                            break;
                        }
                        literal.Append(expression[j]);
                        j++;
                    }
                    sb.Append('\'').Append(literal.ToString().Replace("'", "''")).Append('\'');
                    i = j + 1;
                }
                else if (c == '[')
                {
                    int end = expression.IndexOf(']', i + 1);
                    if (end < 0) { sb.Append(c); i++; continue; }
                    sb.Append(Engine.Quote(expression.Substring(i + 1, end - i - 1)));
                    i = end + 1;
                }
                else { sb.Append(c); i++; }
            }
            return sb.ToString();
        }

        // Divide nos separadores fora de parênteses e de textos (newline também, se pedido)
        public static List<string> Split(string text, char separator, bool lines)
        {
            List<string> parts = new List<string>();
            StringBuilder current = new StringBuilder();
            int depth = 0;
            char quote = '\0';
            foreach (char c in text)
            {
                if (quote != '\0') { current.Append(c); if (c == quote) quote = '\0'; continue; }
                if (c == '"' || c == '\'') { quote = c; current.Append(c); continue; }
                if (c == '(' || c == '[') depth++;
                if (c == ')' || c == ']') depth--;
                if (depth == 0 && (c == separator || (lines && (c == '\n' || c == '\r'))))
                {
                    if (current.ToString().Trim().Length > 0) parts.Add(current.ToString().Trim());
                    current.Clear();
                    continue;
                }
                current.Append(c);
            }
            if (current.ToString().Trim().Length > 0) parts.Add(current.ToString().Trim());
            return parts;
        }

        // "Nome = expressão" (o "=" de atribuição, não o de comparação: só um nome antes dele)
        public static bool Assignment(string statement, out string name, out string expression)
        {
            name = expression = null;
            string s = statement.Trim();
            int eq = -1;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '=' && (i + 1 >= s.Length || s[i + 1] != '=') && (i == 0 || "<>!=".IndexOf(s[i - 1]) < 0)) { eq = i; break; }
                if (s[i] == '"' || s[i] == '\'' || s[i] == '(') return false;
            }
            if (eq <= 0) return false;
            string left = s.Substring(0, eq).Trim();
            if (left.StartsWith("[") && left.EndsWith("]")) left = left.Substring(1, left.Length - 2);
            else if (!left.All(ch => char.IsLetterOrDigit(ch) || ch == '_')) return false;
            if (left.Length == 0) return false;
            name = left;
            expression = s.Substring(eq + 1).Trim();
            return expression.Length > 0;
        }
    }

    // Excel do próprio processo (o SageTypes roda dentro dele)
    static class Excel
    {
        [DllImport("oleacc.dll")]
        static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object obj);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string title);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);

        // Caminho para os Read: completo (C:\..., \\servidor\...) ou endereço (https://, s3://)
        // fica como está; relativo ("vendas.csv", "dados\vendas.csv") parte da pasta da pasta de
        // trabalho ativa, como o Python parte da pasta atual. Pasta de trabalho não salva (ou fora
        // do Excel): a pasta atual do processo.
        public static string ResolvePath(string path)
        {
            if (string.IsNullOrEmpty(path) || path.Contains("://") || System.IO.Path.IsPathRooted(path)) return path;
            return System.IO.Path.Combine(LocalFolder(), path);
        }

        static string LocalFolder()
        {
            try
            {
                object app = InProcess();
                if (app != null)
                {
                    dynamic book = ((dynamic)app).ActiveWorkbook;
                    string folder = book == null ? "" : (string)book.Path;
                    if (!string.IsNullOrEmpty(folder) && !folder.Contains("://")) return folder;
                }
            }
            catch (Exception) { } // Excel ocupado (ex.: editando uma célula): a pasta atual
            return Environment.CurrentDirectory;
        }

        public static object Application(object fallback)
        {
            object app = InProcess();
            if (app != null) return app;
            if (fallback != null) return fallback;
            return Marshal.GetActiveObject("Excel.Application");
        }

        // O Excel deste processo (null fora do Excel)
        static object InProcess()
        {
            uint me = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            for (IntPtr main = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "XLMAIN", null); main != IntPtr.Zero; main = FindWindowEx(IntPtr.Zero, main, "XLMAIN", null))
            {
                uint pid;
                GetWindowThreadProcessId(main, out pid);
                if (pid != me) continue;
                IntPtr desk = FindWindowEx(main, IntPtr.Zero, "XLDESK", null);
                IntPtr sheet = desk == IntPtr.Zero ? IntPtr.Zero : FindWindowEx(desk, IntPtr.Zero, "EXCEL7", null);
                if (sheet == IntPtr.Zero) continue;
                Guid iid = new Guid("00020400-0000-0000-C000-000000000046");
                object window;
                if (AccessibleObjectFromWindow(sheet, 0xFFFFFFF0, ref iid, out window) == 0 && window != null) return ((dynamic)window).Application;
            }
            return null;
        }
    }
}
