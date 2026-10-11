using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SageTypes
{
    // DataGrid: mostra um DataFrame num UserForm (Caixa de ferramentas > Sage Framework DataGrid).
    //
    //   Set gridVendas.DataFrame = df
    //   gridVendas.AutoPaging = True        ' padrão: pagina quando as linhas não cabem
    //
    //   gridVendas.ColumnFormats = Array("valor", "#,##0.00", "percentual", "0.00%", "data", "dd/mm/yyyy")
    //
    //   Private Sub gridVendas_DblClick()
    //       MsgBox gridVendas.Cell(gridVendas.SelectedRow, "Cliente")
    //   End Sub
    //
    // Com AutoPaging, a grade calcula quantas linhas cabem na altura do controle; se todas
    // couberem, mostra todas; senão, mostra uma página por vez, com a barra "◀ Página 1 de N ▶"
    // embaixo (e refaz a conta quando o controle muda de tamanho). Sem AutoPaging, rola. Só as
    // linhas à vista são lidas do DuckDB, então funciona com milhões de linhas.
    //
    // Interface dual: acrescente membros só no fim, com o próximo DispId (ver StringS). As
    // propriedades gravadas no arquivo usam os nomes de SaveProperties: não os mude.
    [ComVisible(true), Guid("44893F4B-2FF3-4412-B6B2-5F942C3C4C34"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _DataGrid
    {
        [DispId(1)] DataFrame DataFrame { get; set; }
        [DispId(2)] bool AutoPaging { get; set; }
        [DispId(3)] int Page { get; set; }
        [DispId(4)] int PageCount { get; }
        [DispId(5)] int RowsPerPage { get; }
        [DispId(6)] long SelectedRow { get; set; }
        [DispId(7)] long RowCount { get; }
        [DispId(8)] object Cell(long Row, object Column);
        [DispId(9)] void NextPage();
        [DispId(10)] void PreviousPage();
        [DispId(11)] bool ShowIndex { get; set; }
        [DispId(12)] string FontName { get; set; }
        [DispId(13)] double FontSize { get; set; }
        [DispId(-501), ComAliasName("stdole.OLE_COLOR")] int BackColor { get; set; }
        [DispId(-513), ComAliasName("stdole.OLE_COLOR")] int ForeColor { get; set; }
        [DispId(14), ComAliasName("stdole.OLE_COLOR")] int BorderColor { get; set; }
        [DispId(15), ComAliasName("stdole.OLE_COLOR")] int HeaderBackColor { get; set; }
        [DispId(16)] object ColumnFormats { get; set; }
        [DispId(1016), PropertyLet("ColumnFormats")] void LetColumnFormats(object Value);
        [DispId(17)] string SelectedColumn { get; }
        [DispId(18)] void Copy();
        [DispId(19), ComAliasName("stdole.OLE_COLOR")] int SelectionColor { get; set; }
    }

    [ComVisible(true), Guid("34F0EAA8-EE62-41E7-95DF-5D41E589054E"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface __DataGrid
    {
        [DispId(1)] void Click();
        [DispId(2)] void DblClick();
        [DispId(3)] void SelChange();
        [DispId(4)] void PageChange();
    }

    public delegate void DataGridEventHandler();

    [ComVisible(true), Guid("116738F7-D352-40B2-8E87-2B0294450F9E"), ProgId("Sage.DataGrid")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_DataGrid)), ComSourceInterfaces(typeof(__DataGrid))]
    public sealed class DataGrid : OleControl, _DataGrid
    {
        public event DataGridEventHandler Click;
        public event DataGridEventHandler DblClick;
        public event DataGridEventHandler SelChange;
        public event DataGridEventHandler PageChange;

        const int SystemWindow = unchecked((int)0x80000005), SystemWindowText = unchecked((int)0x80000008);
        const int DefaultBorder = 0xB3ADAB, DefaultHeader = 0xF3F3F3, DefaultSelection = 0xF7E4CC; // RGB(204, 228, 247)

        DataFrame frame;
        bool autoPaging, showIndex;
        string fontName;
        double fontSize;
        int backColor, foreColor, borderColor, headerBackColor, selectionColor;
        Font font;
        GridView view;
        // Formatos do VBA (Format) por coluna; só no código, não ficam gravados no form
        readonly Dictionary<string, string> formats = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public DataGrid() : base(240, 140) { InitProperties(); }

        protected override string TypeName { get { return "Sage DataGrid"; } }

        protected override void InitProperties()
        {
            autoPaging = true; showIndex = true;
            fontName = "Segoe UI"; fontSize = 9;
            backColor = SystemWindow; foreColor = SystemWindowText; borderColor = DefaultBorder; headerBackColor = DefaultHeader;
            selectionColor = DefaultSelection;
            ResetFont();
        }

        protected override void SaveProperties(IDictionary<string, object> p)
        {
            p["AutoPaging"] = autoPaging;
            p["ShowIndex"] = showIndex;
            p["FontName"] = fontName;
            p["FontSize"] = fontSize;
            p["BackColor"] = backColor;
            p["ForeColor"] = foreColor;
            p["BorderColor"] = borderColor;
            p["HeaderBackColor"] = headerBackColor;
            p["SelectionColor"] = selectionColor;
        }

        protected override void LoadProperties(IDictionary<string, object> p)
        {
            autoPaging = Get(p, "AutoPaging", autoPaging);
            showIndex = Get(p, "ShowIndex", showIndex);
            fontName = Get(p, "FontName", fontName);
            fontSize = Get(p, "FontSize", fontSize);
            backColor = Get(p, "BackColor", backColor);
            foreColor = Get(p, "ForeColor", foreColor);
            borderColor = Get(p, "BorderColor", borderColor);
            headerBackColor = Get(p, "HeaderBackColor", headerBackColor);
            selectionColor = Get(p, "SelectionColor", selectionColor);
            ResetFont();
        }

        protected override Control CreateContent()
        {
            view = new GridView(this);
            view.Dock = DockStyle.Fill;
            view.SetData(frame);
            return view;
        }

        void ResetFont()
        {
            Font old = font;
            font = new Font(fontName, (float)Math.Max(1, fontSize), FontStyle.Regular, GraphicsUnit.Point);
            if (view != null) view.Font = font;
            if (old != null) old.Dispose();
        }

        void Restyle()
        {
            if (view != null) { view.Font = font; view.Relayout(); view.Invalidate(); }
            Changed();
        }

        internal Font GridFont { get { return font; } }
        internal Color Back { get { return ColorTranslator.FromOle(backColor); } }
        internal Color Fore { get { return ColorTranslator.FromOle(foreColor); } }
        internal Color Border { get { return ColorTranslator.FromOle(borderColor); } }
        internal Color Header { get { return ColorTranslator.FromOle(headerBackColor); } }
        internal Color SelectionBack { get { return ColorTranslator.FromOle(selectionColor); } }
        internal bool Paging { get { return autoPaging; } }
        internal bool Index { get { return showIndex; } }
        internal bool IsUserMode { get { return UserMode; } }
        internal bool Forward(ref Message m) { return ForwardKey(ref m); }
        internal void Focused(bool got) { ContentFocus(got); }
        internal void RaiseClick() { if (EventsEnabled && Click != null) Click(); }
        internal void RaiseDblClick() { if (EventsEnabled && DblClick != null) DblClick(); }
        internal void RaiseSelChange() { if (EventsEnabled && SelChange != null) SelChange(); }
        internal void RaisePageChange() { if (EventsEnabled && PageChange != null) PageChange(); }

        // Sem janela (designer): uma grade de exemplo
        protected override void Paint(Graphics g, Rectangle r, bool windowed)
        {
            using (SolidBrush b = new SolidBrush(Back)) g.FillRectangle(b, r);
            if (windowed) return; // a grade desenha o resto
            int rowH = font.Height + 8;
            using (SolidBrush h = new SolidBrush(Header)) g.FillRectangle(h, r.Left, r.Top, r.Width, rowH);
            using (Pen line = new Pen(Color.FromArgb(225, 225, 225)))
            {
                for (int y = r.Top + rowH; y < r.Bottom; y += rowH) g.DrawLine(line, r.Left, y, r.Right, y);
                int colW = Math.Max(60, r.Width / 3);
                for (int x = r.Left + colW; x < r.Right; x += colW) g.DrawLine(line, x, r.Top, x, r.Bottom);
            }
            TextRenderer.DrawText(g, "Sage DataGrid", font, new Rectangle(r.Left + 6, r.Top, r.Width - 12, rowH), Fore,
                TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            using (Pen pen = new Pen(Border)) g.DrawRectangle(pen, r.Left, r.Top, r.Width - 1, r.Height - 1);
        }

        // ------------------------------------------------------------------
        // _DataGrid
        // ------------------------------------------------------------------

        public DataFrame DataFrame
        {
            get { return frame; }
            set
            {
                if (value != null) CheckFormats(value, formats.Keys);
                frame = value;
                if (view != null) view.SetData(frame);
            }
        }

        // ColumnFormats: DictionaryS, Scripting.Dictionary ou Array("coluna", "formato", ...),
        // como o DictionaryS recebe; os formatos são os da função Format do VBA
        public object ColumnFormats
        {
            get
            {
                DictionaryS result = new DictionaryS();
                foreach (KeyValuePair<string, string> f in formats) result.LetValue(f.Key, f.Value);
                return result;
            }
            set { LetColumnFormats(value); }
        }

        public void LetColumnFormats(object Value)
        {
            Dictionary<string, string> parsed = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            object source = Interop.Unwrap(Value);
            if (source != null && !(source is DBNull) && !Interop.IsMissing(source))
                foreach (KeyValuePair<object, object> pair in Interop.PairsOf(source))
                    parsed[Interop.Text(pair.Key)] = Interop.Text(pair.Value);
            if (frame != null) CheckFormats(frame, parsed.Keys);
            formats.Clear();
            foreach (KeyValuePair<string, string> f in parsed) formats[f.Key] = f.Value;
            if (view != null) view.Reformat();
        }

        // Os nomes dos formatos têm que ser colunas do DataFrame
        static void CheckFormats(DataFrame df, IEnumerable<string> names)
        {
            List<string> columns = df.ColumnNames;
            foreach (string name in names)
            {
                bool found = false;
                foreach (string c in columns) if (string.Equals(c, name, StringComparison.OrdinalIgnoreCase)) { found = true; break; }
                if (!found)
                    throw Interop.Error(9, "KeyError: a coluna '" + name + "' de ColumnFormats não está no DataFrame (colunas: " +
                        string.Join(", ", columns.ToArray()) + ")");
            }
        }

        // O texto da célula: com o formato da coluna (o mesmo motor do Format do VBA), ou o padrão
        internal string Format(string column, object value, string type)
        {
            string format;
            if (value != null && formats.TryGetValue(column, out format) && format.Length > 0)
            {
                object input = value is long ? (object)(decimal)(long)value : value is ulong ? (object)(decimal)(ulong)value
                    : value is DateTimeOffset ? (object)((DateTimeOffset)value).DateTime : value;
                string text;
                try { if (VarFormat(ref input, format, 1, 1, 0, out text) == 0) return text; }
                catch (Exception) { }
            }
            return DataFrame.Display(value, type);
        }

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
        static extern int VarFormat([MarshalAs(UnmanagedType.Struct)] ref object input, string format, int firstDay, int firstWeek, int flags,
            [MarshalAs(UnmanagedType.BStr)] out string output);

        public bool AutoPaging { get { return autoPaging; } set { autoPaging = value; Restyle(); } }
        public bool ShowIndex { get { return showIndex; } set { showIndex = value; Restyle(); } }

        public int Page
        {
            get { return view != null ? view.Page : 1; }
            set { if (view != null) view.GoToPage(value); }
        }

        public int PageCount { get { return view != null ? view.PageCount : 1; } }
        public int RowsPerPage { get { return view != null ? view.RowsPerPage : 0; } }
        public long RowCount { get { return frame != null ? frame.Count : 0; } }

        public long SelectedRow
        {
            get { return view != null ? view.Selected : -1; }
            set { if (view != null) view.SelectRow(value, true); }
        }

        public object Cell(long Row, object Column)
        {
            if (frame == null) throw Interop.Error(91, "O DataGrid não tem DataFrame");
            long count = frame.Count;
            if (Row < 0) Row += count;
            if (Row < 0 || Row >= count) throw Interop.Error(9, "IndexError: linha " + Row + " fora do DataFrame (" + count + " linhas)");
            List<string> names = frame.ColumnNames;
            object c = Interop.Unwrap(Column);
            int index;
            if (c is string)
            {
                index = names.IndexOf((string)c);
                if (index < 0) throw Interop.Error(9, "KeyError: '" + c + "'");
            }
            else
            {
                index = Convert.ToInt32(c, CultureInfo.InvariantCulture);
                if (index < 0) index += names.Count;
                if (index < 0 || index >= names.Count) throw Interop.Error(9, "IndexError: coluna fora do DataFrame");
            }
            List<object[]> rows = frame.Rows(Row, 1);
            return rows.Count == 0 ? null : Interop.Wrap(rows[0][index]);
        }

        // A coluna da célula ativa ("" sem seleção)
        public string SelectedColumn { get { return view != null ? view.FocusColumn : ""; } }

        // Copia a seleção para a área de transferência, como o Ctrl+C
        public void Copy() { if (view != null) view.CopySelection(); }

        public void NextPage() { if (view != null) view.GoToPage(view.Page + 1); }
        public void PreviousPage() { if (view != null) view.GoToPage(view.Page - 1); }

        public string FontName
        {
            get { return fontName; }
            set { if (!string.IsNullOrEmpty(value)) { fontName = value; ResetFont(); Restyle(); } }
        }

        public double FontSize
        {
            get { return fontSize; }
            set
            {
                if (value <= 0 || value > 400) throw Interop.Error(380, "ValueError: FontSize deve ficar entre 1 e 400");
                fontSize = value; ResetFont(); Restyle();
            }
        }

        public int BackColor { get { return backColor; } set { backColor = value; Restyle(); } }
        public int ForeColor { get { return foreColor; } set { foreColor = value; Restyle(); } }
        public int BorderColor { get { return borderColor; } set { borderColor = value; Restyle(); } }
        public int HeaderBackColor { get { return headerBackColor; } set { headerBackColor = value; Restyle(); } }
        public int SelectionColor { get { return selectionColor; } set { selectionColor = value; Restyle(); } }
    }

    // A grade: desenho próprio, linhas lidas do DataFrame só quando aparecem
    sealed class GridView : Control
    {
        const int WM_KEYDOWN = 0x100, WM_CHAR = 0x102, WM_NCHITTEST = 0x84, HTTRANSPARENT = -1;
        static readonly Color NullBack = Color.FromArgb(75, 0, 130);
        // Célula NULL: "Vazio" em português, "Empty" (o vazio do VBA) nos outros idiomas
        static string NullText { get { return CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "pt" ? "Vazio" : "Empty"; } }
        static readonly Color Line = Color.FromArgb(225, 225, 225);

        readonly DataGrid owner;
        readonly HScrollBar hScroll = new HScrollBar();
        readonly VScrollBar vScroll = new VScrollBar();
        DataFrame frame;
        List<string> names = new List<string>(), types = new List<string>();
        long count;
        long top;                       // primeira linha à vista
        long anchorRow = -1, focusRow = -1; // seleção: da âncora até a célula ativa (focus)
        int anchorCol = -1, focusCol = -1;
        int visibleRows = 1, rowsPerPage = 1;
        bool pager;
        int[] widths = new int[0];
        int indexWidth;
        long cacheStart = -1;
        List<object[]> cache = new List<object[]>();
        Rectangle prevButton, nextButton;

        public GridView(DataGrid owner)
        {
            this.owner = owner;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            Font = owner.GridFont;
            hScroll.Visible = false;
            vScroll.Visible = false;
            hScroll.Scroll += delegate { Invalidate(); };
            vScroll.Scroll += delegate { top = vScroll.Value; Invalidate(); };
            Controls.Add(hScroll);
            Controls.Add(vScroll);
            GotFocus += delegate { owner.Focused(true); Invalidate(); };
            LostFocus += delegate { owner.Focused(false); Invalidate(); };
        }

        float DpiScale { get { using (Graphics g = CreateGraphics()) return g.DpiX / 96f; } }
        int RowHeight { get { return Font.Height + (int)Math.Round(8 * DpiScale); } }
        int HeaderHeight { get { return RowHeight + 2; } }

        public int RowsPerPage { get { return owner.Paging ? rowsPerPage : visibleRows; } }
        public int PageCount { get { return !owner.Paging || count == 0 ? 1 : (int)Math.Min(int.MaxValue, (count + rowsPerPage - 1) / rowsPerPage); } }
        public int Page { get { return !owner.Paging ? 1 : (int)(top / Math.Max(1, rowsPerPage)) + 1; } }
        public long Selected { get { return HasSelection ? focusRow : -1; } }

        public void SetData(DataFrame value)
        {
            frame = value;
            names = frame != null ? frame.ColumnNames : new List<string>();
            types = frame != null ? frame.ColumnTypes : new List<string>();
            count = frame != null ? frame.Count : 0;
            top = 0; anchorRow = focusRow = -1; anchorCol = focusCol = -1; cacheStart = -1; cache.Clear();
            hScroll.Value = 0;
            Relayout();
            MeasureColumns();
            Relayout();
            Invalidate();
        }

        // Quantas linhas cabem, se há páginas, e as barras de rolagem
        public void Relayout()
        {
            if (Width <= 0 || Height <= 0) return;
            int rowH = RowHeight, pagerH = rowH + (int)Math.Round(6 * DpiScale);
            for (int pass = 0; pass < 2; pass++)
            {
                int width = ClientSize.Width - (vScroll.Visible ? vScroll.Width : 0);
                hScroll.Visible = TotalWidth > width && names.Count > 0;
                int available = ClientSize.Height - HeaderHeight - (hScroll.Visible ? hScroll.Height : 0);
                if (owner.Paging)
                {
                    vScroll.Visible = false;
                    int fit = Math.Max(1, available / rowH);
                    pager = count > fit;
                    if (pager) fit = Math.Max(1, (available - pagerH) / rowH);
                    rowsPerPage = fit;
                    visibleRows = fit;
                    top = (top / rowsPerPage) * rowsPerPage; // começo de uma página
                }
                else
                {
                    pager = false;
                    visibleRows = Math.Max(1, available / rowH);
                    vScroll.Visible = count > visibleRows;
                    top = Math.Max(0, Math.Min(top, count - visibleRows));
                }
            }
            int bottom = ClientSize.Height - (pager ? pagerH : 0);
            if (hScroll.Visible)
            {
                int width = ClientSize.Width - (vScroll.Visible ? vScroll.Width : 0);
                hScroll.SetBounds(0, bottom - hScroll.Height, width, hScroll.Height);
                hScroll.Maximum = Math.Max(0, TotalWidth - 1);
                hScroll.LargeChange = Math.Max(1, width);
                hScroll.SmallChange = 40;
                if (hScroll.Value > Math.Max(0, TotalWidth - width)) hScroll.Value = Math.Max(0, TotalWidth - width);
            }
            else hScroll.Value = 0;
            if (vScroll.Visible)
            {
                vScroll.SetBounds(ClientSize.Width - vScroll.Width, HeaderHeight, vScroll.Width,
                    bottom - HeaderHeight - (hScroll.Visible ? hScroll.Height : 0));
                vScroll.Maximum = (int)Math.Min(int.MaxValue - 1, Math.Max(0, count - 1));
                vScroll.LargeChange = visibleRows;
                vScroll.SmallChange = 1;
                vScroll.Value = (int)Math.Min(top, vScroll.Maximum);
            }
        }

        int TotalWidth
        {
            get
            {
                int w = owner.Index ? indexWidth : 0;
                foreach (int c in widths) w += c;
                return w;
            }
        }

        // Larguras: o título da coluna e os valores das primeiras linhas, entre 50 e 320 px
        void MeasureColumns()
        {
            float scale = DpiScale;
            int pad = (int)Math.Round(16 * scale);
            widths = new int[names.Count];
            List<object[]> sample = Rows(top, Math.Max(visibleRows, 50));
            using (Graphics g = CreateGraphics())
            using (Font bold = new Font(Font, FontStyle.Bold))
            {
                for (int c = 0; c < names.Count; c++)
                {
                    int w = TextRenderer.MeasureText(g, names[c], bold).Width;
                    foreach (object[] row in sample)
                        w = Math.Max(w, TextRenderer.MeasureText(g, CellText(row[c], c), Font).Width);
                    widths[c] = Math.Max((int)(50 * scale), Math.Min((int)(320 * scale), w + pad));
                }
                indexWidth = TextRenderer.MeasureText(g, Math.Max(0, count - 1).ToString("N0", CultureInfo.CurrentCulture), Font).Width + pad;
            }
        }

        string CellText(object value, int column)
        {
            return value == null ? NullText : owner.Format(names[column], value, types[column]);
        }

        // Os formatos das colunas mudaram: larguras e desenho
        public void Reformat()
        {
            MeasureColumns();
            Relayout();
            Invalidate();
        }

        static bool IsNumber(string type)
        {
            return type.EndsWith("INT") || type == "INTEGER" || type == "DOUBLE" || type == "FLOAT" || type.StartsWith("DECIMAL");
        }

        // As linhas de [start, start + n), do cache ou do DuckDB
        List<object[]> Rows(long start, int n)
        {
            if (frame == null || n <= 0) return new List<object[]>();
            if (cacheStart < 0 || start < cacheStart || start + n > cacheStart + cache.Count && cacheStart + cache.Count < count)
            {
                cacheStart = start;
                cache = frame.Rows(start, Math.Max(n, 200));
            }
            List<object[]> result = new List<object[]>();
            for (long i = start; i < start + n && i - cacheStart < cache.Count; i++) result.Add(cache[(int)(i - cacheStart)]);
            return result;
        }

        public void GoToPage(int page)
        {
            if (!owner.Paging || frame == null) return;
            int target = Math.Max(1, Math.Min(PageCount, page));
            if (target == Page) return;
            top = (long)(target - 1) * rowsPerPage;
            Invalidate();
            owner.RaisePageChange();
        }

        // ------------------------------------------------------------------
        // Seleção, como no st.dataframe do Streamlit: uma célula ativa (focus) e um retângulo
        // dela até a âncora; linha inteira (clique no índice) ou coluna inteira (no cabeçalho)
        // ------------------------------------------------------------------

        long R1 { get { return Math.Min(anchorRow, focusRow); } }
        long R2 { get { return Math.Max(anchorRow, focusRow); } }
        int C1 { get { return Math.Min(anchorCol, focusCol); } }
        int C2 { get { return Math.Max(anchorCol, focusCol); } }
        bool HasSelection { get { return focusRow >= 0 && focusCol >= 0 && count > 0 && names.Count > 0; } }

        bool IsSelected(long row, int col)
        {
            return HasSelection && row >= R1 && row <= R2 && col >= C1 && col <= C2;
        }

        public string FocusColumn { get { return HasSelection ? names[focusCol] : ""; } }

        // Seleciona de (anchorRow, anchorCol) até (row, col); extend = mantém a âncora (Shift)
        void SelectCells(long row, int col, bool extend, bool show)
        {
            if (frame == null || count == 0 || names.Count == 0) return;
            row = Math.Max(0, Math.Min(count - 1, row));
            col = Math.Max(0, Math.Min(names.Count - 1, col));
            long before = R1 * 31 + R2 * 17 + C1 * 7 + C2 + focusRow * 13 + focusCol;
            if (!extend || anchorRow < 0) { anchorRow = row; anchorCol = col; }
            focusRow = row; focusCol = col;
            if (show) Reveal(row, col);
            Invalidate();
            if (R1 * 31 + R2 * 17 + C1 * 7 + C2 + focusRow * 13 + focusCol != before) owner.RaiseSelChange();
        }

        void SelectRange(long r1, int c1, long r2, int c2)
        {
            anchorRow = r1; anchorCol = c1;
            SelectCells(r2, c2, true, false);
        }

        // Compatível com SelectedRow: a linha inteira
        public void SelectRow(long row, bool show)
        {
            if (frame == null || count == 0) return;
            if (row < 0) { anchorRow = focusRow = -1; anchorCol = focusCol = -1; Invalidate(); owner.RaiseSelChange(); return; }
            row = Math.Min(count - 1, row);
            anchorRow = row; anchorCol = 0;
            SelectCells(row, names.Count - 1, true, show);
            focusCol = 0;
            Invalidate();
        }

        // Rola/pagina até a célula aparecer
        void Reveal(long row, int col)
        {
            int before = Page;
            if (owner.Paging) top = row / rowsPerPage * rowsPerPage;
            else if (row < top) top = row;
            else if (row >= top + visibleRows) top = row - visibleRows + 1;
            if (vScroll.Visible) vScroll.Value = (int)Math.Min(top, vScroll.Maximum);
            if (owner.Paging && Page != before) owner.RaisePageChange();
            if (hScroll.Visible && col >= 0)
            {
                int left = owner.Index ? indexWidth : 0;
                for (int c = 0; c < col; c++) left += widths[c];
                int right = left + widths[col];
                int view = ClientSize.Width - (vScroll.Visible ? vScroll.Width : 0);
                int fixedLeft = owner.Index ? indexWidth : 0;
                if (left - hScroll.Value < fixedLeft) hScroll.Value = Math.Max(0, left - fixedLeft);
                else if (right - hScroll.Value > view) hScroll.Value = Math.Min(Math.Max(0, hScroll.Maximum - hScroll.LargeChange + 1), right - view);
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            int before = Page;
            Relayout();
            Invalidate();
            if (owner.Paging && frame != null && Page != before) owner.RaisePageChange();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Rectangle client = ClientRectangle;
            using (SolidBrush b = new SolidBrush(owner.Back)) g.FillRectangle(b, client);
            int rowH = RowHeight, headerH = HeaderHeight;
            int x0 = -hScroll.Value;
            int width = client.Width - (vScroll.Visible ? vScroll.Width : 0);
            int pagerH = pager ? rowH + (int)Math.Round(6 * DpiScale) : 0;
            int bodyBottom = client.Height - pagerH - (hScroll.Visible ? hScroll.Height : 0);
            Color fore = owner.Fore, muted = SystemColors.GrayText;
            // Sem o foco, a mesma cor acinzentada
            Color chosen = owner.SelectionBack;
            int gray = (chosen.R + chosen.G + chosen.B) / 3;
            Color selection = Focused ? chosen : Color.FromArgb((chosen.R + gray * 2) / 3, (chosen.G + gray * 2) / 3, (chosen.B + gray * 2) / 3);
            Color headerSelected = Color.FromArgb(Math.Max(0, owner.Header.R - 22), Math.Max(0, owner.Header.G - 22), Math.Max(0, owner.Header.B - 22));
            TextFormatFlags flags = TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine;
            int pad = (int)Math.Round(6 * DpiScale);
            Rectangle active = Rectangle.Empty;

            g.SetClip(new Rectangle(0, 0, width, bodyBottom));
            // Cabeçalho (as colunas da seleção, destacadas)
            using (SolidBrush h = new SolidBrush(owner.Header)) g.FillRectangle(h, 0, 0, width, headerH);
            using (Font bold = new Font(Font, FontStyle.Bold))
            using (SolidBrush hs = new SolidBrush(headerSelected))
            {
                int x = x0;
                if (owner.Index) x += indexWidth;
                for (int c = 0; c < names.Count; c++)
                {
                    if (HasSelection && c >= C1 && c <= C2) g.FillRectangle(hs, x, 0, widths[c], headerH);
                    TextRenderer.DrawText(g, names[c], bold, new Rectangle(x + pad, 0, widths[c] - 2 * pad, headerH), fore,
                        flags | (IsNumber(types[c]) ? TextFormatFlags.Right : TextFormatFlags.Left));
                    x += widths[c];
                }
            }

            // Linhas
            List<object[]> rows = Rows(top, visibleRows);
            using (Pen line = new Pen(Line))
            using (SolidBrush sel = new SolidBrush(selection))
            using (SolidBrush hs = new SolidBrush(headerSelected))
            using (SolidBrush nul = new SolidBrush(NullBack))
            using (SolidBrush nulSel = new SolidBrush(Color.FromArgb(120, 60, 170)))
            {
                for (int r = 0; r < rows.Count; r++)
                {
                    long index = top + r;
                    int y = headerH + r * rowH;
                    int x = x0;
                    if (owner.Index)
                    {
                        if (HasSelection && index >= R1 && index <= R2) g.FillRectangle(hs, x, y, indexWidth, rowH);
                        TextRenderer.DrawText(g, index.ToString("N0", CultureInfo.CurrentCulture), Font,
                            new Rectangle(x + pad, y, indexWidth - 2 * pad, rowH), muted, flags | TextFormatFlags.Right);
                        x += indexWidth;
                    }
                    for (int c = 0; c < names.Count; c++)
                    {
                        object value = rows[r][c];
                        Rectangle cell = new Rectangle(x, y, widths[c], rowH);
                        bool selected = IsSelected(index, c);
                        if (value == null)
                        {
                            g.FillRectangle(selected ? nulSel : nul, cell.X + 1, cell.Y + 1, cell.Width - 1, cell.Height - 1);
                            TextRenderer.DrawText(g, NullText, Font, new Rectangle(cell.X + pad, y, cell.Width - 2 * pad, rowH), Color.White, flags);
                        }
                        else
                        {
                            if (selected) g.FillRectangle(sel, cell.X, cell.Y, cell.Width, cell.Height - 1);
                            TextRenderer.DrawText(g, CellText(value, c), Font, new Rectangle(cell.X + pad, y, cell.Width - 2 * pad, rowH), fore,
                                flags | (IsNumber(types[c]) ? TextFormatFlags.Right : TextFormatFlags.Left));
                        }
                        if (index == focusRow && c == focusCol) active = cell;
                        x += widths[c];
                    }
                    g.DrawLine(line, 0, y + rowH - 1, width, y + rowH - 1);
                }
                g.DrawLine(line, 0, headerH - 1, width, headerH - 1);
                // Divisórias das colunas
                int cx = x0 + (owner.Index ? indexWidth : 0);
                int linesBottom = Math.Min(bodyBottom, headerH + rows.Count * rowH);
                if (owner.Index) g.DrawLine(line, cx - 1, 0, cx - 1, linesBottom);
                foreach (int w in widths) { cx += w; g.DrawLine(line, cx - 1, 0, cx - 1, linesBottom); }
            }
            // Célula ativa: contorno azul
            if (!active.IsEmpty)
            {
                int t = Math.Max(2, (int)Math.Round(2 * DpiScale));
                using (Pen pen = new Pen(Focused ? Color.FromArgb(0, 120, 212) : Color.FromArgb(150, 150, 150), t))
                    g.DrawRectangle(pen, active.X + t / 2, active.Y + t / 2 - 1, active.Width - t, active.Height - t + 1);
            }
            g.ResetClip();

            // Paginação
            if (pager)
            {
                int y = client.Height - pagerH;
                using (SolidBrush h = new SolidBrush(owner.Header)) g.FillRectangle(h, 0, y, client.Width, pagerH);
                using (Pen line = new Pen(Line)) g.DrawLine(line, 0, y, client.Width, y);
                // À direita: "linhas 1–8 de 1.998" e depois "◀ Página 1 de N ▶"
                bool pt = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "pt";
                CultureInfo ci = CultureInfo.CurrentCulture;
                string pageText = (pt ? "Página " : "Page ") + Page.ToString("N0", ci) + (pt ? " de " : " of ") + PageCount.ToString("N0", ci);
                long first = top + 1, last = Math.Min(count, top + rowsPerPage);
                string rowsText = (pt ? "linhas " : "rows ") + first.ToString("N0", ci) + "–" + last.ToString("N0", ci) +
                    (pt ? " de " : " of ") + count.ToString("N0", ci);
                int button = pagerH;
                int pageWidth = TextRenderer.MeasureText(g, pageText, Font, Size.Empty, TextFormatFlags.NoPadding).Width + 2 * pad;
                nextButton = new Rectangle(client.Width - pad - button, y, button, pagerH);
                Rectangle pageRect = new Rectangle(nextButton.Left - pageWidth, y, pageWidth, pagerH);
                prevButton = new Rectangle(pageRect.Left - button, y, button, pagerH);
                TextRenderer.DrawText(g, "◀", Font, prevButton, Page > 1 ? fore : muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                TextRenderer.DrawText(g, pageText, Font, pageRect, fore, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
                TextRenderer.DrawText(g, "▶", Font, nextButton, Page < PageCount ? fore : muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                TextRenderer.DrawText(g, rowsText, Font, new Rectangle(pad, y, prevButton.Left - 2 * pad, pagerH), muted, flags | TextFormatFlags.Right);
            }
            else prevButton = nextButton = Rectangle.Empty;

            using (Pen pen = new Pen(Focused ? Color.FromArgb(0, 120, 212) : owner.Border))
                g.DrawRectangle(pen, 0, 0, client.Width - 1, client.Height - 1);
        }

        // O que está sob o ponto: área (cabeçalho, índice, célula), linha e coluna
        enum Area { None, Header, Index, Cell }

        Area HitTest(Point p, out long row, out int col)
        {
            row = -1; col = -1;
            int rowH = RowHeight;
            int bodyBottom = ClientSize.Height - (pager ? rowH + (int)Math.Round(6 * DpiScale) : 0) - (hScroll.Visible ? hScroll.Height : 0);
            int width = ClientSize.Width - (vScroll.Visible ? vScroll.Width : 0);
            if (p.X < 0 || p.X >= width || p.Y < 0 || p.Y >= bodyBottom || names.Count == 0) return Area.None;
            int x = -hScroll.Value;
            bool inIndex = false;
            if (owner.Index)
            {
                if (p.X >= x && p.X < x + indexWidth) inIndex = true;
                x += indexWidth;
            }
            if (!inIndex)
                for (int c = 0; c < names.Count; c++)
                {
                    if (p.X >= x && p.X < x + widths[c]) { col = c; break; }
                    x += widths[c];
                }
            if (p.Y < HeaderHeight) return col >= 0 ? Area.Header : Area.None;
            long r = top + (p.Y - HeaderHeight) / rowH;
            if (r >= Math.Min(count, top + visibleRows)) return Area.None;
            row = r;
            if (inIndex) return Area.Index;
            return col >= 0 ? Area.Cell : Area.None;
        }

        bool dragging;

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (!owner.IsUserMode) return;
            Focus();
            if (e.Button != MouseButtons.Left) return;
            if (prevButton.Contains(e.Location)) { GoToPage(Page - 1); return; }
            if (nextButton.Contains(e.Location)) { GoToPage(Page + 1); return; }
            bool shift = (ModifierKeys & Keys.Shift) != 0;
            long row; int col;
            switch (HitTest(e.Location, out row, out col))
            {
                case Area.Cell:
                    SelectCells(row, col, shift, false);
                    dragging = true;
                    Capture = true;
                    break;
                case Area.Header: // a coluna inteira
                    if (shift && anchorRow >= 0) SelectRange(0, anchorCol, count - 1, col);
                    else SelectRange(0, col, count - 1, col);
                    focusRow = top;
                    break;
                case Area.Index: // a linha inteira
                    if (shift && anchorRow >= 0) SelectRange(anchorRow, 0, row, names.Count - 1);
                    else SelectRange(row, 0, row, names.Count - 1);
                    focusCol = 0;
                    break;
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (!dragging || (e.Button & MouseButtons.Left) == 0) return;
            long row; int col;
            Point p = new Point(Math.Max(0, Math.Min(ClientSize.Width - 1, e.X)), Math.Max(HeaderHeight, e.Y));
            if (HitTest(p, out row, out col) == Area.Cell && (row != focusRow || col != focusCol)) SelectCells(row, col, true, false);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            dragging = false;
            Capture = false;
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);
            long row; int col;
            if (owner.IsUserMode && e.Button == MouseButtons.Left && HitTest(e.Location, out row, out col) != Area.None) owner.RaiseClick();
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            long row; int col;
            if (owner.IsUserMode && e.Button == MouseButtons.Left && HitTest(e.Location, out row, out col) == Area.Cell) owner.RaiseDblClick();
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            if (frame == null) return;
            if (owner.Paging) GoToPage(Page + (e.Delta < 0 ? 1 : -1));
            else
            {
                top = Math.Max(0, Math.Min(Math.Max(0, count - visibleRows), top + (e.Delta < 0 ? 3 : -3)));
                if (vScroll.Visible) vScroll.Value = (int)Math.Min(top, vScroll.Maximum);
                Invalidate();
            }
        }

        // Sem o laço de mensagens do WinForms: teclas tratadas aqui; Tab, Enter e Esc vão para o form
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_NCHITTEST && !owner.IsUserMode) { m.Result = (IntPtr)HTTRANSPARENT; return; }
            if (m.Msg == WM_KEYDOWN && frame != null && count > 0 && names.Count > 0)
            {
                Keys key = (Keys)(int)m.WParam;
                bool ctrl = (ModifierKeys & Keys.Control) != 0, shift = (ModifierKeys & Keys.Shift) != 0;
                long row = Math.Max(0, focusRow);
                int col = Math.Max(0, focusCol);
                int page = owner.Paging ? rowsPerPage : visibleRows;
                switch (key)
                {
                    case Keys.Tab:
                    case Keys.Escape:
                    case Keys.Return:
                        owner.Forward(ref m);
                        return;
                    case Keys.Down: SelectCells(ctrl ? count - 1 : row + 1, col, shift, true); return;
                    case Keys.Up: SelectCells(ctrl ? 0 : row - 1, col, shift, true); return;
                    case Keys.Right: SelectCells(row, ctrl ? names.Count - 1 : col + 1, shift, true); return;
                    case Keys.Left: SelectCells(row, ctrl ? 0 : col - 1, shift, true); return;
                    case Keys.PageDown: SelectCells(row + page, col, shift, true); return;
                    case Keys.PageUp: SelectCells(row - page, col, shift, true); return;
                    case Keys.Home: SelectCells(ctrl ? 0 : row, 0, shift, true); return;
                    case Keys.End: SelectCells(ctrl ? count - 1 : row, names.Count - 1, shift, true); return;
                    case Keys.A:
                        if (ctrl) { SelectRange(0, 0, count - 1, names.Count - 1); return; }
                        break;
                    case Keys.C:
                        if (ctrl) { CopySelection(); return; }
                        break;
                }
            }
            else if (m.Msg == WM_KEYDOWN && ((Keys)(int)m.WParam == Keys.Tab || (Keys)(int)m.WParam == Keys.Escape || (Keys)(int)m.WParam == Keys.Return))
            {
                owner.Forward(ref m);
                return;
            }
            if (m.Msg == WM_CHAR) return; // sem bipe
            base.WndProc(ref m);
        }

        // Ctrl+C: a seleção separada por tabulações (cola em planilhas), com os formatos das
        // colunas; NULL vira vazio. Lida do DuckDB em blocos.
        public void CopySelection()
        {
            if (!HasSelection) return;
            long r1 = R1, r2 = R2;
            int c1 = C1, c2 = C2;
            StringBuilder s = new StringBuilder();
            for (long start = r1; start <= r2; start += 5000)
            {
                int n = (int)Math.Min(5000, r2 - start + 1);
                foreach (object[] row in frame.Rows(start, n))
                {
                    for (int c = c1; c <= c2; c++)
                    {
                        if (c > c1) s.Append('\t');
                        if (row[c] != null) s.Append(CellText(row[c], c).Replace("\t", " ").Replace("\r", " ").Replace("\n", " "));
                    }
                    s.Append("\r\n");
                }
            }
            try { Clipboard.SetText(s.ToString()); } catch (Exception) { }
        }
    }
}
