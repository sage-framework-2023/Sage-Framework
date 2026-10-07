using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageEditor
{
    // Aba "DataFrame Results" da janela Terminal, como o "Query Results" de um editor de SQL:
    // o df.Show do VBA mostra o DataFrame aqui. A grade é virtual e busca só as linhas visíveis,
    // em páginas de mil, então abre na hora mesmo com milhões de linhas. Embaixo, à esquerda,
    // o título e as colunas; à direita, a quantidade de linhas.
    //
    // O SageTypes não referencia o SageEditor: ele acha a função Show no AppDomain
    // (TerminalWindow publica em "Sage.DataFrameResults") e passa só tipos do .NET.
    sealed class ResultsView : UserControl
    {
        const int PageSize = 1000;
        readonly ResultsGrid grid = new ResultsGrid();
        readonly Panel status = new Panel();
        readonly Label summary = new Label(), count = new Label(), empty = new Label();
        readonly Dictionary<long, string[][]> pages = new Dictionary<long, string[][]>();
        readonly Queue<long> order = new Queue<long>();
        Func<long, int, string[][]> fetch;
        string[] types = new string[0];
        bool failed;
        Theme theme;

        public ResultsView()
        {
            Font = SystemFonts.MessageBoxFont;

            grid.Dock = DockStyle.Fill;
            grid.VirtualMode = true;
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.AllowUserToOrderColumns = false;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText;
            grid.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.BorderStyle = BorderStyle.None;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            grid.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            grid.EnableHeadersVisualStyles = false;
            grid.DefaultCellStyle.Font = new Font("Consolas", 9.75f);
            grid.ColumnHeadersDefaultCellStyle.Font = Font;
            grid.RowTemplate.Height = grid.DefaultCellStyle.Font.Height + 6;
            grid.ColumnHeadersHeight = Font.Height + 10;
            grid.CellValueNeeded += CellValueNeeded;
            grid.CellFormatting += FormatNull;
            grid.CellPainting += PaintRowNumber;
            grid.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if ((e.Control && e.KeyCode == Keys.C) || (e.Control && e.KeyCode == Keys.Insert)) { Copy(); e.Handled = true; e.SuppressKeyPress = true; }
            };
            grid.Visible = false;

            empty.Dock = DockStyle.Fill;
            empty.TextAlign = ContentAlignment.MiddleCenter;
            empty.Text = Strings.ResultsEmpty;

            status.Dock = DockStyle.Bottom;
            status.Height = Font.Height + 8;
            status.Padding = new Padding(8, 0, 10, 0);
            summary.Dock = DockStyle.Fill;
            summary.TextAlign = ContentAlignment.MiddleLeft;
            summary.AutoEllipsis = true;
            count.Dock = DockStyle.Right;
            count.AutoSize = true;
            count.TextAlign = ContentAlignment.MiddleRight;
            count.Padding = new Padding(0, 4, 0, 0);
            status.Controls.Add(summary);
            status.Controls.Add(count);
            status.Paint += delegate(object s, PaintEventArgs e)
            {
                if (theme != null) using (Pen line = new Pen(theme.Border)) e.Graphics.DrawLine(line, 0, 0, status.Width, 0);
            };

            Controls.Add(grid);
            Controls.Add(empty);
            Controls.Add(status);
        }

        // Chamada pelo df.Show (thread do Excel)
        public void ShowFrame(string title, string[] columns, string[] columnTypes, long rows, Func<long, int, string[][]> rowFetch)
        {
            grid.RowCount = 0;
            grid.Columns.Clear();
            pages.Clear();
            order.Clear();
            failed = false;
            fetch = rowFetch;
            types = columnTypes;

            for (int c = 0; c < columns.Length; c++)
            {
                DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                column.HeaderText = columns[c];
                column.ToolTipText = columns[c] + " (" + columnTypes[c] + ")";
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                if (IsNumber(columnTypes[c])) column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                grid.Columns.Add(column);
            }
            int digits = Math.Max(3, (rows - 1).ToString("N0", CultureInfo.CurrentCulture).Length);
            grid.RowHeadersWidth = TextRenderer.MeasureText(new string('8', digits), grid.DefaultCellStyle.Font).Width + 16;
            grid.RowCount = (int)Math.Min(rows, int.MaxValue);
            FitColumns(columns);

            // À esquerda, só o título do df.Show (se houver); à direita, colunas e linhas
            summary.Text = title;
            count.Text = string.Format(CultureInfo.CurrentCulture, columns.Length == 1 ? Strings.ResultsColumn : Strings.ResultsColumns, columns.Length) +
                "   ·   " + string.Format(CultureInfo.CurrentCulture, rows == 1 ? Strings.ResultsRow : Strings.ResultsRows, rows);
            empty.Visible = false;
            grid.Visible = true;
            if (grid.RowCount > 0 && grid.ColumnCount > 0) grid.CurrentCell = grid[0, 0];
        }

        public void FocusGrid()
        {
            if (grid.Visible) grid.Focus();
        }

        // Largura pelo cabeçalho, pelas primeiras e pelas últimas linhas (os valores costumam
        // crescer até o fim, como um id); até 400 px
        void FitColumns(string[] columns)
        {
            List<string[]> sample = new List<string[]>();
            string[][] first = Page(0);
            if (first != null) sample.AddRange(first.Take(200));
            long lastPage = (grid.RowCount - 1) / PageSize;
            if (lastPage > 0)
            {
                string[][] last = Page(lastPage);
                if (last != null) sample.AddRange(last.Skip(Math.Max(0, last.Length - 200)));
            }
            for (int c = 0; c < columns.Length; c++)
            {
                int width = TextRenderer.MeasureText(columns[c], grid.ColumnHeadersDefaultCellStyle.Font).Width;
                foreach (string[] row in sample)
                    if (c < row.Length) width = Math.Max(width, TextRenderer.MeasureText(row[c] ?? "", grid.DefaultCellStyle.Font).Width);
                grid.Columns[c].Width = Math.Max(60, Math.Min(400, width + 20));
            }
        }

        string[][] Page(long page)
        {
            string[][] rows;
            if (pages.TryGetValue(page, out rows)) return rows;
            if (fetch == null || failed) return null;
            try { rows = fetch(page * PageSize, PageSize); }
            catch (Exception ex)
            {
                // A tabela deixou de existir (ex.: o DuckDB foi fechado): para de buscar
                failed = true;
                Log.Error(ex);
                return null;
            }
            pages[page] = rows;
            order.Enqueue(page);
            if (order.Count > 20) pages.Remove(order.Dequeue()); // até 20 mil linhas em memória
            return rows;
        }

        void CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            string[][] rows = Page(e.RowIndex / PageSize);
            int i = e.RowIndex % PageSize;
            if (rows != null && i < rows.Length && e.ColumnIndex < rows[i].Length) e.Value = rows[i][e.ColumnIndex] ?? (object)Null;
        }

        // Célula NULL: "Empty" (o vazio do VBA) num fundo roxo, como o NULL do Azure Data Studio
        static readonly object Null = new NullCell();
        sealed class NullCell { public override string ToString() { return "Empty"; } }
        Color nullBack, nullFore;

        void FormatNull(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value != Null) return;
            e.Value = "Empty";
            e.CellStyle.BackColor = nullBack;
            e.CellStyle.ForeColor = nullFore;
            e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            e.FormattingApplied = true;
        }

        // Ctrl+C: as células selecionadas, separadas por tabulação (como o Excel cola), com o
        // NULL em branco (e não o texto "Empty")
        void Copy()
        {
            if (grid.SelectedCells.Count == 0) return;
            int top = int.MaxValue, bottom = -1, left = int.MaxValue, right = -1;
            HashSet<long> selected = new HashSet<long>();
            foreach (DataGridViewCell cell in grid.SelectedCells)
            {
                top = Math.Min(top, cell.RowIndex); bottom = Math.Max(bottom, cell.RowIndex);
                left = Math.Min(left, cell.ColumnIndex); right = Math.Max(right, cell.ColumnIndex);
                selected.Add((long)cell.RowIndex << 32 | (uint)cell.ColumnIndex);
            }
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            for (int r = top; r <= bottom; r++)
            {
                for (int c = left; c <= right; c++)
                {
                    if (c > left) sb.Append('\t');
                    if (!selected.Contains((long)r << 32 | (uint)c)) continue;
                    string[][] rows = Page(r / PageSize);
                    string value = rows != null && r % PageSize < rows.Length ? rows[r % PageSize][c] : null;
                    if (value != null) sb.Append(value);
                }
                sb.Append("\r\n");
            }
            try { Clipboard.SetText(sb.ToString()); }
            catch (Exception ex) { Log.Error(ex); }
        }

        // Cabeçalho da linha: só o número (base 0, como o índice do pandas), sem a seta da
        // linha atual da grade
        void PaintRowNumber(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex >= 0 || e.RowIndex < 0) return;
            e.PaintBackground(e.ClipBounds, false);
            string label = e.RowIndex.ToString("N0", CultureInfo.CurrentCulture);
            Color color = theme != null ? theme.Muted : SystemColors.GrayText;
            Rectangle r = e.CellBounds;
            TextRenderer.DrawText(e.Graphics, label, grid.DefaultCellStyle.Font,
                new Rectangle(r.Left, r.Top, r.Width - 8, r.Height), color, TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
            e.Handled = true;
        }

        static bool IsNumber(string type)
        {
            return type.EndsWith("INT") || type == "INTEGER" || type == "DOUBLE" || type == "FLOAT" || type.StartsWith("DECIMAL");
        }

        public void ApplyTheme(Theme t)
        {
            theme = t;
            bool dark = t.Background.GetBrightness() < 0.5f;
            Color selection = Blend(t.Background, t.Accent, dark ? 0.45 : 0.25);
            // NULL: roxo com texto branco em qualquer tema
            nullBack = Color.FromArgb(75, 0, 130);
            nullFore = Color.White;
            BackColor = t.Background;
            grid.BackgroundColor = t.Background;
            grid.CornerColor = t.Background;
            grid.GridColor = t.Border;
            grid.DefaultCellStyle.BackColor = t.Background;
            grid.DefaultCellStyle.ForeColor = t.Foreground;
            grid.DefaultCellStyle.SelectionBackColor = selection;
            grid.DefaultCellStyle.SelectionForeColor = t.Foreground;
            foreach (DataGridViewCellStyle header in new[] { grid.ColumnHeadersDefaultCellStyle, grid.RowHeadersDefaultCellStyle })
            {
                header.BackColor = t.Sidebar;
                header.ForeColor = t.Foreground;
                header.SelectionBackColor = t.Sidebar;
                header.SelectionForeColor = t.Foreground;
            }
            empty.BackColor = t.Background;
            empty.ForeColor = t.Muted;
            status.BackColor = t.Sidebar;
            summary.ForeColor = t.Muted;
            count.ForeColor = t.Foreground;
            // Barras de rolagem da grade no tema escuro (Windows 10 e 11)
            string name = dark ? "DarkMode_Explorer" : "Explorer";
            foreach (Control bar in grid.Controls)
            {
                if (bar.IsHandleCreated) SetWindowTheme(bar.Handle, name, null);
                else { Control c = bar; c.HandleCreated += delegate { SetWindowTheme(c.Handle, name, null); }; }
            }
            Invalidate(true);
        }

        static Color Blend(Color a, Color b, double k)
        {
            return Color.FromArgb((int)(a.R + (b.R - a.R) * k), (int)(a.G + (b.G - a.G) * k), (int)(a.B + (b.B - a.B) * k));
        }

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        static extern int SetWindowTheme(IntPtr hwnd, string appName, string idList);

        // Grade que fica com as setas, PgUp/PgDn, Home/End e Ctrl+C (o VBE as usaria)
        sealed class ResultsGrid : DataGridView
        {
            public ResultsGrid() { DoubleBuffered = true; }

            // Canto entre as duas barras de rolagem: a grade o pinta na cor do sistema (branco)
            public Color CornerColor = SystemColors.Control;

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                if (VerticalScrollBar.Visible && HorizontalScrollBar.Visible)
                    using (SolidBrush b = new SolidBrush(CornerColor))
                        e.Graphics.FillRectangle(b, VerticalScrollBar.Left, HorizontalScrollBar.Top, VerticalScrollBar.Width, HorizontalScrollBar.Height);
            }

            protected override void WndProc(ref Message m)
            {
                const int WM_GETDLGCODE = 0x0087, DLGC_WANTALLKEYS = 0x0004, DLGC_WANTCHARS = 0x0080, DLGC_WANTARROWS = 0x0001;
                if (m.Msg == WM_GETDLGCODE)
                {
                    m.Result = (IntPtr)(DLGC_WANTALLKEYS | DLGC_WANTCHARS | DLGC_WANTARROWS);
                    return;
                }
                base.WndProc(ref m);
            }
        }
    }
}
