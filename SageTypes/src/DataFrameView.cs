using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace SageTypes
{
    // Janela do DataFrame.Show: grade em modo virtual, que busca só as linhas visíveis
    // (em páginas de mil), então abre na hora mesmo com dezenas de milhões de linhas.
    // Modal, na thread do Excel (onde o DuckDB é usado), como um UserForm.Show.
    sealed class DataFrameView : Form
    {
        const int PageSize = 1000;
        readonly DataFrame frame;
        readonly List<string> types;
        readonly DataGridView grid = new DataGridView();
        readonly Dictionary<long, List<object[]>> pages = new Dictionary<long, List<object[]>>();
        readonly Queue<long> order = new Queue<long>();

        public static void ShowDialog(DataFrame frame, string title)
        {
            using (DataFrameView view = new DataFrameView(frame, title)) view.ShowDialog();
        }

        // A aba "DataFrame Results" da janela Terminal do SageEditor. Os dois rodam no mesmo
        // processo (e AppDomain) do Excel, mas um não referencia o outro: o SageEditor publica
        // no AppDomain uma função só com tipos do .NET, que recebe o título, os nomes e os tipos
        // das colunas, a quantidade de linhas e uma função que busca um trecho de linhas já
        // formatadas (null: a célula é NULL, e a aba a mostra como Empty, destacada). A função
        // guarda o DataFrame, e com ele a tabela no DuckDB, enquanto a aba o mostrar.
        public const string EditorSlot = "Sage.DataFrameResults";

        public static bool ShowInEditor(DataFrame frame, string title)
        {
            Func<string, string[], string[], long, Func<long, int, string[][]>, bool> show =
                AppDomain.CurrentDomain.GetData(EditorSlot) as Func<string, string[], string[], long, Func<long, int, string[][]>, bool>;
            if (show == null) return false;
            List<string> types = frame.ColumnTypes;
            Func<long, int, string[][]> fetch = delegate(long start, int count)
            {
                List<object[]> rows = frame.Rows(start, count);
                string[][] text = new string[rows.Count][];
                for (int r = 0; r < rows.Count; r++)
                {
                    text[r] = new string[types.Count];
                    for (int c = 0; c < types.Count; c++)
                        text[r][c] = rows[r][c] == null ? null : DataFrame.Display(rows[r][c], types[c]);
                }
                return text;
            };
            try { return show(title, frame.ColumnNames.ToArray(), types.ToArray(), frame.Count, fetch); }
            catch (Exception) { return false; } // SageEditor de outra versão ou descarregado: a janela própria
        }

        DataFrameView(DataFrame frame, string title)
        {
            this.frame = frame;
            types = frame.ColumnTypes;
            long rows = frame.Count;
            List<string> names = frame.ColumnNames;

            Text = title + " - " + rows.ToString("N0", CultureInfo.CurrentCulture) + " linhas x " + names.Count + " colunas";
            Font = new Font("Segoe UI", 9f);
            ClientSize = new Size(1000, 600);
            StartPosition = FormStartPosition.CenterScreen;
            ShowIcon = false;
            KeyPreview = true;
            KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) Close(); };

            grid.Dock = DockStyle.Fill;
            grid.VirtualMode = true;
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.RowHeadersWidth = 90;
            grid.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            grid.BackgroundColor = SystemColors.Window;
            grid.BorderStyle = BorderStyle.None;
            grid.DefaultCellStyle.Font = new Font("Consolas", 9f);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9f);
            grid.EnableHeadersVisualStyles = false;
            for (int c = 0; c < names.Count; c++)
            {
                DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                column.HeaderText = names[c] + Environment.NewLine + types[c];
                column.Width = 120;
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                if (IsNumber(types[c])) column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                grid.Columns.Add(column);
            }
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            grid.RowCount = (int)Math.Min(rows, int.MaxValue);
            grid.CellValueNeeded += CellValueNeeded;
            grid.RowPostPaint += delegate(object s, DataGridViewRowPostPaintEventArgs e)
            {
                // Número da linha (base 0, como o índice do pandas)
                string label = e.RowIndex.ToString("N0", CultureInfo.CurrentCulture);
                TextRenderer.DrawText(e.Graphics, label, grid.DefaultCellStyle.Font,
                    new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth - 6, e.RowBounds.Height),
                    SystemColors.GrayText, TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
            };
            Controls.Add(grid);
        }

        void CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            long page = e.RowIndex / PageSize;
            List<object[]> rows;
            if (!pages.TryGetValue(page, out rows))
            {
                rows = frame.Rows(page * PageSize, PageSize);
                pages[page] = rows;
                order.Enqueue(page);
                if (order.Count > 20) pages.Remove(order.Dequeue()); // até 20 mil linhas em memória
            }
            int i = e.RowIndex % PageSize;
            if (i < rows.Count) e.Value = DataFrame.Display(rows[i][e.ColumnIndex], types[e.ColumnIndex]);
        }

        static bool IsNumber(string type)
        {
            return type.EndsWith("INT") || type == "INTEGER" || type == "DOUBLE" || type == "FLOAT" || type.StartsWith("DECIMAL");
        }
    }
}
