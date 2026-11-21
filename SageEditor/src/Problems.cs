using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageEditor
{
    // Aba Problemas do Painel: os problemas do código encontrados pela análise do Sage (VbaLinter),
    // sem compilar e sem mexer no projeto. Roda pelo botão Verificar da aba, ao abrir a aba pela
    // primeira vez e ao salvar uma pasta de trabalho (VBProject.Saved passa de falso a verdadeiro).
    // Projetos protegidos por senha ficam de fora (o código deles não pode ser lido).
    static class Problems
    {
        public static dynamic Vbe;

        public sealed class Item
        {
            public string File, Folder, Component, Procedure, Message;
            public int Line, Column, Length;
        }

        public static List<Item> Last = new List<Item>();
        public static bool Checked;          // já houve uma verificação
        public static event Action Changed;  // a lista mudou (a aba e o número no título)

        // Na thread do Excel
        public static void Run()
        {
            if (Vbe == null) return;
            List<Item> items = new List<Item>();
            foreach (dynamic project in Vbe.VBProjects)
            {
                try
                {
                    if ((int)project.Protection != 0) continue; // vbext_pp_locked
                    string file = ErrorToTerminal.FileOf(project);
                    foreach (dynamic component in project.VBComponents)
                    {
                        dynamic module = component.CodeModule;
                        int count = (int)module.CountOfLines;
                        if (count == 0) continue;
                        string code = (string)module.Lines(1, count);
                        string folder = ErrorToTerminal.Folder((int)component.Type);
                        string name = (string)component.Name;
                        foreach (VbaLinter.Problem p in VbaLinter.Analyze(code))
                            items.Add(new Item
                            {
                                File = file, Folder = folder, Component = name, Procedure = p.Procedure,
                                Message = p.Message, Line = p.Line, Column = p.Column, Length = p.Length,
                            });
                    }
                }
                catch (Exception ex) { Log.Error(ex); }
            }
            Last = items;
            Checked = true;
            Log.Info("Problemas: " + items.Count);
            if (Changed != null) Changed();
        }

        // Vai até o problema: abre o módulo e seleciona o trecho
        public static void Go(Item item)
        {
            try
            {
                foreach (dynamic project in Vbe.VBProjects)
                {
                    if (ErrorToTerminal.FileOf(project) != item.File) continue;
                    dynamic pane = project.VBComponents.Item(item.Component).CodeModule.CodePane;
                    pane.Show();
                    pane.SetSelection(item.Line, item.Column, item.Line, item.Column + item.Length);
                    pane.Window.SetFocus();
                    return;
                }
            }
            catch (Exception ex) { Log.Error(ex); }
        }

        // Pelo timer do Connect (a cada ~1 s): verifica de novo quando uma pasta de trabalho é salva
        // Cada projeto pela identidade do objeto COM (o nome do arquivo muda no Salvar como)
        static readonly Dictionary<IntPtr, bool> saved = new Dictionary<IntPtr, bool>();
        static int tick;

        public static void Poll()
        {
            if (Vbe == null || ++tick % 7 != 0) return;
            bool changed = false;
            HashSet<IntPtr> seen = new HashSet<IntPtr>();
            try
            {
                foreach (dynamic project in Vbe.VBProjects)
                {
                    IntPtr key = Identity(project);
                    bool now = (bool)project.Saved;
                    bool before;
                    if (saved.TryGetValue(key, out before) && !before && now) changed = true;
                    saved[key] = now;
                    seen.Add(key);
                }
            }
            catch (Exception) { return; } // VBE ocupado: confere na próxima vez
            foreach (IntPtr key in new List<IntPtr>(saved.Keys)) if (!seen.Contains(key)) saved.Remove(key);
            if (changed) Run();
        }

        static IntPtr Identity(object o)
        {
            IntPtr unknown = Marshal.GetIUnknownForObject(o);
            Marshal.Release(unknown); // só o endereço, que identifica o objeto enquanto ele existe
            return unknown;
        }
    }

    // A lista da aba: os problemas agrupados por componente (o arquivo e a pasta ao lado do nome),
    // cada um com a mensagem e a linha. O clique (ou Enter) leva ao trecho no código.
    sealed class ProblemsView : UserControl
    {
        readonly ListBox list = new ListBox();
        readonly Label empty = new Label();
        Theme theme;
        int hover = -1;

        sealed class Row
        {
            public bool Header;
            public string Text, Detail;
            public int Count;
            public Problems.Item Item;
        }

        public ProblemsView()
        {
            list.DrawMode = DrawMode.OwnerDrawFixed;
            list.BorderStyle = BorderStyle.None;
            list.IntegralHeight = false;
            list.Dock = DockStyle.Fill;
            list.Font = SystemFonts.MessageBoxFont;
            list.ItemHeight = list.Font.Height + 8;
            list.DrawItem += DrawRow;
            list.MouseClick += delegate(object s, MouseEventArgs e) { Open(list.IndexFromPoint(e.Location)); };
            list.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.Handled = true; Open(list.SelectedIndex); } };
            list.MouseMove += delegate(object s, MouseEventArgs e)
            {
                int i = list.IndexFromPoint(e.Location);
                if (i != hover) { hover = i; list.Invalidate(); }
            };
            list.MouseLeave += delegate { if (hover != -1) { hover = -1; list.Invalidate(); } };
            list.HandleCreated += delegate { DarkScrollBar(); };

            empty.Dock = DockStyle.Fill;
            empty.Font = SystemFonts.MessageBoxFont;
            empty.TextAlign = ContentAlignment.TopLeft;
            empty.Padding = new Padding(12, 10, 12, 0);
            empty.Text = Strings.ProblemsNotChecked;

            Controls.Add(list);
            Controls.Add(empty);
            Problems.Changed += delegate { if (!IsDisposed) Refill(); };
            Refill();
        }

        // Verificar (o botão da aba): roda a análise e mostra o resultado
        public void Check()
        {
            empty.Text = Strings.ProblemsChecking;
            empty.Visible = true;
            list.Visible = false;
            empty.Refresh();
            Problems.Run();
        }

        public void FocusList() { if (list.Visible) list.Focus(); }

        void Refill()
        {
            list.BeginUpdate();
            list.Items.Clear();
            string group = null;
            Row header = null;
            foreach (Problems.Item item in Problems.Last)
            {
                string key = item.File + "\\" + item.Folder + "\\" + item.Component;
                if (key != group)
                {
                    group = key;
                    header = new Row { Header = true, Text = item.Component, Detail = item.File + "\\" + item.Folder };
                    list.Items.Add(header);
                }
                header.Count++;
                string where = string.Format(Strings.ProblemsLocation, item.Line, item.Column);
                if (!string.IsNullOrEmpty(item.Procedure)) where = item.Procedure + ", " + where;
                list.Items.Add(new Row { Text = item.Message, Detail = where, Item = item });
            }
            list.EndUpdate();
            bool any = list.Items.Count > 0;
            list.Visible = any;
            empty.Visible = !any;
            empty.Text = Problems.Checked ? Strings.ProblemsNone : Strings.ProblemsNotChecked;
        }

        void Open(int index)
        {
            if (index < 0 || index >= list.Items.Count) return;
            Row row = (Row)list.Items[index];
            if (row.Item != null) Problems.Go(row.Item);
        }

        public void ApplyTheme(Theme t)
        {
            theme = t;
            BackColor = t.Background;
            list.BackColor = t.Background;
            list.ForeColor = t.Foreground;
            empty.BackColor = t.Background;
            empty.ForeColor = t.Muted;
            DarkScrollBar();
            list.Invalidate();
        }

        void DarkScrollBar()
        {
            if (theme != null && list.IsHandleCreated) Native.SetWindowTheme(list.Handle, theme.IsDark ? "DarkMode_Explorer" : "Explorer", null);
        }

        void DrawRow(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= list.Items.Count) return;
            Theme t = theme ?? Theme.Find(Settings.ColorTheme);
            Row row = (Row)list.Items[e.Index];
            Graphics g = e.Graphics;
            Rectangle r = e.Bounds;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            Color back = selected ? t.Hover : e.Index == hover ? t.Hover : t.Background;
            using (SolidBrush b = new SolidBrush(back)) g.FillRectangle(b, r);
            const TextFormatFlags flags = TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix |
                TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis;
            // Medir sem as reticências: com elas, a medida volta a largura das reticências
            const TextFormatFlags measure = TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;
            float scale = g.DpiX / 96f;
            int pad = (int)Math.Round(8 * scale);
            if (row.Header)
            {
                // ▾ Componente   Arquivo\Pasta   (n)
                int x = r.Left + pad;
                TextRenderer.DrawText(g, "▾", list.Font, new Rectangle(x, r.Top, r.Height, r.Height), t.Muted, flags);
                x += (int)Math.Round(14 * scale);
                using (Font bold = new Font(list.Font, FontStyle.Bold))
                {
                    int w = TextRenderer.MeasureText(g, row.Text, bold, Size.Empty, measure).Width;
                    TextRenderer.DrawText(g, row.Text, bold, new Rectangle(x, r.Top, w + 2, r.Height), t.Foreground, flags);
                    x += w + pad;
                }
                string detail = row.Detail + "   (" + row.Count + ")";
                TextRenderer.DrawText(g, detail, list.Font, new Rectangle(x, r.Top, Math.Max(0, r.Right - x - pad), r.Height), t.Muted, flags);
                return;
            }
            // ⊗ mensagem   procedimento, linha n, coluna n
            int left = r.Left + pad + (int)Math.Round(20 * scale);
            int size = (int)Math.Round(12 * scale);
            Rectangle icon = new Rectangle(left, r.Top + (r.Height - size) / 2, size, size);
            DrawErrorIcon(g, icon, t.IsDark ? Color.FromArgb(0xF1, 0x4C, 0x4C) : Color.FromArgb(0xE5, 0x14, 0x00));
            int textLeft = icon.Right + pad;
            int msgWidth = TextRenderer.MeasureText(g, row.Text, list.Font, Size.Empty, measure).Width;
            TextRenderer.DrawText(g, row.Text, list.Font, new Rectangle(textLeft, r.Top, Math.Max(0, Math.Min(msgWidth + 2, r.Right - textLeft - pad)), r.Height), t.Foreground, flags);
            int detailLeft = textLeft + msgWidth + 2 * pad;
            if (detailLeft < r.Right - pad)
                TextRenderer.DrawText(g, row.Detail, list.Font, new Rectangle(detailLeft, r.Top, r.Right - detailLeft - pad, r.Height), t.Muted, flags);
        }

        // O ícone de erro: um círculo com um ×, como o dos editores modernos
        static void DrawErrorIcon(Graphics g, Rectangle r, Color color)
        {
            System.Drawing.Drawing2D.SmoothingMode old = g.SmoothingMode;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using (Pen pen = new Pen(color, Math.Max(1f, r.Width / 10f)))
            {
                g.DrawEllipse(pen, r.Left + 0.5f, r.Top + 0.5f, r.Width - 1f, r.Height - 1f);
                float m = r.Width * 0.32f;
                g.DrawLine(pen, r.Left + m, r.Top + m, r.Right - m, r.Bottom - m);
                g.DrawLine(pen, r.Left + m, r.Bottom - m, r.Right - m, r.Top + m);
            }
            g.SmoothingMode = old;
        }
    }
}
