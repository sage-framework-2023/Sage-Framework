using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageVBE
{
    // Números de linha à esquerda da janela de código: [números][margem][código].
    // A faixa é reservada na área não-cliente da janela (WM_NCCALCSIZE), então o VBE
    // continua cuidando de tudo o que é dele (clique, cursor, seleção, rolagem) no
    // espaço que sobra; os números são desenhados nessa faixa. A altura das linhas e
    // o topo do texto vêm do próprio desenho do código (Observe, chamado pelo desvio
    // de texto); a primeira linha visível e a linha atual vêm do CodePane.
    static class LineNumbers
    {
        public static dynamic Vbe;

        sealed class Pane
        {
            public int LineHeight, LastY = -1, TextTop = int.MaxValue;
            public int Width;                       // largura reservada (0 = sem faixa)
            public int TopLine = -1, Current = -1, Total = -1;
        }

        static readonly Dictionary<IntPtr, Pane> panes = new Dictionary<IntPtr, Pane>();
        static Font font;
        static int fontSize;

        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr hwnd, ref POINT p);
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        const uint SWP_NOSIZE = 0x1, SWP_NOMOVE = 0x2, SWP_NOZORDER = 0x4, SWP_NOACTIVATE = 0x10, SWP_FRAMECHANGED = 0x20;

        static bool Enabled { get { return Settings.LineNumbers && Vbe != null && ThemeEngine.Current != null; } }

        static Pane Get(IntPtr hwnd)
        {
            Pane p;
            if (!panes.TryGetValue(hwnd, out p))
            {
                if (panes.Count > 100) panes.Clear();
                panes[hwnd] = p = new Pane();
            }
            return p;
        }

        // Posição real de cada trecho de código desenhado (coordenadas da área cliente)
        public static void Observe(IntPtr hwnd, int x, int y)
        {
            Pane p = Get(hwnd);
            if (p.LastY >= 0 && y != p.LastY)
            {
                int diff = Math.Abs(y - p.LastY);
                if (diff > 4 && (p.LineHeight == 0 || diff < p.LineHeight)) p.LineHeight = diff;
            }
            p.LastY = y;
            if (y < p.TextTop) p.TextTop = y;
        }

        // ------------------------------------------------------------------
        // Faixa na área não-cliente
        // ------------------------------------------------------------------

        // WM_NCCALCSIZE, depois do cálculo padrão: tira a faixa do lado esquerdo do cliente
        public static void AdjustClient(IntPtr hwnd, IntPtr lParam)
        {
            Pane p = Get(hwnd);
            p.Width = Enabled ? WidthFor(p.Total > 0 ? p.Total : 999) : 0;
            if (p.Width == 0) return;
            Native.RECT r = (Native.RECT)Marshal.PtrToStructure(lParam, typeof(Native.RECT));
            r.Left += p.Width;
            if (r.Right < r.Left) r.Right = r.Left;
            Marshal.StructureToPtr(r, lParam, false);
        }

        static int WidthFor(int total)
        {
            EnsureFont(16);
            string widest = new string('8', Math.Max(3, total.ToString().Length));
            return TextRenderer.MeasureText(widest, font, Size.Empty, TextFormatFlags.NoPadding).Width + 18;
        }

        // Recalcula a área não-cliente de todas as janelas de código (ligar/desligar, tema)
        public static void RefreshFrames(IEnumerable<IntPtr> windows)
        {
            foreach (IntPtr hwnd in windows)
                if (IsCodePane(hwnd))
                    SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        }

        public static bool IsCodePane(IntPtr hwnd)
        {
            return Native.ClassName(hwnd) == "VbaWindow" && Native.ClassName(Native.GetParent(hwnd)) == "MDIClient";
        }

        // Desenha a faixa com os números (chamado no WM_NCPAINT, depois de pintar o
        // código e depois de rolagem, teclas e cliques)
        public static void Paint(IntPtr hwnd, Theme t)
        {
            Pane p;
            if (!panes.TryGetValue(hwnd, out p) || p.Width == 0 || t == null) return;

            Native.RECT win, client;
            Native.GetWindowRect(hwnd, out win);
            Native.GetClientRect(hwnd, out client);
            POINT origin = new POINT();
            ClientToScreen(hwnd, ref origin);
            int left = origin.X - win.Left - p.Width, top = origin.Y - win.Top;
            int height = client.Bottom;
            if (height <= 0) return;

            // Linhas visíveis (só quando já se sabe onde o texto fica)
            int topLine = 0, total = 0, current = 0, visible = 0;
            bool rows = p.LineHeight > 0 && p.TextTop != int.MaxValue && Read(hwnd, out topLine, out total, out current, out visible);
            if (rows)
            {
                p.TopLine = topLine; p.Current = current;
                if (p.Total > 0 && Digits(total) != Digits(p.Total))
                {
                    p.Total = total;
                    RefreshFrames(new IntPtr[] { hwnd }); // mais dígitos: faixa mais larga
                    return;
                }
                p.Total = total;
            }

            IntPtr dc = Native.GetWindowDC(hwnd);
            try
            {
                using (Bitmap strip = new Bitmap(p.Width, height))
                using (Graphics g = Graphics.FromImage(strip))
                {
                    g.Clear(FromRef(t.Window));
                    if (rows)
                    {
                        EnsureFont(p.LineHeight);
                        Color normal = FromRef(Blend(t.Window, t.WindowText, 0.42));
                        Color active = FromRef(t.WindowText);
                        for (int i = 0; ; i++)
                        {
                            int y = p.TextTop + i * p.LineHeight, line = topLine + i;
                            if (i > visible || y + p.LineHeight > height - 17 || line > total) break;
                            Rectangle cell = new Rectangle(0, y, p.Width - 10, p.LineHeight);
                            TextRenderer.DrawText(g, line.ToString(), font, cell, line == current ? active : normal,
                                FromRef(t.Window), TextFormatFlags.Right | TextFormatFlags.VerticalCenter |
                                TextFormatFlags.SingleLine | TextFormatFlags.NoPadding);
                        }
                    }
                    using (Graphics target = Graphics.FromHdc(dc))
                        target.DrawImageUnscaled(strip, left, top);
                }
            }
            finally { Native.ReleaseDC(hwnd, dc); }
        }

        // Verificação periódica: rolagem, linha atual ou total de linhas mudou
        public static void Poll()
        {
            if (!Enabled) return;
            dynamic pane = Vbe.ActiveCodePane;
            if (pane == null) return;
            IntPtr hwnd = HwndOf(pane);
            Pane p;
            if (hwnd == IntPtr.Zero || !panes.TryGetValue(hwnd, out p) || p.Width == 0) return;

            int topLine, total, current;
            if (!ReadPane(pane, out topLine, out total, out current)) return;
            if (topLine != p.TopLine || current != p.Current || total != p.Total)
                Paint(hwnd, ThemeEngine.Current);
        }

        // ------------------------------------------------------------------

        static bool Read(IntPtr hwnd, out int topLine, out int total, out int current, out int visible)
        {
            topLine = total = current = visible = 0;
            dynamic pane = PaneFor(hwnd);
            if (pane == null || !ReadPane(pane, out topLine, out total, out current)) return false;
            try { visible = pane.CountOfVisibleLines; } catch (Exception) { return false; }
            return true;
        }

        static bool ReadPane(dynamic pane, out int topLine, out int total, out int current)
        {
            topLine = total = current = 0;
            try
            {
                topLine = pane.TopLine;
                total = pane.CodeModule.CountOfLines;
                int sl = 0, sc = 0, el = 0, ec = 0;
                pane.GetSelection(ref sl, ref sc, ref el, ref ec);
                current = sl;
                return true;
            }
            catch (Exception) { return false; } // VBE ocupado
        }

        // Window.HWnd vem 0 para janelas de código; o título da janela é igual ao
        // Window.Caption do CodePane ("Módulo1 (Código)").
        static dynamic PaneFor(IntPtr hwnd)
        {
            string title = TitleOf(hwnd);
            if (title.Length == 0) return null;
            try
            {
                foreach (dynamic pane in Vbe.CodePanes)
                    if ((string)pane.Window.Caption == title) return pane;
            }
            catch (Exception) { }
            return null;
        }

        static IntPtr HwndOf(dynamic pane)
        {
            string caption;
            try { caption = pane.Window.Caption; }
            catch (Exception) { return IntPtr.Zero; }
            foreach (IntPtr hwnd in panes.Keys)
                if (Native.IsWindow(hwnd) && TitleOf(hwnd) == caption) return hwnd;
            return IntPtr.Zero;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern int GetWindowText(IntPtr hwnd, System.Text.StringBuilder text, int max);

        static string TitleOf(IntPtr hwnd)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder(260);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }
        static int Digits(int n) { return Math.Max(3, n.ToString().Length); }

        static void EnsureFont(int lineHeight)
        {
            int size = Math.Max(9, lineHeight - 4);
            if (font != null && fontSize == size) return;
            if (font != null) font.Dispose();
            font = new Font("Consolas", size, FontStyle.Regular, GraphicsUnit.Pixel);
            fontSize = size;
        }

        static int Blend(int a, int b, double k)
        {
            int r = (int)((a & 0xFF) + ((b & 0xFF) - (a & 0xFF)) * k);
            int g = (int)(((a >> 8) & 0xFF) + (((b >> 8) & 0xFF) - ((a >> 8) & 0xFF)) * k);
            int bl = (int)(((a >> 16) & 0xFF) + (((b >> 16) & 0xFF) - ((a >> 16) & 0xFF)) * k);
            return r | (g << 8) | (bl << 16);
        }

        static Color FromRef(int c)
        {
            return Color.FromArgb(c & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF);
        }
    }
}
