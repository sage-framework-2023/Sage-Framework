using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace SageEditor
{
    // Decoradores de cor na janela de código: cada RGB(r, g, b) e cada hexadecimal Long (&H80FF&, &H8000000F): as cores que os forms
    // entendem; as do sistema aparecem com a cor real do Windows. O VBE tira os zeros à esquerda
    // (&H0080FF& vira &H80FF&) e o & quando o número já é Long (8 dígitos); &HFF sem & é Integer
    // e fica de fora. Cada cor fica sublinhada nela mesma, no próprio desenho do VBE: o desvio do
    // texto (Syntax) chama Begin antes e End depois de cada trecho de texto normal. O clique no
    // texto da cor abre o seletor (ColorPicker), que reescreve a cor no
    // mesmo formato pelo CodeModule (ReplaceLine; sem Desfazer). A linha vem da posição na tela
    // (LineNumbers.RowToLine), e o texto da linha é conferido antes de qualquer mudança.
    static class ColorDecorators
    {
        public static dynamic Vbe;

        // Uma cor sublinhada: o texto dela (área cliente) e qual cor da linha (a n-ésima, da esquerda)
        sealed class Hit
        {
            public IntPtr Hwnd;
            public Rectangle Box;
            public int Y, X;
            public string Literal;
        }

        // Uma cor encontrada num trecho, a desenhar depois do texto
        internal sealed class Pending
        {
            public IntPtr Hwnd;
            public int Y, StartX, EndX, Height;
            public readonly List<Swatch> Swatches = new List<Swatch>();
        }

        internal sealed class Swatch
        {
            public Rectangle Box;
            public int Color;
            public string Literal;
            public int X;
        }

        static readonly List<Hit> hits = new List<Hit>();
        static readonly Native.SubclassProc subclassProc = SubclassProc;
        static readonly UIntPtr SubclassId = (UIntPtr)0x5A72;
        static readonly HashSet<IntPtr> attached = new HashSet<IntPtr>();
        static IntPtr vbeWindow;
        static int tick, lastTop = -1;
        static IntPtr lastPane;

        // RGB(255, 128, 0), &H80FF& (com &: Long) e &H8000000F (5 a 8 dígitos: Long sem o &)
        static readonly Regex pattern = new Regex(@"\bRGB\s*\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)|&H([0-9A-F]{1,8})&|&H([0-9A-F]{5,8})\b(?!&)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        static bool Enabled { get { return Settings.ColorDecorators && Vbe != null; } }

        // ------------------------------------------------------------------
        // Ciclo de vida (thread de interface do Excel)
        // ------------------------------------------------------------------

        public static void Start(IntPtr vbe) { vbeWindow = vbe; }

        // Subclassifica as janelas de código novas; rolagem por teclado: redesenha para refazer
        // as posições das cores
        public static void Poll()
        {
            if (vbeWindow == IntPtr.Zero) return;
            if (++tick % 7 == 0)
                Native.EnumChildWindows(vbeWindow, delegate(IntPtr h, IntPtr l)
                {
                    if (!attached.Contains(h) && LineNumbers.IsCodePane(h) &&
                        Native.SetWindowSubclass(h, subclassProc, SubclassId, UIntPtr.Zero))
                        attached.Add(h);
                    return true;
                }, IntPtr.Zero);
            if (!Enabled) return;
            try
            {
                dynamic pane = Vbe.ActiveCodePane;
                if (pane == null) return;
                int top = pane.TopLine;
                IntPtr hwnd = LineNumbers.HwndOfPane(pane);
                if (hwnd != lastPane) { lastPane = hwnd; lastTop = top; return; }
                if (top != lastTop)
                {
                    lastTop = top;
                    Forget(hwnd);
                    Native.InvalidateRect(hwnd, IntPtr.Zero, false);
                }
            }
            catch (Exception) { } // VBE ocupado
        }

        public static void Shutdown()
        {
            ColorPicker.CloseOpen();
            foreach (IntPtr h in attached) Native.RemoveWindowSubclass(h, subclassProc, SubclassId);
            attached.Clear();
            hits.Clear();
        }

        // Ligou ou desligou nas Configurações: redesenha as janelas de código
        public static void Repaint()
        {
            hits.Clear();
            foreach (IntPtr h in attached) if (Native.IsWindow(h)) Native.InvalidateRect(h, IntPtr.Zero, false);
        }

        static void Forget(IntPtr hwnd) { hits.RemoveAll(h => h.Hwnd == hwnd); }

        // ------------------------------------------------------------------
        // Desenho (chamado pelo desvio do texto, Syntax)
        // ------------------------------------------------------------------

        // Antes do trecho: as cores dele e onde ficam. null: nada a desenhar
        internal static Pending Begin(IntPtr hdc, IntPtr hwnd, int px, int py, IntPtr text, int count, IntPtr dx)
        {
            if (!Enabled || count < 6) return null;
            byte[] bytes = new byte[count];
            Marshal.Copy(text, bytes, 0, count);
            string s = Encoding.Default.GetString(bytes);
            if (s.IndexOf("RGB", StringComparison.OrdinalIgnoreCase) < 0 && s.IndexOf("&H", StringComparison.OrdinalIgnoreCase) < 0) return null;

            Pending p = null;
            Size one;
            GetTextExtentPoint32A(hdc, text, 1, out one);
            foreach (Match m in Matches(s))
            {
                int color;
                if (!ColorOf(m, out color)) continue;
                if (p == null) p = new Pending { Hwnd = hwnd, Y = py, StartX = px, EndX = px + Extent(hdc, text, count, dx), Height = one.Height };
                int left = px + Extent(hdc, text, m.Index, dx), right = px + Extent(hdc, text, m.Index + m.Length, dx);
                // Sublinhado na cor embaixo do texto dela
                p.Swatches.Add(new Swatch { Color = color, Literal = m.Value, X = left,
                    Box = new Rectangle(left, py + one.Height - 2, Math.Max(2, right - left), 2) });
            }
            return p;
        }

        // Depois do trecho (por cima do texto e do fundo dele)
        internal static void End(IntPtr hdc, Pending p)
        {
            if (p == null) return;
            // As cores antigas desta faixa da linha saem (o trecho foi redesenhado)
            hits.RemoveAll(h => h.Hwnd == p.Hwnd && h.Y == p.Y && h.X >= p.StartX && h.X < p.EndX);
            foreach (Swatch w in p.Swatches)
            {
                Native.RECT r = new Native.RECT { Left = w.Box.Left, Top = w.Box.Top, Right = w.Box.Right, Bottom = w.Box.Bottom };
                IntPtr fill = Native.CreateSolidBrush(w.Color);
                FillRect(hdc, ref r, fill);
                Native.DeleteObject(fill);
                // O clique vale no texto inteiro da cor
                hits.Add(new Hit { Hwnd = p.Hwnd, Box = new Rectangle(w.Box.X, p.Y, w.Box.Width, p.Height), Y = p.Y, X = w.X, Literal = w.Literal });
            }
            if (hits.Count > 2000) hits.RemoveRange(0, hits.Count - 2000);
        }

        // Largura em pixels dos primeiros n caracteres do trecho
        static int Extent(IntPtr hdc, IntPtr text, int n, IntPtr dx)
        {
            if (n <= 0) return 0;
            if (dx != IntPtr.Zero)
            {
                int sum = 0;
                for (int i = 0; i < n; i++) sum += Marshal.ReadInt32(dx, i * 4);
                return sum;
            }
            Size size;
            GetTextExtentPoint32A(hdc, text, n, out size);
            return size.Width;
        }

        // As cores do texto, fora de strings ("...") e de comentários (')
        static List<Match> Matches(string s)
        {
            List<Match> list = new List<Match>();
            bool quoted = false;
            int limit = s.Length;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '"') quoted = !quoted;
                else if (s[i] == '\'' && !quoted) { limit = i; break; }
            }
            foreach (Match m in pattern.Matches(s))
            {
                if (m.Index >= limit) break;
                int quotes = 0;
                for (int i = 0; i < m.Index; i++) if (s[i] == '"') quotes++;
                if (quotes % 2 == 0) list.Add(m);
            }
            return list;
        }

        // A cor (COLORREF, 0x00BBGGRR) de RGB(...) ou &H...&
        static bool ColorOf(Match m, out int color)
        {
            color = 0;
            if (m.Groups[1].Success)
            {
                int r = int.Parse(m.Groups[1].Value), g = int.Parse(m.Groups[2].Value), b = int.Parse(m.Groups[3].Value);
                if (r > 255 || g > 255 || b > 255) return false;
                color = r | (g << 8) | (b << 16);
                return true;
            }
            uint v = Convert.ToUInt32(m.Groups[4].Success ? m.Groups[4].Value : m.Groups[5].Value, 16);
            uint high = v >> 24;
            if (high == 0x80)
            {
                int index = (int)(v & 0xFF);
                if (index > 30) return false;
                color = (int)GetSysColor(index) & 0xFFFFFF; // a cor do Windows, sem o tema do editor
                return true;
            }
            if (high != 0) return false;
            color = (int)(v & 0xFFFFFF);
            return true;
        }

        // ------------------------------------------------------------------
        // Mouse na janela de código
        // ------------------------------------------------------------------

        const int WM_LBUTTONDOWN = 0x0201, WM_SETCURSOR = 0x0020, WM_VSCROLL = 0x0115, WM_HSCROLL = 0x0114,
            WM_MOUSEWHEEL = 0x020A, HTCLIENT = 1, IDC_HAND = 32649;
        const int VK_SHIFT = 0x10, VK_CONTROL = 0x11, VK_MENU = 0x12;

        static IntPtr SubclassProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            try
            {
                if (msg == Native.WM_NCDESTROY)
                {
                    Native.RemoveWindowSubclass(hwnd, subclassProc, SubclassId);
                    attached.Remove(hwnd);
                    Forget(hwnd);
                }
                else if (Enabled)
                {
                    switch (msg)
                    {
                        case Native.WM_PAINT:
                        {
                            // A área repintada perde as cores antigas; as novas vêm com o texto
                            Native.RECT update;
                            if (GetUpdateRect(hwnd, out update, false))
                                hits.RemoveAll(h => h.Hwnd == hwnd && h.Y >= update.Top && h.Y < update.Bottom &&
                                    h.Box.Right > update.Left && h.Box.Left < update.Right);
                            break;
                        }
                        case WM_VSCROLL:
                        case WM_HSCROLL:
                        case WM_MOUSEWHEEL:
                        {
                            // A rolagem move os pixels sem redesenhar: redesenha tudo
                            IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                            Forget(hwnd);
                            Native.InvalidateRect(hwnd, IntPtr.Zero, false);
                            return result;
                        }
                        case WM_SETCURSOR:
                            if (((int)lParam & 0xFFFF) == HTCLIENT && HitAt(hwnd, CursorPoint(hwnd)) != null)
                            {
                                SetCursor(LoadCursor(IntPtr.Zero, (IntPtr)IDC_HAND));
                                return (IntPtr)1;
                            }
                            break;
                        case WM_LBUTTONDOWN:
                        {
                            if (IsDown(VK_SHIFT) || IsDown(VK_CONTROL) || IsDown(VK_MENU)) break;
                            long lp = lParam.ToInt64();
                            Point at = new Point((short)(lp & 0xFFFF), (short)((lp >> 16) & 0xFFFF));
                            Hit hit = HitAt(hwnd, at);
                            if (hit != null && Open(hwnd, hit)) return IntPtr.Zero;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex) { Log.Error(ex); }
            return Native.DefSubclassProc(hwnd, msg, wParam, lParam);
        }

        static Hit HitAt(IntPtr hwnd, Point p)
        {
            foreach (Hit h in hits)
            {
                Rectangle box = h.Box;
                box.Inflate(1, 1);
                if (h.Hwnd == hwnd && box.Contains(p)) return h;
            }
            return null;
        }

        static Point CursorPoint(IntPtr hwnd)
        {
            POINT p;
            GetCursorPos(out p);
            ScreenToClient(hwnd, ref p);
            return new Point(p.X, p.Y);
        }

        static bool IsDown(int vk) { return GetKeyState(vk) < 0; }

        // ------------------------------------------------------------------
        // Seletor e reescrita da cor
        // ------------------------------------------------------------------

        // A cor clicada: a linha (pela posição na tela) e qual ocorrência na linha (pela ordem
        // das cores da esquerda para a direita), conferida com o texto da linha
        static bool Open(IntPtr hwnd, Hit hit)
        {
            dynamic pane;
            int line;
            if (!LineNumbers.RowToLine(hwnd, hit.Y, out pane, out line)) return false;
            dynamic module = pane.CodeModule;
            string text = (string)module.Lines(line, 1) ?? "";
            int ordinal = 0;
            foreach (Hit h in hits) if (h.Hwnd == hwnd && h.Y == hit.Y && h.X < hit.X) ordinal++;
            List<Match> found = Matches(text);
            if (ordinal >= found.Count || !string.Equals(found[ordinal].Value, hit.Literal, StringComparison.OrdinalIgnoreCase)) return false;
            int color;
            if (!ColorOf(found[ordinal], out color)) return false;

            Literal literal = new Literal(module, line, text, ordinal, found[ordinal]);
            Point screen = new Point(hit.Box.Left, hit.Box.Bottom + 2);
            POINT sp = new POINT { X = screen.X, Y = hit.Y + Math.Max(hit.Box.Height, 14) + 2 };
            ClientToScreen(hwnd, ref sp);
            ColorPicker.Open(new Point(sp.X, sp.Y), color, delegate(int c) { literal.Set(c); }, delegate { literal.Restore(); },
                (IntPtr)(long)Vbe.MainWindow.HWnd);
            return true;
        }

        // A cor dentro da linha: reescrita no mesmo formato a cada mudança no seletor
        sealed class Literal
        {
            readonly dynamic module;
            readonly int line, ordinal;
            readonly string original;
            readonly bool rgb;

            public Literal(dynamic module, int line, string text, int ordinal, Match m)
            {
                this.module = module; this.line = line; this.ordinal = ordinal; original = text;
                rgb = m.Groups[1].Success;
            }

            public void Set(int color)
            {
                int r = color & 0xFF, g = (color >> 8) & 0xFF, b = (color >> 16) & 0xFF;
                // Hexadecimal sempre com &: sem ele, até 4 dígitos seria Integer (&HFFFF = -1)
                string value = rgb ? string.Format("RGB({0}, {1}, {2})", r, g, b)
                    : "&H" + (color & 0xFFFFFF).ToString("X") + "&";
                string text = (string)module.Lines(line, 1) ?? "";
                List<Match> found = Matches(text);
                if (ordinal >= found.Count) return;
                Match m = found[ordinal];
                string updated = text.Substring(0, m.Index) + value + text.Substring(m.Index + m.Length);
                if (updated != text) module.ReplaceLine(line, updated);
            }

            public void Restore()
            {
                string text = (string)module.Lines(line, 1) ?? "";
                if (text != original) module.ReplaceLine(line, original);
            }
        }

        [StructLayout(LayoutKind.Sequential)] struct Size { public int Width, Height; }
        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        [DllImport("gdi32.dll")] static extern bool GetTextExtentPoint32A(IntPtr hdc, IntPtr text, int count, out Size size);
        [DllImport("user32.dll")] static extern int FillRect(IntPtr hdc, ref Native.RECT rect, IntPtr brush);
        [DllImport("user32.dll")] static extern uint GetSysColor(int index);
        [DllImport("user32.dll")] static extern bool GetUpdateRect(IntPtr hwnd, out Native.RECT rect, bool erase);
        [DllImport("user32.dll")] static extern IntPtr SetCursor(IntPtr cursor);
        [DllImport("user32.dll")] static extern IntPtr LoadCursor(IntPtr instance, IntPtr name);
        [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT p);
        [DllImport("user32.dll")] static extern bool ScreenToClient(IntPtr hwnd, ref POINT p);
        [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr hwnd, ref POINT p);
        [DllImport("user32.dll")] static extern short GetKeyState(int vk);
    }
}
