using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SageEditor
{
    // Abas das janelas abertas (módulos, classes, formulários) no topo da área de
    // código, como no VS Code. A faixa fica na área não-cliente do MDIClient
    // (WM_NCCALCSIZE), então as janelas de código, maximizadas ou não, ocupam o
    // espaço que sobra embaixo dela e o VBE não precisa saber que ela existe.
    // Clique ativa a janela e, arrastando, muda a aba de lugar; o "×" ou o botão do
    // meio fecha; o botão direito abre o menu (Fechar, Fechar Outras...). A lista vem dos filhos do MDIClient (Win32,
    // sem COM), então funciona também com o VBA ocupado. Independe do ThemeEngine:
    // funciona com o tema Padrão do VBE.
    static class EditorTabs
    {
        sealed class Tab
        {
            public IntPtr Hwnd;
            public string Name, Kind, Project, Detail;
            public Rectangle Bounds, Close;
        }

        static readonly Native.SubclassProc subclassProc = SubclassProc;
        static readonly UIntPtr SubclassId = (UIntPtr)0x5A6F;
        static IntPtr vbeWindow, mdi;

        static readonly List<IntPtr> order = new List<IntPtr>(); // ordem de abertura, como no VS Code
        static readonly List<Tab> tabs = new List<Tab>();
        static IntPtr active;
        static string signature;
        static int hover = -1, scroll;
        static bool hoverClose, tracking;
        static IntPtr dragHwnd;      // aba sendo arrastada para mudar de lugar
        static Font font;
        static int dpi, height;
        static ContextMenuStrip menu;

        const int WM_QUIT = 0x0012, WM_KEYDOWN = 0x0100, WM_MOUSEFIRST = 0x0200, WM_MOUSELAST = 0x020E;
        const int VK_LBUTTON = 0x01, VK_RBUTTON = 0x02, VK_ESCAPE = 0x1B, SM_SWAPBUTTON = 23;
        const uint PM_REMOVE = 0x1;
        const int WM_SIZE = 0x0005, WM_NCCALCSIZE = 0x0083, WM_NCHITTEST = 0x0084, WM_SYSCOMMAND = 0x0112,
            WM_NCMOUSEMOVE = 0x00A0, WM_NCLBUTTONDOWN = 0x00A1, WM_NCLBUTTONDBLCLK = 0x00A3,
            WM_NCRBUTTONDOWN = 0x00A4, WM_NCRBUTTONUP = 0x00A5, WM_NCMBUTTONDOWN = 0x00A7, WM_NCMBUTTONUP = 0x00A8,
            WM_NCMOUSELEAVE = 0x02A2, WM_MDIRESTORE = 0x0223, WM_MDIACTIVATE = 0x0222, WM_MDIGETACTIVE = 0x0229;
        const int HTBORDER = 18, SC_CLOSE = 0xF060;
        const uint GW_HWNDNEXT = 2, GW_CHILD = 5, TME_LEAVE = 0x2, TME_NONCLIENT = 0x10;
        const uint SWP_NOSIZE = 0x1, SWP_NOMOVE = 0x2, SWP_NOZORDER = 0x4, SWP_NOACTIVATE = 0x10, SWP_FRAMECHANGED = 0x20;

        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] struct TRACKMOUSEEVENT { public int cbSize; public uint dwFlags; public IntPtr hwndTrack; public uint dwHoverTime; }
        [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr hwnd, ref POINT p);
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr hwnd, uint cmd);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] static extern bool TrackMouseEvent(ref TRACKMOUSEEVENT tme);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);
        [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr hwnd);
        [DllImport("user32.dll")] static extern IntPtr SetCapture(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool ReleaseCapture();
        [DllImport("user32.dll")] static extern IntPtr GetCapture();
        [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT p);
        [DllImport("user32.dll")] static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
        [DllImport("user32.dll")] static extern bool PeekMessage(out MSG msg, IntPtr hwnd, uint min, uint max, uint remove);
        [DllImport("user32.dll")] static extern bool TranslateMessage(ref MSG msg);
        [DllImport("user32.dll")] static extern IntPtr DispatchMessage(ref MSG msg);
        [DllImport("user32.dll")] static extern void PostQuitMessage(int code);
        [StructLayout(LayoutKind.Sequential)] struct MSG { public IntPtr hwnd; public int message; public IntPtr wParam, lParam; public uint time; public POINT pt; }

        static bool Shown { get { return mdi != IntPtr.Zero && Settings.EditorTabs; } }

        // ------------------------------------------------------------------
        // Ciclo de vida (thread de interface do Excel)
        // ------------------------------------------------------------------

        public static void Start(IntPtr vbe)
        {
            vbeWindow = vbe;
            Attach();
        }

        // Configuração ou tema mudou
        public static void Refresh()
        {
            Attach();
            if (mdi == IntPtr.Zero) return;
            Relayout(mdi);
            signature = null;
            Poll();
        }

        public static void Shutdown()
        {
            if (menu != null) { menu.Dispose(); menu = null; }
            if (mdi == IntPtr.Zero) return;
            IntPtr h = mdi;
            mdi = IntPtr.Zero;
            Native.RemoveWindowSubclass(h, subclassProc, SubclassId);
            Relayout(h);
        }

        // Verificação periódica: janela aberta, fechada, ativada ou renomeada
        public static void Poll()
        {
            Attach();
            if (!Shown) return;
            string sig = Collect();
            if (sig != signature)
            {
                signature = sig;
                Paint();
            }
        }

        static void Attach()
        {
            if (mdi != IntPtr.Zero && Native.IsWindow(mdi)) return;
            mdi = IntPtr.Zero;
            if (vbeWindow == IntPtr.Zero) return;
            IntPtr found = IntPtr.Zero;
            Native.EnumChildWindows(vbeWindow, delegate(IntPtr h, IntPtr l)
            {
                if (Native.ClassName(h) != "MDIClient") return true;
                found = h;
                return false;
            }, IntPtr.Zero);
            if (found == IntPtr.Zero || !Native.SetWindowSubclass(found, subclassProc, SubclassId, UIntPtr.Zero)) return;
            mdi = found;
            signature = null;
            Relayout(mdi);
        }

        // Recalcula a área não-cliente; o WM_SIZE faz o MDIClient reposicionar a janela maximizada
        static void Relayout(IntPtr h)
        {
            SetWindowPos(h, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
            Native.RECT c;
            Native.GetClientRect(h, out c);
            Native.SendMessage(h, WM_SIZE, IntPtr.Zero, (IntPtr)((c.Bottom << 16) | (c.Right & 0xFFFF)));
            Native.RedrawWindow(h, IntPtr.Zero, IntPtr.Zero, Native.RDW_INVALIDATE | Native.RDW_FRAME | Native.RDW_ALLCHILDREN);
        }

        // ------------------------------------------------------------------
        // Mensagens do MDIClient
        // ------------------------------------------------------------------

        static IntPtr SubclassProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            try
            {
                if (msg == Native.WM_NCDESTROY)
                {
                    Native.RemoveWindowSubclass(hwnd, subclassProc, SubclassId);
                    if (hwnd == mdi) mdi = IntPtr.Zero;
                }
                else if (hwnd == mdi && Shown)
                {
                    bool c;
                    int i;
                    switch (msg)
                    {
                        case WM_NCCALCSIZE:
                        {
                            IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                            EnsureFont();
                            Native.RECT r = (Native.RECT)Marshal.PtrToStructure(lParam, typeof(Native.RECT));
                            r.Top += height;
                            if (r.Bottom < r.Top) r.Bottom = r.Top;
                            Marshal.StructureToPtr(r, lParam, false);
                            return result;
                        }

                        case Native.WM_NCPAINT:
                        {
                            IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                            Paint();
                            return result;
                        }

                        case WM_NCHITTEST:
                            if (InStrip(lParam)) return (IntPtr)HTBORDER;
                            break;

                        case WM_NCMOUSEMOVE:
                            if ((int)wParam != HTBORDER) break;
                            if (!tracking) TrackLeave();
                            i = HitTest(lParam, out c);
                            if (i != hover || c != hoverClose) { hover = i; hoverClose = c; Paint(); }
                            return IntPtr.Zero;

                        case WM_NCMOUSELEAVE:
                            tracking = false;
                            if (hover != -1) { hover = -1; hoverClose = false; Paint(); }
                            break;

                        case WM_NCLBUTTONDOWN:
                        case WM_NCLBUTTONDBLCLK:
                            if ((int)wParam != HTBORDER) break;
                            i = HitTest(lParam, out c);
                            if (i >= 0)
                            {
                                if (c) CloseWindow(tabs[i].Hwnd);
                                else
                                {
                                    // Ativa já no clique; se o mouse andar com o botão pressionado, arrasta
                                    IntPtr h = tabs[i].Hwnd;
                                    Activate(h);
                                    Drag(h, ScreenPoint(lParam));
                                }
                            }
                            return IntPtr.Zero;

                        case WM_NCMBUTTONDOWN:
                        case WM_NCRBUTTONDOWN:
                            if ((int)wParam == HTBORDER) return IntPtr.Zero;
                            break;

                        case WM_NCMBUTTONUP:
                            if ((int)wParam != HTBORDER) break;
                            i = HitTest(lParam, out c);
                            if (i >= 0) CloseWindow(tabs[i].Hwnd);
                            return IntPtr.Zero;

                        case WM_NCRBUTTONUP:
                            if ((int)wParam != HTBORDER) break;
                            i = HitTest(lParam, out c);
                            if (i >= 0) ShowMenu(i, ScreenPoint(lParam));
                            return IntPtr.Zero;
                    }
                }
            }
            catch (Exception ex) { Log.Error(ex); }
            return Native.DefSubclassProc(hwnd, msg, wParam, lParam);
        }

        static void TrackLeave()
        {
            TRACKMOUSEEVENT tme = new TRACKMOUSEEVENT();
            tme.cbSize = Marshal.SizeOf(typeof(TRACKMOUSEEVENT));
            tme.dwFlags = TME_LEAVE | TME_NONCLIENT;
            tme.hwndTrack = mdi;
            tracking = TrackMouseEvent(ref tme);
        }

        // ------------------------------------------------------------------
        // Janelas abertas
        // ------------------------------------------------------------------

        static string Collect()
        {
            List<IntPtr> open = new List<IntPtr>();
            for (IntPtr h = GetWindow(mdi, GW_CHILD); h != IntPtr.Zero; h = GetWindow(h, GW_HWNDNEXT))
                if (IsWindowVisible(h) && !Native.ClassName(h).StartsWith("#")) open.Add(h);

            // Novas no fim; GetWindow vem da mais recente para a mais antiga
            order.RemoveAll(h => !open.Contains(h));
            for (int i = open.Count - 1; i >= 0; i--)
                if (!order.Contains(open[i])) order.Add(open[i]);
            active = Native.SendMessage(mdi, WM_MDIGETACTIVE, IntPtr.Zero, IntPtr.Zero);

            tabs.Clear();
            StringBuilder sig = new StringBuilder();
            foreach (IntPtr h in order)
            {
                string title = TitleOf(h);
                tabs.Add(Parse(h, title));
                sig.Append(h.ToInt64()).Append('=').Append(title).Append('|');
            }
            sig.Append(active.ToInt64());

            // Mesmo nome em projetos diferentes: mostra o projeto, como o VS Code mostra a pasta
            foreach (Tab a in tabs)
                foreach (Tab b in tabs)
                    if (a != b && a.Name == b.Name && a.Kind == b.Kind && a.Project != b.Project)
                    {
                        a.Detail = (a.Detail + "  " + a.Project).Trim();
                        break;
                    }
            return sig.ToString();
        }

        // "Pasta1.xlsm - Módulo1 (Código)" -> nome "Módulo1"; outros tipos de janela
        // ("UserForm1 (UserForm)") mostram o tipo ao lado do nome
        static Tab Parse(IntPtr hwnd, string title)
        {
            Tab tab = new Tab();
            tab.Hwnd = hwnd;
            tab.Kind = tab.Project = "";
            string name = title;
            int p = name.LastIndexOf(" (");
            if (p > 0 && name.EndsWith(")"))
            {
                tab.Kind = name.Substring(p + 2, name.Length - p - 3);
                name = name.Substring(0, p);
            }
            int d = name.LastIndexOf(" - ");
            if (d >= 0)
            {
                tab.Project = name.Substring(0, d);
                name = name.Substring(d + 3);
            }
            tab.Name = name.Length > 0 ? name : title;
            tab.Detail = Native.ClassName(hwnd) == "VbaWindow" ? "" : tab.Kind;
            return tab;
        }

        static string TitleOf(IntPtr hwnd)
        {
            StringBuilder sb = new StringBuilder(260);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        // Arrastar com o botão pressionado. Loop próprio, como o Windows faz ao mover
        // janelas: acompanha o cursor (GetCursorPos/GetAsyncKeyState) em vez de esperar
        // WM_MOUSEMOVE, que o loop de mensagens do VBE não entrega com a captura ativa.
        // Pintura, teclado e timers continuam sendo despachados; o mouse é descartado.
        static void Drag(IntPtr hwnd, Point start)
        {
            int button = GetSystemMetrics(SM_SWAPBUTTON) != 0 ? VK_RBUTTON : VK_LBUTTON;
            Size min = SystemInformation.DragSize;
            bool dragging = false;
            dragHwnd = hwnd;
            SetCapture(mdi);
            try
            {
                MSG msg;
                while ((GetAsyncKeyState(button) & 0x8000) != 0 && GetCapture() == mdi)
                {
                    while (PeekMessage(out msg, IntPtr.Zero, 0, 0, PM_REMOVE))
                    {
                        if (msg.message == WM_QUIT) { PostQuitMessage((int)msg.wParam); return; }
                        if (msg.message >= WM_MOUSEFIRST && msg.message <= WM_MOUSELAST) continue;
                        if (msg.message == WM_KEYDOWN && (int)msg.wParam == VK_ESCAPE) return;
                        TranslateMessage(ref msg);
                        DispatchMessage(ref msg);
                    }
                    if (mdi == IntPtr.Zero) return;

                    POINT p;
                    GetCursorPos(out p);
                    if (!dragging)
                        dragging = Math.Abs(p.X - start.X) > min.Width / 2 || Math.Abs(p.Y - start.Y) > min.Height / 2;
                    if (dragging)
                    {
                        POINT o = new POINT();
                        ClientToScreen(mdi, ref o);
                        Reorder(p.X - o.X);
                    }
                    System.Threading.Thread.Sleep(10);
                }
            }
            finally
            {
                dragHwnd = IntPtr.Zero;
                if (GetCapture() == mdi) ReleaseCapture();
            }
        }

        // Troca com a vizinha quando o cursor passa do meio dela. Comparar com o meio
        // da vizinha (e não com a borda) evita que abas de larguras diferentes fiquem
        // trocando de lugar sem parar.
        static void Reorder(int x)
        {
            EnsureFont();
            Native.RECT c;
            Native.GetClientRect(mdi, out c);
            Layout(c.Right);
            int i = tabs.FindIndex(t => t.Hwnd == dragHwnd);
            if (i < 0) return;
            hover = -1;
            bool moved = false;
            while (true)
            {
                int j;
                if (i < tabs.Count - 1 && x > Center(tabs[i + 1])) j = i + 1;
                else if (i > 0 && x < Center(tabs[i - 1])) j = i - 1;
                else break;
                Tab t = tabs[i]; tabs[i] = tabs[j]; tabs[j] = t;
                IntPtr h = order[i]; order[i] = order[j]; order[j] = h;
                i = j;
                moved = true;
                Layout(c.Right);
            }
            if (moved) Paint();
        }

        static int Center(Tab t) { return t.Bounds.Left + t.Bounds.Width / 2; }

        static void Activate(IntPtr hwnd)
        {
            if (IsIconic(hwnd)) Native.SendMessage(mdi, WM_MDIRESTORE, hwnd, IntPtr.Zero);
            Native.SendMessage(mdi, WM_MDIACTIVATE, hwnd, IntPtr.Zero);
            Poll();
        }

        // Como o "×" da própria janela (o VBE só esconde a janela de código)
        static void CloseWindow(IntPtr hwnd)
        {
            PostMessage(hwnd, WM_SYSCOMMAND, (IntPtr)SC_CLOSE, IntPtr.Zero);
        }

        static void ShowMenu(int index, Point screen)
        {
            if (menu != null) menu.Dispose();
            Theme t = Theme.Find(Settings.ColorTheme);
            List<IntPtr> all = new List<IntPtr>();
            foreach (Tab tab in tabs) all.Add(tab.Hwnd);
            IntPtr target = all[index];

            menu = new ContextMenuStrip();
            menu.Renderer = new ToolStripProfessionalRenderer(new MenuColors(t));
            menu.BackColor = t.Sidebar;
            menu.ForeColor = t.Foreground;
            AddItem(Strings.Close, true, delegate { CloseWindow(target); });
            AddItem(Strings.CloseOthers, all.Count > 1, delegate { foreach (IntPtr h in all) if (h != target) CloseWindow(h); });
            AddItem(Strings.CloseToRight, index < all.Count - 1, delegate { for (int i = index + 1; i < all.Count; i++) CloseWindow(all[i]); });
            menu.Items.Add(new ToolStripSeparator());
            AddItem(Strings.CloseAll, true, delegate { foreach (IntPtr h in all) CloseWindow(h); });
            menu.Show(screen);
        }

        static void AddItem(string text, bool enabled, EventHandler click)
        {
            ToolStripMenuItem item = new ToolStripMenuItem(text, null, click);
            item.Enabled = enabled;
            menu.Items.Add(item);
        }

        sealed class MenuColors : ProfessionalColorTable
        {
            readonly Theme t;
            public MenuColors(Theme theme) { t = theme; }
            public override Color ToolStripDropDownBackground { get { return t.Sidebar; } }
            public override Color ImageMarginGradientBegin { get { return t.Sidebar; } }
            public override Color ImageMarginGradientMiddle { get { return t.Sidebar; } }
            public override Color ImageMarginGradientEnd { get { return t.Sidebar; } }
            public override Color MenuBorder { get { return t.Border; } }
            public override Color MenuItemBorder { get { return t.Hover; } }
            public override Color MenuItemSelected { get { return t.Hover; } }
            public override Color SeparatorDark { get { return t.Border; } }
            public override Color SeparatorLight { get { return t.Sidebar; } }
        }

        // ------------------------------------------------------------------
        // Posição e desenho
        // ------------------------------------------------------------------

        static void EnsureFont()
        {
            int d = 96;
            try { d = (int)GetDpiForWindow(mdi); } catch (EntryPointNotFoundException) { }
            if (d <= 0) d = 96;
            if (font != null && dpi == d) return;
            if (font != null) font.Dispose();
            dpi = d;
            font = new Font("Segoe UI", S(12), FontStyle.Regular, GraphicsUnit.Pixel); // 9 pt
            height = S(24);
        }

        static int S(int px) { return (int)Math.Round(px * dpi / 96.0); }
        static int IconSize { get { return S(16); } } // ícone do componente antes do nome (ProjectIcons)

        static Point ScreenPoint(IntPtr lParam)
        {
            long v = lParam.ToInt64();
            return new Point((short)(v & 0xFFFF), (short)((v >> 16) & 0xFFFF));
        }

        // Ponto da tela -> coordenadas da faixa (0,0 = canto superior esquerdo)
        static Point StripPoint(IntPtr lParam)
        {
            Point p = ScreenPoint(lParam);
            POINT o = new POINT();
            ClientToScreen(mdi, ref o);
            return new Point(p.X - o.X, p.Y - (o.Y - height));
        }

        static bool InStrip(IntPtr lParam)
        {
            Point p = StripPoint(lParam);
            Native.RECT c;
            Native.GetClientRect(mdi, out c);
            return p.Y >= 0 && p.Y < height && p.X >= 0 && p.X < c.Right;
        }

        static int HitTest(IntPtr lParam, out bool onClose)
        {
            onClose = false;
            // A verificação periódica refaz a lista sem posições; elas só vêm do Layout
            EnsureFont();
            Native.RECT c;
            Native.GetClientRect(mdi, out c);
            Layout(c.Right);
            Point p = StripPoint(lParam);
            for (int i = 0; i < tabs.Count; i++)
            {
                if (!tabs[i].Bounds.Contains(p)) continue;
                Rectangle close = tabs[i].Close;
                close.Inflate(S(2), S(2));
                onClose = close.Contains(p);
                return i;
            }
            return -1;
        }

        static int Measure(string text)
        {
            return text.Length == 0 ? 0 : TextRenderer.MeasureText(text, font, Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix).Width;
        }

        // Largura pelo texto; sem espaço, as abas encolhem até um mínimo e depois a
        // faixa rola para manter a aba ativa visível
        static void Layout(int width)
        {
            int padL = S(10), padR = S(4), gap = S(6), closeSize = S(16), minWidth = S(72);
            int n = tabs.Count;
            int[] w = new int[n];
            int total = 0, activeIndex = -1;
            for (int i = 0; i < n; i++)
            {
                int text = Measure(tabs[i].Name);
                if (tabs[i].Detail.Length > 0) text += gap + Measure(tabs[i].Detail);
                if (Settings.Icons) text += IconSize + gap;
                w[i] = padL + text + gap + closeSize + padR;
                total += w[i];
                if (tabs[i].Hwnd == active) activeIndex = i;
            }
            if (total > width && total > 0)
            {
                int fitted = 0;
                for (int i = 0; i < n; i++) { w[i] = Math.Max(minWidth, (int)((long)w[i] * width / total)); fitted += w[i]; }
                total = fitted;
            }

            if (total <= width) scroll = 0;
            else if (activeIndex >= 0)
            {
                int start = 0;
                for (int i = 0; i < activeIndex; i++) start += w[i];
                if (start < scroll) scroll = start;
                if (start + w[activeIndex] > scroll + width) scroll = start + w[activeIndex] - width;
                scroll = Math.Max(0, Math.Min(scroll, total - width));
            }

            int x = -scroll;
            for (int i = 0; i < n; i++)
            {
                tabs[i].Bounds = new Rectangle(x, 0, w[i], height);
                tabs[i].Close = new Rectangle(x + w[i] - padR - closeSize, (height - closeSize) / 2, closeSize, closeSize);
                x += w[i];
            }
        }

        static void Paint()
        {
            if (!Shown) return;
            EnsureFont();
            Native.RECT win, client;
            Native.GetWindowRect(mdi, out win);
            Native.GetClientRect(mdi, out client);
            POINT o = new POINT();
            ClientToScreen(mdi, ref o);
            int width = client.Right;
            int left = o.X - win.Left, top = o.Y - win.Top - height;
            if (width <= 0 || top < 0) return;

            if (signature == null) signature = Collect();
            Layout(width);
            Theme t = Theme.Find(Settings.ColorTheme);

            IntPtr dc = Native.GetWindowDC(mdi);
            try
            {
                using (Bitmap strip = new Bitmap(width, height))
                {
                    using (Graphics g = Graphics.FromImage(strip)) Draw(g, t, width);
                    using (Graphics target = Graphics.FromHdc(dc)) target.DrawImageUnscaled(strip, left, top);
                }
            }
            finally { Native.ReleaseDC(mdi, dc); }
        }

        static void Draw(Graphics g, Theme t, int width)
        {
            g.Clear(t.Sidebar);
            using (Pen line = new Pen(t.Border)) g.DrawLine(line, 0, height - 1, width, height - 1);

            int padL = S(10), gap = S(6);
            const TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine |
                TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix;

            for (int i = 0; i < tabs.Count; i++)
            {
                Tab tab = tabs[i];
                Rectangle r = tab.Bounds;
                if (r.Right <= 0 || r.Left >= width) continue;
                bool on = tab.Hwnd == active, hot = i == hover;
                Color back = on ? t.Background : hot ? t.Hover : t.Sidebar;
                Color fore = on ? t.Foreground : t.Muted;

                // A aba ativa se junta ao editor (sem a linha de baixo) e tem o destaque em cima
                using (SolidBrush b = new SolidBrush(back)) g.FillRectangle(b, r.Left, 0, r.Width, on ? height : height - 1);
                using (Pen p = new Pen(t.Border)) g.DrawLine(p, r.Right - 1, 0, r.Right - 1, height - 1);
                if (on) using (SolidBrush b = new SolidBrush(t.Accent)) g.FillRectangle(b, r.Left, 0, r.Width - 1, Math.Max(1, S(1)));

                int x = r.Left + padL, right = tab.Close.Left - gap;
                if (Settings.Icons && right - x > IconSize)
                {
                    Bitmap icon = ProjectIcons.Image(ProjectIcons.ForTab(tab.Name, tab.Kind), IconSize);
                    if (icon != null) g.DrawImage(icon, x, (height - IconSize) / 2, IconSize, IconSize);
                    x += IconSize + gap;
                }
                if (right > x)
                {
                    int nameWidth = Math.Min(Measure(tab.Name), right - x);
                    TextRenderer.DrawText(g, tab.Name, font, new Rectangle(x, 0, nameWidth, height), fore, back, flags);
                    int dx = x + nameWidth + gap;
                    if (tab.Detail.Length > 0 && right - dx > S(12))
                        TextRenderer.DrawText(g, tab.Detail, font, new Rectangle(dx, 0, right - dx, height),
                            on ? t.Muted : Blend(t.Muted, back, 0.35), back, flags);
                }

                if (on || hot)
                {
                    Rectangle c = tab.Close;
                    SmoothingMode previous = g.SmoothingMode;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    if (hot && hoverClose)
                        using (SolidBrush b = new SolidBrush(on ? t.Hover : t.Border))
                        using (GraphicsPath path = Rounded(c, S(4)))
                            g.FillPath(b, path);
                    int s = S(4), cx = c.Left + c.Width / 2, cy = c.Top + c.Height / 2;
                    using (Pen p = new Pen(fore, Math.Max(1f, S(1) * 1.2f)))
                    {
                        g.DrawLine(p, cx - s, cy - s, cx + s, cy + s);
                        g.DrawLine(p, cx + s, cy - s, cx - s, cy + s);
                    }
                    g.SmoothingMode = previous;
                }
            }
        }

        static GraphicsPath Rounded(Rectangle r, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int d = radius * 2;
            path.AddArc(r.Left, r.Top, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Top, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        static Color Blend(Color a, Color b, double k)
        {
            return Color.FromArgb((int)(a.R + (b.R - a.R) * k), (int)(a.G + (b.G - a.G) * k), (int)(a.B + (b.B - a.B) * k));
        }
    }
}
