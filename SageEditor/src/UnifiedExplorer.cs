using System;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SageEditor
{
    // Explorador unificado: a janela Propriedades fica dentro da janela Projeto, embaixo da árvore,
    // numa seção que recolhe e abre como as do Explorer do VS Code ("▾ PROPRIEDADES - Módulo1").
    //
    // A janela Propriedades continua sendo a do VBE: para ele ela fica fechada (não ocupa espaço
    // no layout dele) e a janela Win32 (wndclass_pbrs) passa a ser filha da janela Projeto (PROJECT).
    // Depois de o VBE arrumar a janela Projeto (WM_SIZE), a árvore fica só com a parte de cima; a
    // faixa da seção e as Propriedades ocupam a de baixo. Clicar na faixa recolhe ou abre; arrastar
    // muda a altura. F4 e Exibir > Janela Propriedades abrem a seção.
    static class UnifiedExplorer
    {
        public static dynamic Vbe;
        static IntPtr vbeWindow, project, tree, properties;
        static ExplorerHeader header;
        static bool collapsed;
        static int height = -1; // altura das Propriedades (px)
        static int tick;
        static ButtonClick menuClick;

        const int vbext_wt_ProjectWindow = 6, vbext_wt_PropertyWindow = 7, PropertiesId = 222;
        const string CollapsedKey = "sage.explorer.propertiesCollapsed", HeightKey = "sage.explorer.propertiesHeight";
        static readonly Native.SubclassProc projectProc = ProjectProc, propertiesProc = PropertiesProc, treeProc = TreeProc;
        static readonly UIntPtr SubclassId = (UIntPtr)0x5A74;

        public static void Start(IntPtr vbe)
        {
            vbeWindow = vbe;
            collapsed = Settings.Get(CollapsedKey, "") == "yes";
            int h;
            if (int.TryParse(Settings.Get(HeightKey, ""), NumberStyles.Integer, CultureInfo.InvariantCulture, out h)) height = h;
            try
            {
                dynamic control = Vbe.CommandBars.FindControl(Type.Missing, PropertiesId);
                Action open = delegate { ShowProperties(); };
                if (control != null) menuClick = new ButtonClick((object)control, open, true);
            }
            catch (Exception ex) { Log.Error(ex); }
        }

        public static void Poll()
        {
            if (vbeWindow == IntPtr.Zero || ++tick % 3 != 0) return;
            if (!Settings.UnifiedExplorer) { Release(true); return; }
            IntPtr p = FindChild(vbeWindow, "PROJECT");
            if (p == IntPtr.Zero || !Native.IsWindowVisible(p)) { Release(false); return; } // fechada: as Propriedades vão junto
            if (p != project) { Release(false); Attach(p); }

            dynamic w = PropertiesWindow();
            if (w == null) return;
            bool vbeShows = (bool)w.Visible;
            if (vbeShows || properties == IntPtr.Zero || !Native.IsWindow(properties) || Native.GetParent(properties) != project)
            {
                // F4 ou Exibir > Janela Propriedades: o VBE mostrou a janela solta; ela volta para a seção, aberta
                if (vbeShows && properties != IntPtr.Zero && collapsed) SetCollapsed(false);
                Borrow(w);
            }
            if (header != null) header.UpdateTitle(Title(properties));
        }

        static void Attach(IntPtr p)
        {
            project = p;
            Native.SetWindowSubclass(project, projectProc, SubclassId, UIntPtr.Zero);
            tree = FindChild(project, "SysTreeView32");
            if (tree != IntPtr.Zero) Native.SetWindowSubclass(tree, treeProc, SubclassId, UIntPtr.Zero);
            header = new ExplorerHeader();
            header.CreateControl();
            SetParent(header.Handle, project);
            header.Toggle += delegate { SetCollapsed(!collapsed); };
            header.Resized += delegate(int y) { Resize(y); };
            header.DragDone += delegate { SaveHeight(); };
            header.ApplyTheme();
            // a barra de título passa a ser a do Sage (sem o ×)
            SetWindowPos(project, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
            Layout();
        }

        static void Borrow(dynamic w)
        {
            IntPtr h = FindPropertiesHwnd();
            if (h == IntPtr.Zero)
            {
                w.Visible = true; // nunca aberta nesta sessão: o VBE cria a janela ao mostrá-la
                h = FindPropertiesHwnd();
                if (h == IntPtr.Zero) return;
            }
            if (h != properties)
            {
                if (properties != IntPtr.Zero && Native.IsWindow(properties)) Native.RemoveWindowSubclass(properties, propertiesProc, SubclassId);
                properties = h;
                Native.SetWindowSubclass(h, propertiesProc, SubclassId, UIntPtr.Zero);
            }
            if ((bool)w.Visible) w.Visible = false; // para o VBE fica fechada: não ocupa espaço
            SetParent(h, project);
            // sem a área não-cliente (título "Propriedades - ..." e bordas): só o conteúdo
            SetWindowPos(h, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
            Layout();
        }

        // Ctrl+R (SageShortcuts, toggleExplorer): fecha a janela Projeto (as Propriedades vão
        // junto) ou abre e põe o foco na árvore
        public static void Toggle()
        {
            foreach (dynamic w in Vbe.Windows)
            {
                if ((int)w.Type != vbext_wt_ProjectWindow) continue;
                if ((bool)w.Visible && project != IntPtr.Zero && Native.IsWindowVisible(project)) w.Visible = false;
                else { w.Visible = true; w.SetFocus(); }
                return;
            }
        }

        // Exibir > Janela Propriedades: abre a seção e põe o foco nela
        public static void ShowProperties()
        {
            if (project == IntPtr.Zero) { try { dynamic w = PropertiesWindow(); if (w != null) w.Visible = true; } catch (Exception) { } return; }
            SetCollapsed(false);
            if (properties != IntPtr.Zero) SetFocus(properties);
        }

        static void SetCollapsed(bool value)
        {
            collapsed = value;
            Settings.Set(CollapsedKey, value ? "yes" : "no");
            Layout();
        }

        // Arrastar a faixa: y é a posição dela na janela Projeto
        static void Resize(int y)
        {
            Native.RECT r;
            Native.GetClientRect(project, out r);
            int h = r.Bottom - y - header.Height;
            height = h;
            if (collapsed && h > 40) collapsed = false;
            Layout();
        }

        static void SaveHeight()
        {
            if (height > 0) Settings.Set(HeightKey, height.ToString(CultureInfo.InvariantCulture));
        }

        // ------------------------------------------------------------------
        // Layout: árvore em cima, a faixa, as Propriedades embaixo
        // ------------------------------------------------------------------

        static bool laying;

        static void Layout()
        {
            if (project == IntPtr.Zero || header == null || laying) return;
            laying = true;
            try
            {
                Native.RECT r;
                Native.GetClientRect(project, out r);
                int width = r.Right, total = r.Bottom, bar = header.Height;
                int treeTop = TreeTop();
                int room = Math.Max(0, total - treeTop - bar);
                if (height < 0) height = room / 2;
                int props = collapsed || properties == IntPtr.Zero ? 0 : Math.Max(60, Math.Min(height, room - 60));
                int barTop = total - props - bar;
                header.Collapsed = collapsed;
                // A faixa fica na frente das Propriedades: a barra de título delas (que fica atrás
                // da faixa, fora da área visível) não pode ser desenhada por cima dela
                if (properties != IntPtr.Zero) SetWindowPos(properties, header.Handle, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                SetWindowPos(header.Handle, HWND_TOP, 0, barTop, width, bar, SWP_NOACTIVATE | SWP_SHOWWINDOW);
                if (tree != IntPtr.Zero) SetWindowPos(tree, IntPtr.Zero, 0, 0, width, Math.Max(0, barTop - treeTop), SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                if (properties != IntPtr.Zero && Native.GetParent(properties) == project)
                {
                    Rectangle p = PropertiesRect();
                    SetWindowPos(properties, IntPtr.Zero, p.X, p.Y, p.Width, p.Height, SWP_NOZORDER | SWP_NOACTIVATE | (collapsed ? SWP_HIDEWINDOW : SWP_SHOWWINDOW));
                }
                header.Invalidate();
            }
            finally { laying = false; }
        }

        // A área das Propriedades: embaixo da faixa até o fim da janela Projeto
        static Rectangle PropertiesRect()
        {
            Native.RECT r;
            Native.GetClientRect(project, out r);
            Rectangle header_ = HeaderRect();
            int y = header_.Bottom, h = Math.Max(0, r.Bottom - y);
            return new Rectangle(0, y, r.Right, h);
        }

        static Rectangle HeaderRect()
        {
            POINT o = new POINT();
            ClientToScreen(project, ref o);
            Native.RECT w;
            Native.GetWindowRect(header.Handle, out w);
            return new Rectangle(w.Left - o.X, w.Top - o.Y, w.Right - w.Left, w.Bottom - w.Top);
        }

        // Onde a árvore começa (abaixo da barra de botões da janela Projeto)
        static int TreeTop()
        {
            if (tree == IntPtr.Zero) return 0;
            Native.RECT w;
            Native.GetWindowRect(tree, out w);
            POINT o = new POINT();
            ClientToScreen(project, ref o);
            return w.Top - o.Y;
        }

        // ------------------------------------------------------------------
        // Subclassing
        // ------------------------------------------------------------------

        const int WM_SIZE = 0x0005, WM_WINDOWPOSCHANGING = 0x0046, WM_SETTEXT = 0x000C;
        const uint SWP_NOSIZE = 0x1, SWP_NOMOVE = 0x2, SWP_NOZORDER = 0x4, SWP_NOACTIVATE = 0x10, SWP_FRAMECHANGED = 0x20, SWP_SHOWWINDOW = 0x40, SWP_HIDEWINDOW = 0x80;
        const int WM_NCCALCSIZE = 0x0083, WM_NCPAINT = 0x0085, WM_NCACTIVATE = 0x0086, WM_NCHITTEST = 0x0084, HTCAPTION = 2;

        [StructLayout(LayoutKind.Sequential)]
        struct NCCALCSIZE_PARAMS { public Native.RECT r0, r1, r2; public IntPtr pos; }

        [DllImport("user32.dll")] static extern IntPtr GetWindowDC(IntPtr hwnd);
        [DllImport("user32.dll")] static extern int ReleaseDC(IntPtr hwnd, IntPtr dc);
        static readonly IntPtr HWND_TOP = IntPtr.Zero;

        [StructLayout(LayoutKind.Sequential)]
        struct WINDOWPOS { public IntPtr hwnd, hwndInsertAfter; public int x, y, cx, cy; public uint flags; }

        // A janela Projeto: depois de o VBE arrumar a árvore, a seção volta para baixo dela
        static IntPtr ProjectProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            // Barra de título do Sage no lugar da do VBE: só o título, sem o botão de fechar
            // (Ctrl+R abre e fecha a janela). Arrastar e o duplo clique nela seguem com o VBE
            // (HTCAPTION), para acoplar e desacoplar.
            if (hwnd == project && IsDocked())
            {
                switch (msg)
                {
                    case WM_NCCALCSIZE:
                        if (wParam != IntPtr.Zero)
                        {
                            NCCALCSIZE_PARAMS p = (NCCALCSIZE_PARAMS)Marshal.PtrToStructure(lParam, typeof(NCCALCSIZE_PARAMS));
                            Native.RECT whole = p.r0;
                            Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                            p = (NCCALCSIZE_PARAMS)Marshal.PtrToStructure(lParam, typeof(NCCALCSIZE_PARAMS));
                            p.r0 = whole;
                            p.r0.Top += TitleHeight;
                            Marshal.StructureToPtr(p, lParam, false);
                        }
                        return IntPtr.Zero;
                    case WM_NCPAINT:
                        PaintTitle(hwnd);
                        return IntPtr.Zero;
                    case WM_NCACTIVATE:
                        PaintTitle(hwnd);
                        return (IntPtr)1;
                    case WM_NCHITTEST:
                    {
                        Native.RECT w;
                        Native.GetWindowRect(hwnd, out w);
                        int y = (short)(((long)lParam >> 16) & 0xFFFF);
                        if (y >= w.Top && y < w.Top + TitleHeight) return (IntPtr)HTCAPTION;
                        break;
                    }
                    case WM_SETTEXT:
                    {
                        IntPtr r = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                        PaintTitle(hwnd);
                        return r;
                    }
                }
            }
            IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
            try
            {
                if (msg == WM_SIZE) Layout();
                else if (msg == Native.WM_NCDESTROY)
                {
                    Native.RemoveWindowSubclass(hwnd, projectProc, SubclassId);
                    if (hwnd == project) { project = IntPtr.Zero; tree = IntPtr.Zero; }
                }
            }
            catch (Exception ex) { Log.Error(ex); }
            return result;
        }

        // Acoplada (dentro da janela principal); flutuando, a moldura flutuante tem a barra dela
        static bool IsDocked()
        {
            return Native.GetAncestor(project, Native.GA_ROOT) == vbeWindow;
        }

        static int TitleHeight { get { return header != null ? header.Height : 24; } }

        static void PaintTitle(IntPtr hwnd)
        {
            Native.RECT w;
            if (!Native.GetWindowRect(hwnd, out w)) return;
            int width = w.Right - w.Left, height = TitleHeight;
            if (width <= 0) return;
            IntPtr dc = GetWindowDC(hwnd);
            if (dc == IntPtr.Zero) return;
            try
            {
                Theme t = Theme.Find(Settings.ColorTheme);
                using (Bitmap strip = new Bitmap(width, height))
                {
                    using (Graphics g = Graphics.FromImage(strip))
                    {
                        g.Clear(t.Sidebar);
                        StringBuilder sb = new StringBuilder(256);
                        GetWindowText(hwnd, sb, sb.Capacity);
                        Font font = SystemFonts.MessageBoxFont;
                        int pad = (int)Math.Round(height / 3.0);
                        TextRenderer.DrawText(g, sb.ToString(), font, new Rectangle(pad, 0, width - pad, height), t.Foreground, t.Sidebar,
                            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
                    }
                    using (Graphics target = Graphics.FromHdc(dc)) target.DrawImageUnscaled(strip, 0, 0);
                }
            }
            finally { ReleaseDC(hwnd, dc); }
        }

        // A árvore não passa da faixa da seção
        static IntPtr TreeProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            try
            {
                if (msg == WM_WINDOWPOSCHANGING && header != null && project != IntPtr.Zero && !laying)
                {
                    WINDOWPOS p = (WINDOWPOS)Marshal.PtrToStructure(lParam, typeof(WINDOWPOS));
                    if ((p.flags & SWP_NOSIZE) == 0)
                    {
                        int limit = HeaderRect().Top - ((p.flags & SWP_NOMOVE) == 0 ? p.y : TreeTop());
                        if (p.cy > limit) { p.cy = Math.Max(0, limit); Marshal.StructureToPtr(p, lParam, false); }
                    }
                }
                else if (msg == Native.WM_NCDESTROY)
                {
                    Native.RemoveWindowSubclass(hwnd, treeProc, SubclassId);
                    if (hwnd == tree) tree = IntPtr.Zero;
                }
            }
            catch (Exception ex) { Log.Error(ex); }
            return Native.DefSubclassProc(hwnd, msg, wParam, lParam);
        }

        // As Propriedades ficam na área da seção (e escondidas com ela recolhida)
        static IntPtr PropertiesProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            // Dentro da seção: sem área não-cliente (a faixa já diz o que é a janela)
            if (Native.GetParent(hwnd) == project && project != IntPtr.Zero)
            {
                if (msg == WM_NCCALCSIZE) return IntPtr.Zero; // área cliente = janela inteira
                if (msg == WM_NCPAINT) return IntPtr.Zero;
                if (msg == WM_NCACTIVATE) return (IntPtr)1;
            }
            try
            {
                if (msg == WM_WINDOWPOSCHANGING && Native.GetParent(hwnd) == project && header != null && !laying)
                {
                    WINDOWPOS p = (WINDOWPOS)Marshal.PtrToStructure(lParam, typeof(WINDOWPOS));
                    Rectangle r = PropertiesRect();
                    p.x = r.X; p.y = r.Y; p.cx = r.Width; p.cy = r.Height;
                    p.flags &= ~(SWP_NOMOVE | SWP_NOSIZE);
                    p.flags |= SWP_NOZORDER; // o VBE não a põe na frente da faixa
                    if (collapsed) p.flags &= ~SWP_SHOWWINDOW; else p.flags &= ~SWP_HIDEWINDOW;
                    Marshal.StructureToPtr(p, lParam, false);
                }
                else if (msg == WM_SETTEXT)
                {
                    IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                    if (header != null) header.UpdateTitle(Title(hwnd));
                    return result;
                }
                else if (msg == Native.WM_NCDESTROY)
                {
                    Native.RemoveWindowSubclass(hwnd, propertiesProc, SubclassId);
                    if (hwnd == properties) properties = IntPtr.Zero;
                }
            }
            catch (Exception ex) { Log.Error(ex); }
            return Native.DefSubclassProc(hwnd, msg, wParam, lParam);
        }

        // ------------------------------------------------------------------

        // Volta tudo para o VBE. showProperties: as Propriedades reaparecem acopladas onde o VBE
        // as deixava (configuração desligada, Excel fechando); false: ficam escondidas (a janela
        // Projeto foi fechada, e elas fecham junto)
        static void Release(bool showProperties)
        {
            if (project == IntPtr.Zero && properties == IntPtr.Zero && header == null) return;
            SaveHeight();
            if (properties != IntPtr.Zero && Native.IsWindow(properties))
            {
                Native.RemoveWindowSubclass(properties, propertiesProc, SubclassId);
                ShowWindow(properties, 0);
                SetParent(properties, vbeWindow);
                SetWindowPos(properties, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
                if (showProperties) try { dynamic w = PropertiesWindow(); if (w != null) w.Visible = true; } catch (Exception) { }
            }
            properties = IntPtr.Zero;
            if (tree != IntPtr.Zero && Native.IsWindow(tree)) Native.RemoveWindowSubclass(tree, treeProc, SubclassId);
            tree = IntPtr.Zero;
            if (header != null) { header.Dispose(); header = null; }
            if (project != IntPtr.Zero && Native.IsWindow(project))
            {
                Native.RemoveWindowSubclass(project, projectProc, SubclassId);
                // a barra de título do VBE volta
                SetWindowPos(project, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
                // o VBE arruma a janela Projeto de novo (a árvore volta a ocupar tudo)
                Native.RECT r;
                Native.GetClientRect(project, out r);
                SendMessage(project, WM_SIZE, IntPtr.Zero, (IntPtr)((r.Bottom << 16) | (r.Right & 0xFFFF)));
            }
            project = IntPtr.Zero;
        }

        public static void ApplyTheme()
        {
            if (header != null) header.ApplyTheme();
        }

        public static void Shutdown()
        {
            Release(true);
            if (menuClick != null) { menuClick.Dispose(); menuClick = null; }
        }

        // ------------------------------------------------------------------

        static dynamic PropertiesWindow()
        {
            foreach (dynamic w in Vbe.Windows)
                if ((int)w.Type == vbext_wt_PropertyWindow) return w;
            return null;
        }

        static IntPtr FindPropertiesHwnd()
        {
            IntPtr h = FindChild(vbeWindow, "wndclass_pbrs");
            if (h != IntPtr.Zero) return h;
            IntPtr found = IntPtr.Zero;
            Native.EnumThreadWindows(Native.GetCurrentThreadId(), delegate(IntPtr top, IntPtr l)
            {
                if (Native.ClassName(top) == "wndclass_pbrs") { found = top; return false; }
                found = FindChild(top, "wndclass_pbrs");
                return found == IntPtr.Zero;
            }, IntPtr.Zero);
            return found;
        }

        static IntPtr FindChild(IntPtr parent, string cls)
        {
            IntPtr found = IntPtr.Zero;
            Native.EnumChildWindows(parent, delegate(IntPtr h, IntPtr l)
            {
                if (Native.ClassName(h) == cls) { found = h; return false; }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        // "Propriedades - Módulo1" -> "Módulo1"
        static string Title(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return "";
            StringBuilder sb = new StringBuilder(256);
            GetWindowText(hwnd, sb, sb.Capacity);
            string t = sb.ToString();
            int p = t.IndexOf(" - ", StringComparison.Ordinal);
            return p >= 0 ? t.Substring(p + 3) : "";
        }

        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr hwnd, ref POINT p);
        [DllImport("user32.dll")] static extern IntPtr SetParent(IntPtr child, IntPtr parent);
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int cmd);
        [DllImport("user32.dll")] static extern IntPtr SetFocus(IntPtr hwnd);
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int m, IntPtr w, IntPtr l);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);
    }

    // A faixa "▾ PROPRIEDADES - Módulo1": clique recolhe ou abre; arrastar muda a altura
    sealed class ExplorerHeader : Control
    {
        public event Action Toggle;
        public event Action<int> Resized;
        public event Action DragDone;
        public bool Collapsed;
        string title = "";
        bool down, dragging, hot;
        int downY;
        float scale = 1;

        public ExplorerHeader()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Font = SystemFonts.MessageBoxFont;
            Cursor = Cursors.Hand;
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            using (Graphics g = CreateGraphics()) scale = g.DpiX / 96f;
            Height = (int)Math.Round(24 * scale);
        }

        public void UpdateTitle(string t)
        {
            if (t == title) return;
            title = t;
            Invalidate();
        }

        public void ApplyTheme() { Invalidate(); }

        protected override void OnPaint(PaintEventArgs e)
        {
            Theme t = Theme.Find(Settings.ColorTheme);
            Graphics g = e.Graphics;
            using (SolidBrush b = new SolidBrush(hot ? t.Hover : t.Sidebar)) g.FillRectangle(b, ClientRectangle);
            using (Pen line = new Pen(t.Border)) g.DrawLine(line, 0, 0, Width, 0);
            // seta: ▸ recolhida, ▾ aberta
            int cx = (int)Math.Round(12 * scale), cy = Height / 2, a = Math.Max(3, (int)Math.Round(4 * scale));
            Point[] arrow = Collapsed
                ? new[] { new Point(cx - a / 2, cy - a), new Point(cx + a / 2 + 1, cy), new Point(cx - a / 2, cy + a) }
                : new[] { new Point(cx - a, cy - a / 2), new Point(cx + a, cy - a / 2), new Point(cx, cy + a / 2 + 1) };
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using (SolidBrush b = new SolidBrush(t.Foreground)) g.FillPolygon(b, arrow);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
            const TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine |
                TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding;
            using (Font bold = new Font(Font, FontStyle.Bold))
            {
                string label = Strings.ExplorerProperties;
                int x = (int)Math.Round(22 * scale);
                // sem as reticências na medida: com elas e sem largura, o MeasureText erra
                int w = TextRenderer.MeasureText(g, label, bold, new Size(int.MaxValue, Height), TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding).Width;
                TextRenderer.DrawText(g, label, bold, new Rectangle(x, 0, Width - x, Height), t.Foreground, flags);
                if (title.Length > 0)
                    TextRenderer.DrawText(g, "  " + title, Font, new Rectangle(x + w, 0, Math.Max(0, Width - x - w), Height), t.Muted, flags);
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left) return;
            down = true; dragging = false; downY = e.Y;
            Capture = true;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (!down)
            {
                Cursor = Cursors.Hand;
                if (!hot) { hot = true; Invalidate(); }
                return;
            }
            if (!dragging && Math.Abs(e.Y - downY) > 3) { dragging = true; Cursor = Cursors.SizeNS; }
            if (dragging && Resized != null)
            {
                // posição da faixa na janela Projeto
                Resized(Top + e.Y - downY);
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (!down) return;
            down = false;
            Capture = false;
            if (!dragging && Toggle != null) Toggle();
            if (dragging && DragDone != null) DragDone();
            dragging = false;
            Cursor = Cursors.Hand;
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (hot) { hot = false; Invalidate(); }
        }
    }
}
