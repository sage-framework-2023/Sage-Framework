using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageEditor
{
    // Layout fixo do VBE, como o do VS Code: o Explorador (a janela Projeto) ocupa a altura inteira,
    // à esquerda ou à direita, e o Painel (a janela Terminal) fica embaixo do código, ao lado dele.
    // O VBE não tem comando para escolher o lado: a posição vem do arrasto e da ordem em que as
    // janelas foram acopladas. Elas são filhas diretas da janela principal (MDIClient, PROJECT,
    // GenericPane e os dois VBSlider); os pedidos do VBE para posicioná-las saem com a posição do
    // Sage (desvios em Overscroll.Hooks). Os divisores são arrastados pelo Sage (o VBSlider calcula
    // pela posição do VBE). As janelas não se desacoplam: não há título para arrastar, o item
    // Encaixável sai dos menus do VBE, e uma janela que fique solta (Ferramentas > Opções >
    // Encaixar) é acoplada de novo. Só com o Explorador unificado ligado.
    static class PanelLayout
    {
        public static dynamic Vbe;

        const string SideKey = "sage.layout.sidebar", WidthKey = "sage.layout.sidebarWidth", HeightKey = "sage.layout.panelHeight";
        const int SliderSize = 6, MinSide = 140, MinCode = 200, MinPanel = 80, MinCodeHeight = 100;
        const int DockableId = 746, HideId = 865, vbext_wt_ProjectWindow = 6;
        static readonly string[] menus = { "Project Window", "Project Window (Break)", "Property Browser", "Immediate Window", "Watch Window", "Locals Window", "Docked Window" };

        static IntPtr main, mdi, project, panel, sliderV, sliderH;
        static bool started, sideLeft, loaded, opening, openingPanel;
        static int sideWidth, panelHeight;
        static Rectangle divider, dividerH, panelBox;
        static readonly Dictionary<IntPtr, Rectangle> targets = new Dictionary<IntPtr, Rectangle>();
        static readonly HashSet<IntPtr> attached = new HashSet<IntPtr>();
        static readonly List<dynamic> hidden = new List<dynamic>();
        static readonly HashSet<string> done = new HashSet<string>();
        static readonly Native.SubclassProc sliderProc = SliderProc;
        static readonly UIntPtr SubclassId = (UIntPtr)0x5A75;
        static IntPtr dragging;
        static int tick;
        static ContextMenuStrip menu;

        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr hwnd, ref POINT p);
        [DllImport("user32.dll")] static extern bool ScreenToClient(IntPtr hwnd, ref POINT p);
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] static extern IntPtr SetCapture(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool ReleaseCapture();
        [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT p);
        [DllImport("user32.dll")] static extern IntPtr SetCursor(IntPtr cursor);
        [DllImport("user32.dll")] static extern bool LockWindowUpdate(IntPtr hwnd);

        const uint SWP_NOSIZE = 0x1, SWP_NOMOVE = 0x2, SWP_NOZORDER = 0x4, SWP_NOACTIVATE = 0x10;
        const int WM_SETCURSOR = 0x0020, WM_MOUSEMOVE = 0x0200, WM_LBUTTONDOWN = 0x0201, WM_LBUTTONUP = 0x0202, WM_LBUTTONDBLCLK = 0x0203,
            WM_CAPTURECHANGED = 0x0215;

        static bool Active { get { return started && Settings.UnifiedExplorer; } }

        public static void Start(IntPtr vbe)
        {
            main = vbe;
            started = true;
            HideDockable();
        }

        public static void Shutdown()
        {
            started = false;
            foreach (IntPtr h in attached) if (Native.IsWindow(h)) Native.RemoveWindowSubclass(h, sliderProc, SubclassId);
            attached.Clear();
            // A personalização dos menus fica gravada: o item volta
            foreach (dynamic c in hidden) { try { c.Visible = true; } catch (Exception) { } }
            hidden.Clear();
            if (menu != null) { menu.Dispose(); menu = null; }
        }

        // ------------------------------------------------------------------
        // Janelas e cálculo
        // ------------------------------------------------------------------

        static void Refresh()
        {
            mdi = project = panel = IntPtr.Zero;
            List<IntPtr> sliders = new List<IntPtr>();
            Native.EnumChildWindows(main, delegate(IntPtr h, IntPtr l)
            {
                if (Native.GetParent(h) != main) return true;
                string cls = Native.ClassName(h);
                if (cls == "MDIClient") mdi = h;
                else if (cls == "PROJECT") project = h;
                else if (cls == "GenericPane" && Native.IsWindowVisible(h)) panel = h;
                else if (cls == "VBSlider")
                {
                    sliders.Add(h);
                    if (attached.Add(h)) Native.SetWindowSubclass(h, sliderProc, SubclassId, UIntPtr.Zero);
                }
                return true;
            }, IntPtr.Zero);
            if (project != IntPtr.Zero && !opening && !Native.IsWindowVisible(project)) project = IntPtr.Zero;

            List<IntPtr> shown = sliders.FindAll(s => Native.IsWindowVisible(s));
            IntPtr vertical = IntPtr.Zero, horizontal = IntPtr.Zero;
            // Quem já é vertical continua vertical; com um só, ele é o do painel que estiver aberto
            foreach (IntPtr s in shown) if (s == sliderV) vertical = s;
            foreach (IntPtr s in shown) if (s == sliderH) horizontal = s;
            foreach (IntPtr s in shown)
            {
                if (s == vertical || s == horizontal) continue;
                Native.RECT r;
                Native.GetWindowRect(s, out r);
                if (vertical == IntPtr.Zero && project != IntPtr.Zero && (r.Bottom - r.Top > r.Right - r.Left || horizontal != IntPtr.Zero || panel == IntPtr.Zero)) vertical = s;
                else if (horizontal == IntPtr.Zero && panel != IntPtr.Zero) horizontal = s;
                else if (vertical == IntPtr.Zero && project != IntPtr.Zero) vertical = s;
            }
            sliderV = vertical;
            sliderH = horizontal;
        }

        static Rectangle ClientRectOf(IntPtr h)
        {
            Native.RECT r;
            Native.GetWindowRect(h, out r);
            POINT a = new POINT { X = r.Left, Y = r.Top };
            ScreenToClient(main, ref a);
            return new Rectangle(a.X, a.Y, r.Right - r.Left, r.Bottom - r.Top);
        }

        // A área das janelas acopladas: a janela principal sem as barras de comando
        static Rectangle Area()
        {
            Native.RECT c;
            Native.GetClientRect(main, out c);
            int left = 0, top = 0, right = c.Right, bottom = c.Bottom;
            Native.EnumChildWindows(main, delegate(IntPtr h, IntPtr l)
            {
                if (Native.GetParent(h) != main || Native.ClassName(h) != "MsoCommandBarDock" || !Native.IsWindowVisible(h)) return true;
                Rectangle r = ClientRectOf(h);
                if (r.Width <= 0 || r.Height <= 0) return true;
                if (r.Width >= r.Height) { if (r.Top <= c.Bottom / 2) top = Math.Max(top, r.Bottom); else bottom = Math.Min(bottom, r.Top); }
                else { if (r.Left <= c.Right / 2) left = Math.Max(left, r.Right); else right = Math.Min(right, r.Left); }
                return true;
            }, IntPtr.Zero);
            return Rectangle.FromLTRB(left, top, Math.Max(left, right), Math.Max(top, bottom));
        }

        // Tamanhos e lado: na primeira vez, os que o VBE já tem
        static void Load()
        {
            if (loaded || mdi == IntPtr.Zero) return;
            int value;
            string side = Settings.Get(SideKey, "");
            if (side.Length > 0) sideLeft = side == "left";
            else if (project != IntPtr.Zero) sideLeft = ClientRectOf(project).Left < ClientRectOf(mdi).Left;
            if (int.TryParse(Settings.Get(WidthKey, ""), NumberStyles.Integer, CultureInfo.InvariantCulture, out value)) sideWidth = value;
            else if (project != IntPtr.Zero) sideWidth = ClientRectOf(project).Width;
            if (int.TryParse(Settings.Get(HeightKey, ""), NumberStyles.Integer, CultureInfo.InvariantCulture, out value)) panelHeight = value;
            else if (panel != IntPtr.Zero) panelHeight = ClientRectOf(panel).Height;
            // Com o VBE ainda abrindo, a janela vem pequena: vale o tamanho padrão
            if (sideWidth < 200) sideWidth = 300;
            if (panelHeight < MinPanel) panelHeight = 260;
            loaded = true;
        }

        static void Compute()
        {
            Refresh();
            Load();
            targets.Clear();
            divider = dividerH = panelBox = Rectangle.Empty;
            if (mdi == IntPtr.Zero) return;
            Rectangle a = Area();
            if (a.Width <= 0 || a.Height <= 0) return;
            Rectangle code = a;
            if (project != IntPtr.Zero)
            {
                int gap = sliderV != IntPtr.Zero || opening ? SliderSize : 0;
                int width = Math.Max(Math.Min(MinSide, a.Width / 2), Math.Min(sideWidth, a.Width - gap - MinCode));
                if (sideLeft)
                {
                    targets[project] = new Rectangle(a.Left, a.Top, width, a.Height);
                    divider = new Rectangle(a.Left + width, a.Top, gap, a.Height);
                    code = Rectangle.FromLTRB(a.Left + width + gap, a.Top, a.Right, a.Bottom);
                }
                else
                {
                    targets[project] = new Rectangle(a.Right - width, a.Top, width, a.Height);
                    divider = new Rectangle(a.Right - width - gap, a.Top, gap, a.Height);
                    code = Rectangle.FromLTRB(a.Left, a.Top, a.Right - width - gap, a.Bottom);
                }
                if (sliderV != IntPtr.Zero) targets[sliderV] = divider;
            }
            if (panel != IntPtr.Zero || openingPanel)
            {
                int gap = sliderH != IntPtr.Zero || openingPanel ? SliderSize : 0;
                int height = Math.Max(Math.Min(MinPanel, code.Height / 2), Math.Min(panelHeight, code.Height - gap - MinCodeHeight));
                panelBox = new Rectangle(code.Left, code.Bottom - height, code.Width, height);
                dividerH = new Rectangle(code.Left, code.Bottom - height - gap, code.Width, gap);
                if (panel != IntPtr.Zero) targets[panel] = panelBox;
                if (sliderH != IntPtr.Zero) targets[sliderH] = dividerH;
                code = Rectangle.FromLTRB(code.Left, code.Top, code.Right, code.Bottom - height - gap);
            }
            targets[mdi] = code;
        }

        static void Apply()
        {
            Compute();
            // Em cópia: mover uma janela faz o VBE reposicionar as outras, e isso recalcula os alvos
            foreach (KeyValuePair<IntPtr, Rectangle> t in new List<KeyValuePair<IntPtr, Rectangle>>(targets))
            {
                Rectangle now = ClientRectOf(t.Key), r = t.Value;
                if (now != r) SetWindowPos(t.Key, IntPtr.Zero, r.X, r.Y, r.Width, r.Height, SWP_NOZORDER | SWP_NOACTIVATE);
            }
        }

        // Abre ou fecha a janela Projeto de uma vez: o VBE faz a troca em vários passos (o código
        // muda de posição e só depois de largura) e cada um seria pintado na tela. A travada
        // junta tudo numa pintura só. Se outra janela já estiver travada, segue sem travar.
        public static void Frozen(Action change)
        {
            bool locked = Active && LockWindowUpdate(main);
            try
            {
                change();
                if (Active) Apply();
            }
            finally { if (locked) LockWindowUpdate(IntPtr.Zero); }
        }

        // Abre a janela Projeto (UnifiedExplorer.Toggle): durante a abertura, os pedidos do VBE já
        // contam com ela, inclusive o do código, que sairia do tamanho antigo
        public static void Open(Action show)
        {
            opening = true;
            try { show(); }
            finally { opening = false; }
            if (Active) Apply();
        }

        // Janela oculta prestes a abrir: já vai para o lugar do Sage, senão o VBE a mostra na
        // posição dele e ela pisca lá. O VBSlider é do Painel ou do Explorador conforme o formato.
        static Rectangle PlaceOfClosed(IntPtr hwnd, string cls, bool tall)
        {
            bool ofPanel = cls == "GenericPane" || (cls == "VBSlider" && !tall);
            bool wasOpening = opening, wasPanel = openingPanel;
            if (ofPanel) openingPanel = true; else opening = true;
            try { Compute(); }
            finally { opening = wasOpening; openingPanel = wasPanel; }
            Rectangle r;
            if (cls == "VBSlider") return ofPanel ? dividerH : divider;
            if (cls == "GenericPane") return panelBox;
            return cls == "PROJECT" && targets.TryGetValue(hwnd, out r) ? r : Rectangle.Empty;
        }

        // O VBE recria o divisor a cada abertura, já visível e na posição dele: sem pedido para
        // desviar, a posição muda aqui, na criação (ThemeEngine.CbtProc, HCBT_CREATEWND). No
        // CREATESTRUCT, cy, cx, y e x vêm depois dos 4 ponteiros.
        public static void Created(IntPtr hwnd, IntPtr parent, IntPtr createStruct)
        {
            if (!Active || parent != main || Native.ClassName(hwnd) != "VBSlider") return;
            int at = 4 * IntPtr.Size;
            Rectangle r = PlaceOfClosed(hwnd, "VBSlider", Marshal.ReadInt32(createStruct, at) > Marshal.ReadInt32(createStruct, at + 4));
            if (r.IsEmpty) return;
            Marshal.WriteInt32(createStruct, at, r.Height);
            Marshal.WriteInt32(createStruct, at + 4, r.Width);
            Marshal.WriteInt32(createStruct, at + 8, r.Y);
            Marshal.WriteInt32(createStruct, at + 12, r.X);
        }

        // Desvios de posição (Overscroll): o pedido do VBE para uma das janelas sai com a posição do Sage
        public static void Adjust(IntPtr hwnd, ref int x, ref int y, ref int cx, ref int cy, ref uint flags)
        {
            if (!Active || Native.GetParent(hwnd) != main) return;
            try
            {
                string cls = Native.ClassName(hwnd);
                if (cls != "MDIClient" && cls != "PROJECT" && cls != "GenericPane" && cls != "VBSlider") return;
                Rectangle r;
                if (!Native.IsWindowVisible(hwnd))
                {
                    Rectangle now = ClientRectOf(hwnd);
                    r = PlaceOfClosed(hwnd, cls, (flags & SWP_NOSIZE) != 0 ? now.Height > now.Width : cy > cx);
                    if (r.IsEmpty) return;
                }
                else
                {
                    Compute();
                    if (!targets.TryGetValue(hwnd, out r)) return;
                }
                x = r.X; y = r.Y; cx = r.Width; cy = r.Height;
                flags &= ~(SWP_NOMOVE | SWP_NOSIZE);
            }
            catch (Exception) { } // nunca atrapalhar o VBE
        }

        // ------------------------------------------------------------------
        // Verificação periódica
        // ------------------------------------------------------------------

        public static void Poll()
        {
            if (!Active || dragging != IntPtr.Zero) return;
            Apply();
            if (++tick % 7 == 0) Redock();
            if (tick % 20 == 0 && done.Count < menus.Length) HideDockable();
        }

        // Janela solta (Ferramentas > Opções > Encaixar ou de antes desta versão): volta a acoplar
        static void Redock()
        {
            try
            {
                foreach (dynamic w in Vbe.Windows)
                    if ((int)w.Type == vbext_wt_ProjectWindow && (bool)w.Visible && project == IntPtr.Zero) Vbe.MainWindow.LinkedWindows.Add(w);
                TerminalWindow.Redock();
            }
            catch (Exception) { } // Encaixável desligado nas Opções, ou o VBE ocupado
        }

        // O Encaixável é o item logo antes do Ocultar: o Id 746 também é o de Remover e Fechar
        // projeto, e o texto muda com o idioma. Alguns menus só têm os itens depois de abertos: a
        // verificação periódica volta aos que faltam.
        static void HideDockable()
        {
            if (!Settings.UnifiedExplorer) return;
            foreach (string name in menus)
            {
                if (done.Contains(name)) continue;
                try
                {
                    dynamic controls = Vbe.CommandBars[name].Controls;
                    int count = controls.Count;
                    for (int i = 1; i < count; i++)
                    {
                        dynamic c = controls[i];
                        if ((int)c.Id != DockableId || (int)controls[i + 1].Id != HideId) continue;
                        if ((bool)c.Visible) { c.Visible = false; hidden.Add(c); }
                        done.Add(name);
                        break;
                    }
                }
                catch (Exception ex) { Log.Error(ex); }
            }
        }

        // ------------------------------------------------------------------
        // Divisores
        // ------------------------------------------------------------------

        static IntPtr SliderProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            try
            {
                if (msg == Native.WM_NCDESTROY)
                {
                    Native.RemoveWindowSubclass(hwnd, sliderProc, SubclassId);
                    attached.Remove(hwnd);
                }
                else if (Active)
                {
                    switch (msg)
                    {
                        case WM_SETCURSOR:
                            SetCursor((hwnd == sliderH ? Cursors.HSplit : Cursors.VSplit).Handle);
                            return (IntPtr)1;
                        case WM_LBUTTONDOWN:
                        case WM_LBUTTONDBLCLK:
                            dragging = hwnd;
                            SetCapture(hwnd);
                            return IntPtr.Zero;
                        case WM_MOUSEMOVE:
                            if (dragging == hwnd) Drag(hwnd);
                            return IntPtr.Zero;
                        case WM_LBUTTONUP:
                            if (dragging == hwnd) { dragging = IntPtr.Zero; ReleaseCapture(); Save(); }
                            return IntPtr.Zero;
                        case WM_CAPTURECHANGED:
                            if (dragging == hwnd) { dragging = IntPtr.Zero; Save(); }
                            return IntPtr.Zero;
                    }
                }
            }
            catch (Exception ex) { Log.Error(ex); }
            return Native.DefSubclassProc(hwnd, msg, wParam, lParam);
        }

        static void Drag(IntPtr slider)
        {
            POINT p;
            GetCursorPos(out p);
            ScreenToClient(main, ref p);
            Rectangle a = Area();
            if (slider == sliderH)
            {
                panelHeight = Math.Max(MinPanel, a.Bottom - p.Y - SliderSize / 2);
            }
            else
            {
                sideWidth = Math.Max(MinSide, (sideLeft ? p.X - a.Left : a.Right - p.X) - SliderSize / 2);
            }
            Apply();
        }

        static void Save()
        {
            Settings.Set(WidthKey, sideWidth.ToString(CultureInfo.InvariantCulture));
            Settings.Set(HeightKey, panelHeight.ToString(CultureInfo.InvariantCulture));
        }

        // ------------------------------------------------------------------
        // Menus
        // ------------------------------------------------------------------

        static void SetSide(bool left)
        {
            sideLeft = left;
            Settings.Set(SideKey, left ? "left" : "right");
            Apply();
        }

        static ContextMenuStrip NewMenu()
        {
            if (menu != null) menu.Dispose();
            Theme t = Theme.Find(Settings.ColorTheme);
            menu = new ContextMenuStrip();
            menu.Renderer = new ToolStripProfessionalRenderer(new EditorTabs.MenuColors(t));
            menu.BackColor = t.Sidebar;
            menu.ForeColor = t.Foreground;
            menu.ShowImageMargin = false;
            return menu;
        }

        static void AddItem(string text, string shortcut, EventHandler click)
        {
            ToolStripMenuItem item = new ToolStripMenuItem(text, null, click);
            item.ForeColor = menu.ForeColor;
            if (shortcut != null) item.ShortcutKeyDisplayString = shortcut;
            menu.Items.Add(item);
        }

        // Clique direito na barra do Explorador
        public static void ShowExplorerMenu()
        {
            NewMenu();
            AddItem(sideLeft ? Strings.MoveSideBarRight : Strings.MoveSideBarLeft, null, delegate { SetSide(!sideLeft); });
            AddItem(Strings.HideSideBar, "Ctrl+R", delegate { UnifiedExplorer.Toggle(); });
            menu.Show(Cursor.Position);
        }

        // Clique direito na faixa das abas do Painel
        public static void ShowPanelMenu()
        {
            NewMenu();
            AddItem(Strings.HidePanel, "Ctrl+J", delegate { TerminalWindow.Hide(); });
            menu.Show(Cursor.Position);
        }
    }
}
