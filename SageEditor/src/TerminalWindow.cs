using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SageEditor
{
    // Janela "Terminal": uma janela acoplável do VBE (Windows.CreateToolWindow) com três abas:
    //   Imediata                a Verificação imediata de verdade, trazida para dentro da aba
    //   Variáveis Locais        a janela Variáveis locais de verdade, idem
    //   Inspeção de Variáveis   a Inspeção de variáveis de verdade, idem
    //   Terminal                um PowerShell dentro do Excel (PsConsole)
    //   DataFrame Results       o df.Show do VBA, como o Query Results do SQL (ResultsView)
    //
    // As duas janelas do VBE continuam sendo dele (depuração, Debug.Print, F8...). Para o VBE
    // elas ficam "fechadas" (Visible = False), então ele não reserva espaço nem desenha
    // divisórias para elas; a janela Win32 (VbaWindow) passa a ser filha da área das abas.
    // Um subclassing impede o VBE de movê-la ou de escondê-la enquanto a aba dela está ativa,
    // e a verificação periódica a traz de volta se o VBE a reabrir (ex.: Adicionar inspeção).
    //
    // Ctrl+G (Verificação imediata) fica desativado no VBE; Exibir > Verificação imediata e
    // Exibir > Inspeção de variáveis abrem a aba correspondente. O Ctrl+J do SageShortcuts
    // abre e fecha a janela (Connect.ToggleTerminal).
    static class TerminalWindow
    {
        public const string HostProgId = "Sage.TerminalHost";
        const string PositionGuid = "{DF6B0202-ED1E-4931-B770-80AF9A506B4D}"; // o VBE guarda a posição por ele
        public const int TabImmediate = 0, TabLocals = 1, TabWatch = 2, TabConsole = 3, TabResults = 4;
        const int vbext_wt_Watch = 3, vbext_wt_Locals = 4, vbext_wt_Immediate = 5;
        // Exibir > Verificação imediata / Variáveis locais / Inspeção de variáveis
        const int ImmediateId = 2554, LocalsId = 2555, WatchId = 2556;

        public static dynamic Vbe;
        static object addIn;
        static IntPtr vbeWindow;
        static dynamic toolWindow;
        static TerminalHost host;
        static bool failed;
        static int tick;
        static readonly List<ButtonClick> menuClicks = new List<ButtonClick>();

        sealed class Borrowed
        {
            public readonly int Type, Tab;
            public IntPtr Hwnd;
            public bool Once; // já esteve na aba: um pedido do VBE para mostrá-la abre a aba
            // Borda e barra de título da janela do VBE: ficam fora da área da aba
            public int Left, Top, Right, Bottom;
            public Borrowed(int type, int tab) { Type = type; Tab = tab; }
        }

        static readonly Borrowed[] borrowed = { new Borrowed(vbext_wt_Immediate, TabImmediate), new Borrowed(vbext_wt_Locals, TabLocals), new Borrowed(vbext_wt_Watch, TabWatch) };
        static readonly Native.SubclassProc borrowedProc = BorrowedProc;
        static readonly UIntPtr SubclassId = (UIntPtr)0x5A72;

        const int WH_KEYBOARD = 2, HC_ACTION = 0, VK_G = 0x47, VK_SHIFT = 0x10, VK_CONTROL = 0x11, VK_MENU = 0x12;
        static readonly Native.HookProc keyboardProc = KeyboardProc;
        static IntPtr keyboardHook;

        const int WM_WINDOWPOSCHANGING = 0x0046;
        const uint SWP_NOSIZE = 0x1, SWP_NOMOVE = 0x2, SWP_NOZORDER = 0x4, SWP_NOACTIVATE = 0x10, SWP_FRAMECHANGED = 0x20,
            SWP_SHOWWINDOW = 0x40, SWP_HIDEWINDOW = 0x80;
        const int SW_HIDE = 0, SW_SHOW = 5;

        [StructLayout(LayoutKind.Sequential)]
        struct WINDOWPOS { public IntPtr hwnd, hwndInsertAfter; public int x, y, cx, cy; public uint flags; }

        [DllImport("user32.dll")] static extern IntPtr SetParent(IntPtr child, IntPtr parent);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int cmd);
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] static extern IntPtr SetFocus(IntPtr hwnd);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern short GetKeyState(int vk);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);

        public static bool Ready { get { return host != null; } }

        public static void Start(IntPtr vbe, object addInInstance)
        {
            vbeWindow = vbe;
            addIn = addInInstance;
            keyboardHook = Native.SetWindowsHookEx(WH_KEYBOARD, keyboardProc, IntPtr.Zero, Native.GetCurrentThreadId());
        }

        // ------------------------------------------------------------------
        // Verificação periódica (thread de interface)
        // ------------------------------------------------------------------

        public static void Poll()
        {
            if (Vbe == null || failed) return;
            if (host == null) { Create(); return; }
            FitHost();
            if (++tick % 3 != 0) return;
            foreach (Borrowed b in borrowed)
            {
                dynamic w = WindowOfType(b.Type);
                if (w == null) continue;
                bool vbeShows = (bool)w.Visible;
                if (vbeShows || b.Hwnd == IntPtr.Zero || !Native.IsWindow(b.Hwnd) || Native.GetParent(b.Hwnd) != host.Panel.Content.Handle)
                {
                    bool wasHere = b.Once;
                    Borrow(b, w);
                    // O VBE pediu para mostrar a janela (ex.: Depurar > Adicionar inspeção)
                    if (vbeShows && wasHere) Show(b.Tab);
                }
            }
        }

        static void Create()
        {
            try
            {
                // Pela reflexão: o binder do dynamic não converte o AddIn (1º argumento), e o
                // último (DocObj) volta por referência
                object windows = Vbe.Windows;
                object[] args = { new DispatchWrapper(addIn), HostProgId, Strings.TerminalTitle, PositionGuid, new DispatchWrapper(null) };
                ParameterModifier byRef = new ParameterModifier(5);
                byRef[4] = true;
                toolWindow = windows.GetType().InvokeMember("CreateToolWindow", BindingFlags.InvokeMethod, null, windows,
                    args, new ParameterModifier[] { byRef }, null, null);
                object doc = args[4];
                host = doc as TerminalHost;
                if (host == null)
                {
                    failed = true;
                    Log.Info("Terminal: o controle " + HostProgId + " não foi carregado (" + (doc == null ? "nulo" : doc.GetType().FullName) + ")");
                    return;
                }
                host.Panel.TabChanged += delegate { LayoutBorrowed(); };
                host.Panel.Content.Resize += delegate { LayoutBorrowed(); };
                InterceptMenus();
                DockFirstTime();
                AppDomain.CurrentDomain.SetData(ResultsSlot, showResults);
                Log.Info("Terminal criado");
            }
            catch (Exception ex)
            {
                failed = true;
                Log.Error(ex);
            }
        }

        // Na primeira vez, o VBE cria a janela flutuando, pequena: ela vai para baixo do código,
        // acoplada, como a Verificação imediata. Depois o VBE lembra onde o usuário a deixou.
        const string DockedKey = "sage.terminal.docked";

        static void DockFirstTime()
        {
            if (Settings.Get(DockedKey, "") == "yes") return;
            try { toolWindow.Visible = true; } catch (Exception ex) { Log.Error(ex); }
            try { toolWindow.Height = 260; } catch (COMException) { } // só vale flutuando
            try { Vbe.MainWindow.LinkedWindows.Add(toolWindow); Settings.Set(DockedKey, "yes"); }
            catch (Exception ex) { Log.Error(ex); }
            try { toolWindow.Visible = false; } catch (Exception ex) { Log.Error(ex); }
        }

        // O VBE redimensiona a moldura da janela, mas não o controle dentro dela: o controle
        // acompanha a área da moldura (que muda ao acoplar, desacoplar e arrastar a divisória)
        static IntPtr frame;
        static bool wasFloating;
        static readonly Native.SubclassProc frameProc = FrameProc;
        const int WM_SIZE = 0x0005;

        static void FitHost()
        {
            if (host == null || !host.IsHandleCreated) return;
            IntPtr parent = Native.GetParent(host.Handle);
            if (parent != frame)
            {
                if (frame != IntPtr.Zero && Native.IsWindow(frame)) Native.RemoveWindowSubclass(frame, frameProc, SubclassId);
                frame = parent;
                if (frame != IntPtr.Zero)
                {
                    Native.SetWindowSubclass(frame, frameProc, SubclassId, UIntPtr.Zero);
                    // Recalcula a área não-cliente (some a barra do VBE)
                    SetWindowPos(frame, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
                    wasFloating = false;
                }
            }
            if (frame == IntPtr.Zero) return;
            bool floating = Native.GetAncestor(frame, Native.GA_ROOT) != vbeWindow;
            if (floating && !wasFloating) FloatingSize();
            wasFloating = floating;
            Native.RECT r;
            Native.GetClientRect(frame, out r);
            int width = r.Right - r.Left, height = r.Bottom - r.Top;
            if (width > 0 && height > 0 && (host.Left != 0 || host.Top != 0 || host.Width != width || host.Height != height))
            {
                SetWindowPos(host.Handle, IntPtr.Zero, 0, 0, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
                Repaint();
            }
        }

        // A barra que o VBE desenhou antes de a área não-cliente sumir fica na tela até o
        // conteúdo se pintar por cima: repinta tudo (abas e janelas emprestadas)
        const uint RDW_INVALIDATE = 0x1, RDW_ALLCHILDREN = 0x80, RDW_UPDATENOW = 0x100, RDW_FRAME = 0x400;
        [DllImport("user32.dll")] static extern bool RedrawWindow(IntPtr hwnd, IntPtr rect, IntPtr region, uint flags);

        static void Repaint()
        {
            if (host != null && host.IsHandleCreated)
                RedrawWindow(host.Handle, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW | RDW_FRAME);
        }

        // Ao desacoplar, o VBE abre a janela flutuante com 200 x 200: pequena demais para as abas
        static void FloatingSize()
        {
            if (frame == IntPtr.Zero || Native.GetAncestor(frame, Native.GA_ROOT) == vbeWindow) return;
            Native.RECT r;
            Native.GetWindowRect(frame, out r);
            if (r.Right - r.Left >= 400) return;
            try
            {
                toolWindow.Width = 700;
                toolWindow.Height = 300;
            }
            catch (COMException) { }
        }

        // A moldura acoplada (GenericPane) desenha a barra "Terminal ×" na área não-cliente
        // dela. Sem área não-cliente, as abas ficam no topo, com o × delas; arrastar a área
        // vazia das abas faz o que a barra fazia (StartDrag).
        const int WM_NCCALCSIZE = 0x0083, WM_NCPAINT = 0x0085, WM_NCACTIVATE = 0x0086, WM_PAINT = 0x000F, WM_ERASEBKGND = 0x0014,
            WM_SETFOCUS = 0x0007, WM_KILLFOCUS = 0x0008, WM_SETTEXT = 0x000C, WM_SHOWWINDOW = 0x0018;

        [StructLayout(LayoutKind.Sequential)]
        struct NCCALCSIZE_PARAMS { public Native.RECT r0, r1, r2; public IntPtr pos; }

        static IntPtr FrameProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            if (msg == WM_NCCALCSIZE)
            {
                // A área cliente passa a ser a janela inteira
                if (wParam != IntPtr.Zero)
                {
                    NCCALCSIZE_PARAMS p = (NCCALCSIZE_PARAMS)Marshal.PtrToStructure(lParam, typeof(NCCALCSIZE_PARAMS));
                    Native.RECT whole = p.r0;
                    Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                    p = (NCCALCSIZE_PARAMS)Marshal.PtrToStructure(lParam, typeof(NCCALCSIZE_PARAMS));
                    p.r0 = whole;
                    Marshal.StructureToPtr(p, lParam, false);
                }
                return IntPtr.Zero;
            }
            if (msg == WM_NCPAINT) return IntPtr.Zero;
            if (msg == WM_NCACTIVATE) return (IntPtr)1; // sem redesenhar a barra
            // O conteúdo (as abas) cobre a moldura inteira: ela não precisa pintar nada
            if (msg == WM_PAINT)
            {
                ValidateRect(hwnd, IntPtr.Zero);
                return IntPtr.Zero;
            }
            if (msg == WM_ERASEBKGND) return (IntPtr)1;

            IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
            try
            {
                if (msg == WM_SIZE) FitHost();
                else if (msg == WM_SETFOCUS || msg == WM_KILLFOCUS || msg == WM_SHOWWINDOW || msg == WM_SETTEXT) Repaint();
                else if (msg == Native.WM_NCDESTROY)
                {
                    Native.RemoveWindowSubclass(hwnd, frameProc, SubclassId);
                    if (hwnd == frame) frame = IntPtr.Zero;
                }
            }
            catch (Exception ex) { Log.Error(ex); }
            return result;
        }

        // df.Show do VBA (SageTypes): mostra o DataFrame na aba DataFrame Results. O SageTypes
        // acha esta função no AppDomain (os dois rodam no processo do Excel, sem referência um
        // ao outro); false: ele usa a janela própria.
        const string ResultsSlot = "Sage.DataFrameResults";
        static readonly Func<string, string[], string[], long, Func<long, int, string[][]>, bool> showResults = ShowResults;

        static bool ShowResults(string title, string[] columns, string[] types, long rows, Func<long, int, string[][]> fetch)
        {
            if (!Ready || host.InvokeRequired) return false;
            host.Panel.Results.ShowFrame(title, columns, types, rows, fetch);
            Show(TabResults);
            return true;
        }

        // Exibir > Verificação imediata / Inspeção de variáveis (e os botões iguais nas
        // barras): abrem a aba, no lugar da janela solta do VBE
        static void InterceptMenus()
        {
            foreach (int id in new int[] { ImmediateId, LocalsId, WatchId })
            {
                int tab = id == ImmediateId ? TabImmediate : id == LocalsId ? TabLocals : TabWatch;
                dynamic controls = Vbe.CommandBars.FindControls(Type.Missing, id);
                if (controls == null) continue;
                foreach (dynamic control in controls)
                {
                    Action open = delegate { Show(tab); };
                    try { menuClicks.Add(new ButtonClick((object)control, open, true)); }
                    catch (Exception ex) { Log.Error(ex); }
                }
            }
        }

        static dynamic WindowOfType(int type)
        {
            foreach (dynamic w in Vbe.Windows)
                if ((int)w.Type == type) return w;
            return null;
        }

        // ------------------------------------------------------------------
        // Janelas do VBE dentro das abas
        // ------------------------------------------------------------------

        static void Borrow(Borrowed b, dynamic w)
        {
            string caption = w.Caption;
            IntPtr h = FindPane(caption);
            if (h == IntPtr.Zero)
            {
                // Nunca foi aberta nesta sessão: o VBE cria a janela ao mostrá-la
                w.Visible = true;
                h = FindPane(caption);
                if (h == IntPtr.Zero) return;
            }
            if (h != b.Hwnd)
            {
                if (b.Hwnd != IntPtr.Zero && Native.IsWindow(b.Hwnd)) Native.RemoveWindowSubclass(b.Hwnd, borrowedProc, SubclassId);
                b.Hwnd = h;
                Native.SetWindowSubclass(h, borrowedProc, SubclassId, UIntPtr.Zero);
            }
            // Para o VBE ela fica fechada: não reserva espaço no layout dele
            if ((bool)w.Visible) w.Visible = false;
            SetParent(h, host.Panel.Content.Handle);
            MeasureFrame(b);
            b.Once = true;
            LayoutBorrowed();
        }

        // Distância da janela até a área cliente dela, em cada lado (barra de título "Verificação
        // imediata ×" e bordas): a janela fica maior que a aba, e só o conteúdo aparece
        static void MeasureFrame(Borrowed b)
        {
            Native.RECT window, client;
            if (!Native.GetWindowRect(b.Hwnd, out window) || !Native.GetClientRect(b.Hwnd, out client)) return;
            POINT origin = new POINT();
            ClientToScreen(b.Hwnd, ref origin);
            b.Left = Math.Max(0, origin.X - window.Left);
            b.Top = Math.Max(0, origin.Y - window.Top);
            b.Right = Math.Max(0, window.Right - (origin.X + client.Right));
            b.Bottom = Math.Max(0, window.Bottom - (origin.Y + client.Bottom));
        }

        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr hwnd, ref POINT p);

        static Rectangle Placement(Borrowed b)
        {
            Size size = host.Panel.Content.ClientSize;
            return new Rectangle(-b.Left, -b.Top, size.Width + b.Left + b.Right, size.Height + b.Top + b.Bottom);
        }

        static IntPtr FindPane(string caption)
        {
            IntPtr found = IntPtr.Zero;
            Native.EnumWindowsProc match = delegate(IntPtr h, IntPtr l)
            {
                if (Native.ClassName(h) == "VbaWindow" && Title(h) == caption) { found = h; return false; }
                return true;
            };
            Native.EnumChildWindows(vbeWindow, match, IntPtr.Zero);
            if (found == IntPtr.Zero)
                Native.EnumThreadWindows(Native.GetCurrentThreadId(), delegate(IntPtr top, IntPtr l)
                {
                    if (found == IntPtr.Zero && top != vbeWindow) Native.EnumChildWindows(top, match, IntPtr.Zero);
                    return found == IntPtr.Zero;
                }, IntPtr.Zero);
            return found;
        }

        static string Title(IntPtr hwnd)
        {
            StringBuilder sb = new StringBuilder(256);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        static Borrowed FindBorrowed(IntPtr hwnd)
        {
            foreach (Borrowed b in borrowed) if (b.Hwnd == hwnd) return b;
            return null;
        }

        static void LayoutBorrowed()
        {
            if (host == null) return;
            int active = host.Panel.ActiveTab;
            foreach (Borrowed b in borrowed)
            {
                if (b.Hwnd == IntPtr.Zero || Native.GetParent(b.Hwnd) != host.Panel.Content.Handle) continue;
                Rectangle r = Placement(b);
                SetWindowPos(b.Hwnd, IntPtr.Zero, r.X, r.Y, r.Width, r.Height, SWP_NOZORDER | SWP_NOACTIVATE);
                ShowWindow(b.Hwnd, b.Tab == active ? SW_SHOW : SW_HIDE);
            }
        }

        // O VBE continua mandando na janela dele: posição e visibilidade ficam as da aba
        static IntPtr BorrowedProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            try
            {
                if (msg == WM_WINDOWPOSCHANGING && host != null && Native.GetParent(hwnd) == host.Panel.Content.Handle)
                {
                    Borrowed b = FindBorrowed(hwnd);
                    if (b != null)
                    {
                        WINDOWPOS p = (WINDOWPOS)Marshal.PtrToStructure(lParam, typeof(WINDOWPOS));
                        Rectangle r = Placement(b);
                        p.x = r.X; p.y = r.Y; p.cx = r.Width; p.cy = r.Height;
                        p.flags &= ~(SWP_NOMOVE | SWP_NOSIZE);
                        if (b.Tab == host.Panel.ActiveTab) p.flags &= ~SWP_HIDEWINDOW;
                        else p.flags &= ~SWP_SHOWWINDOW;
                        Marshal.StructureToPtr(p, lParam, false);
                    }
                }
                else if (msg == Native.WM_NCDESTROY)
                {
                    Native.RemoveWindowSubclass(hwnd, borrowedProc, SubclassId);
                    Borrowed b = FindBorrowed(hwnd);
                    if (b != null) b.Hwnd = IntPtr.Zero;
                }
            }
            catch (Exception ex) { Log.Error(ex); }
            return Native.DefSubclassProc(hwnd, msg, wParam, lParam);
        }

        // ------------------------------------------------------------------
        // Abrir e fechar
        // ------------------------------------------------------------------

        public static void Toggle()
        {
            if (!Ready) return;
            if ((bool)toolWindow.Visible) toolWindow.Visible = false;
            else Show(-1);
        }

        // Fecha se a aba já está à vista; senão abre nela
        public static void ToggleTab(int tab)
        {
            if (!Ready) return;
            if ((bool)toolWindow.Visible && host.Panel.ActiveTab == tab) toolWindow.Visible = false;
            else Show(tab);
        }

        public static void Show(int tab)
        {
            if (!Ready) return;
            if (tab >= 0) host.Panel.SelectTab(tab);
            toolWindow.Visible = true;
            FitHost();
            Repaint();
            try { toolWindow.SetFocus(); } catch (COMException) { }
            FocusActive();
        }

        static void FocusActive()
        {
            int active = host.Panel.ActiveTab;
            if (active == TabConsole) { host.Panel.Console.FocusInput(); return; }
            if (active == TabResults) { host.Panel.Results.FocusGrid(); return; }
            foreach (Borrowed b in borrowed)
                if (b.Tab == active && b.Hwnd != IntPtr.Zero) SetFocus(b.Hwnd);
        }

        // Arrastar a área vazia das abas: o VBE recebe o clique como se fosse na barra de título
        // dele (acoplar, desacoplar, mover); o duplo clique alterna entre acoplada e flutuante
        const int WM_NCLBUTTONDOWN = 0x00A1, WM_NCLBUTTONDBLCLK = 0x00A3, HTCAPTION = 2;
        [DllImport("user32.dll")] static extern bool ReleaseCapture();
        [DllImport("user32.dll")] static extern bool ValidateRect(IntPtr hwnd, IntPtr rect);
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

        public static void StartDrag(Point screen, bool doubleClick)
        {
            if (frame == IntPtr.Zero) return;
            ReleaseCapture();
            // Flutuando, quem tem a barra de título é a janela flutuante do VBE
            IntPtr root = Native.GetAncestor(frame, Native.GA_ROOT);
            IntPtr target = root != IntPtr.Zero && root != vbeWindow ? root : frame;
            IntPtr point = (IntPtr)((screen.Y << 16) | (screen.X & 0xFFFF));
            SendMessage(target, doubleClick ? WM_NCLBUTTONDBLCLK : WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, point);
        }

        public static void Hide()
        {
            if (Ready) toolWindow.Visible = false;
        }

        public static void ApplyTheme()
        {
            if (host != null) host.Panel.ApplyTheme();
        }

        // ------------------------------------------------------------------
        // Ctrl+G desativado no VBE
        // ------------------------------------------------------------------

        static IntPtr KeyboardProc(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (code == HC_ACTION && (int)wParam == VK_G && GetKeyState(VK_CONTROL) < 0 &&
                    GetKeyState(VK_MENU) >= 0 && GetKeyState(VK_SHIFT) >= 0 && VbeInFront())
                    return (IntPtr)1;
            }
            catch (Exception ex) { Log.Error(ex); }
            return Native.CallNextHookEx(keyboardHook, code, wParam, lParam);
        }

        // O VBE (ou uma janela dele, como o Terminal flutuante) em primeiro plano; no
        // Excel, o Ctrl+G continua sendo o Ir para
        static bool VbeInFront()
        {
            IntPtr front = GetForegroundWindow();
            return front != IntPtr.Zero && (front == vbeWindow || Native.GetAncestor(front, Native.GA_ROOTOWNER) == vbeWindow);
        }

        // ------------------------------------------------------------------

        public static void Shutdown()
        {
            if (keyboardHook != IntPtr.Zero) { Native.UnhookWindowsHookEx(keyboardHook); keyboardHook = IntPtr.Zero; }
            AppDomain.CurrentDomain.SetData(ResultsSlot, null);
            foreach (ButtonClick c in menuClicks) c.Dispose();
            menuClicks.Clear();
            if (frame != IntPtr.Zero && Native.IsWindow(frame))
            {
                Native.RemoveWindowSubclass(frame, frameProc, SubclassId);
                SetWindowPos(frame, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
            }
            frame = IntPtr.Zero;
            // As janelas do VBE voltam para ele (fechadas, como ele as vê) antes de ele salvar o layout
            foreach (Borrowed b in borrowed)
            {
                if (b.Hwnd != IntPtr.Zero && Native.IsWindow(b.Hwnd))
                {
                    Native.RemoveWindowSubclass(b.Hwnd, borrowedProc, SubclassId);
                    ShowWindow(b.Hwnd, SW_HIDE);
                    SetParent(b.Hwnd, vbeWindow);
                }
                b.Hwnd = IntPtr.Zero;
            }
            if (host != null) host.Panel.Console.Shutdown();
            host = null;
            toolWindow = null;
            addIn = null;
            failed = false;
        }
    }

    // O controle ActiveX que o VBE hospeda na janela acoplável
    [ComVisible(true), Guid("946CA9EF-210D-4F1E-A571-6129E9967700"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITerminalHost
    {
    }

    [ComVisible(true), Guid("31E81384-D124-4B77-A222-BEE1ACED5A35"), ProgId(TerminalWindow.HostProgId),
     ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(ITerminalHost))]
    public class TerminalHost : UserControl, ITerminalHost
    {
        internal readonly TerminalPanel Panel;

        public TerminalHost()
        {
            Panel = new TerminalPanel();
            Panel.Dock = DockStyle.Fill;
            Controls.Add(Panel);
        }
    }

    // As abas (desenhadas como as abas da área de código) e a área de conteúdo
    sealed class TerminalPanel : Control
    {
        public readonly Panel Content = new Panel();
        public readonly PsConsole Console = new PsConsole();
        public readonly ResultsView Results = new ResultsView();
        public event Action<int> TabChanged;
        public int ActiveTab { get; private set; }

        readonly Rectangle[] tabs = new Rectangle[5];
        int hot = -1;
        float scale = 1;

        public TerminalPanel()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Font = SystemFonts.MessageBoxFont;
            Console.Dock = DockStyle.Fill;
            Console.Visible = false;
            Content.Controls.Add(Console);
            Results.Dock = DockStyle.Fill;
            Results.Visible = false;
            Content.Controls.Add(Results);
            Controls.Add(Content);
            ApplyTheme();
        }

        string[] Titles { get { return new string[] { Strings.TerminalImmediate, Strings.TerminalLocals, Strings.TerminalWatch, Strings.TerminalConsole, Strings.TerminalResults }; } }

        int StripHeight { get { return (int)Math.Round(26 * scale); } }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            using (Graphics g = CreateGraphics()) scale = g.DpiX / 96f;
            PerformLayout();
        }

        protected override void OnLayout(LayoutEventArgs e)
        {
            Content.SetBounds(0, StripHeight, Width, Math.Max(0, Height - StripHeight));
        }

        public void SelectTab(int tab)
        {
            ActiveTab = tab;
            Console.Visible = tab == TerminalWindow.TabConsole;
            Results.Visible = tab == TerminalWindow.TabResults;
            Invalidate();
            if (TabChanged != null) TabChanged(tab);
        }

        public void ApplyTheme()
        {
            Theme t = Theme.Find(Settings.ColorTheme);
            BackColor = t.Sidebar;
            Content.BackColor = t.Background;
            Console.ApplyTheme(t);
            Results.ApplyTheme(t);
            Invalidate();
        }

        // O × fica na ponta direita da faixa das abas, alinhado a elas
        Rectangle CloseButton
        {
            get
            {
                int size = StripHeight;
                return new Rectangle(Width - size, 0, size, size - 1);
            }
        }

        const int HotClose = 100;

        protected override void OnPaint(PaintEventArgs e)
        {
            Theme t = Theme.Find(Settings.ColorTheme);
            Graphics g = e.Graphics;
            int height = StripHeight;
            using (SolidBrush b = new SolidBrush(t.Sidebar)) g.FillRectangle(b, 0, 0, Width, height);
            using (Pen line = new Pen(t.Border)) g.DrawLine(line, 0, height - 1, Width, height - 1);
            const TextFormatFlags flags = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
                TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;
            string[] titles = Titles;
            int x = 0, pad = (int)Math.Round(14 * scale);
            for (int i = 0; i < titles.Length; i++)
            {
                int width = TextRenderer.MeasureText(g, titles[i], Font, Size.Empty, flags).Width + 2 * pad;
                Rectangle r = new Rectangle(x, 0, width, height);
                tabs[i] = r;
                bool on = i == ActiveTab;
                Color back = on ? t.Background : i == hot ? t.Hover : t.Sidebar;
                using (SolidBrush b = new SolidBrush(back)) g.FillRectangle(b, r.Left, 0, r.Width, on ? height : height - 1);
                if (on) using (SolidBrush b = new SolidBrush(t.Accent)) g.FillRectangle(b, r.Left, height - Math.Max(2, (int)scale * 2), r.Width, Math.Max(2, (int)scale * 2));
                TextRenderer.DrawText(g, titles[i], Font, r, on ? t.Foreground : t.Muted, flags);
                x += width;
            }

            Rectangle close = CloseButton;
            if (hot == HotClose) using (SolidBrush b = new SolidBrush(t.Hover)) g.FillRectangle(b, close);
            int arm = Math.Max(3, (int)Math.Round(4 * scale));
            int cx = close.Left + close.Width / 2, cy = close.Top + close.Height / 2;
            using (Pen pen = new Pen(hot == HotClose ? t.Foreground : t.Muted, Math.Max(1f, scale)))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.DrawLine(pen, cx - arm, cy - arm, cx + arm, cy + arm);
                g.DrawLine(pen, cx - arm, cy + arm, cx + arm, cy - arm);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
            }
        }

        // Aba (0 a 2), o × (HotClose), a área vazia da faixa (-2) ou fora dela (-1)
        int HitTest(Point p)
        {
            if (p.Y < 0 || p.Y >= StripHeight) return -1;
            if (CloseButton.Contains(p)) return HotClose;
            for (int i = 0; i < tabs.Length; i++) if (tabs[i].Contains(p)) return i;
            return -2;
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left) return;
            int i = HitTest(e.Location);
            if (i >= 0 && i < tabs.Length) TerminalWindow.Show(i);
            else if (i == -2) TerminalWindow.StartDrag(PointToScreen(e.Location), e.Clicks > 1);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button == MouseButtons.Left && HitTest(e.Location) == HotClose) TerminalWindow.Hide();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int i = HitTest(e.Location);
            if (i == -2) i = -1;
            if (i != hot) { hot = i; Invalidate(); }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (hot != -1) { hot = -1; Invalidate(); }
        }
    }
}
