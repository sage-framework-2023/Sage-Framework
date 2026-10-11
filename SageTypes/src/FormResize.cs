using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace SageTypes
{
    // Resizable Me: o UserForm ganha borda para redimensionar (e os botões de maximizar e
    // minimizar), e os controles ancorados acompanham o tamanho:
    //
    //   Private Sub UserForm_Initialize()
    //       Resizable Me, Anchors:=Array("gridDados", "Fill", "lblInfo", "Top Width", "btnFechar", "Bottom Right")
    //   End Sub
    //
    // Âncoras (separadas por espaço; as posições são as do designer):
    //   Left, Top        ficam onde estão (o padrão)
    //   Right, Bottom    acompanham a borda direita / de baixo (o controle se move)
    //   Width, Height    esticam junto com o form
    //   Fill             = Width Height
    //
    // A janela do form é a do próprio objeto (IOleWindow); o tamanho mínimo, por padrão, é o do
    // designer. Tudo é solto quando a janela fecha.
    static class FormResize
    {
        sealed class Anchor
        {
            public string Name;
            public bool Right, Bottom, Width, Height;
            public double Left0, Top0, Width0, Height0;
        }

        sealed class State
        {
            public IntPtr Hwnd;
            public object Form;
            public List<Anchor> Anchors = new List<Anchor>();
            public double InsideWidth0, InsideHeight0;
            public int MinWidthPx, MinHeightPx;
            public bool Busy;
        }

        static readonly Dictionary<IntPtr, State> states = new Dictionary<IntPtr, State>();
        static readonly SubclassProc proc = Proc;
        const int SubclassId = 0x5A6E;

        public static void Apply(object form, object anchors, object minWidth, object minHeight, object maxButton, object minButton)
        {
            if (form == null || Interop.IsMissing(form)) throw Interop.Error(5, "ValueError: Resizable precisa do formulário (Me)");
            IntPtr hwnd = FindForm(form);
            if (hwnd == IntPtr.Zero) throw Interop.Error(5, "ValueError: a janela do formulário não foi encontrada (chame Resizable no UserForm_Initialize ou depois)");

            State state;
            if (states.TryGetValue(hwnd, out state)) RemoveWindowSubclass(hwnd, proc, (UIntPtr)SubclassId);
            state = new State { Hwnd = hwnd, Form = form };
            // Tamanho interno de referência (o do designer), em pixels da área cliente da janela
            
            RECT client;
            GetClientRect(hwnd, out client);
            state.InsideWidth0 = client.Right;
            state.InsideHeight0 = client.Bottom;
            object controls = Get(form, "Controls");

            // Âncoras: nome do controle, posição: Array("nome", "Fill", ...) ou DictionaryS
            if (anchors != null && !Interop.IsMissing(anchors) && !(anchors is DBNull))
                foreach (KeyValuePair<object, object> pair in Interop.PairsOf(Interop.Unwrap(anchors)))
                {
                    string name = Interop.Text(pair.Key);
                    object control;
                    try { control = Call(controls, "Item", name); }
                    catch (Exception) { throw Interop.Error(9, "KeyError: o controle '" + name + "' não está no formulário"); }
                    Anchor a = Parse(name, Interop.Text(pair.Value));
                    a.Left0 = Number(Get(control, "Left")); a.Top0 = Number(Get(control, "Top"));
                    a.Width0 = Number(Get(control, "Width")); a.Height0 = Number(Get(control, "Height"));
                    Marshal.ReleaseComObject(control);
                    state.Anchors.Add(a);
                }

            
            // Tamanho mínimo, em pontos (como Width/Height do form); o padrão é o do designer
            RECT r;
            GetWindowRect(hwnd, out r);
            double pxPerPt = Dpi(hwnd) / 72.0;
            double outerW = r.Right - r.Left, outerH = r.Bottom - r.Top;
            double formW = Number(Get(form, "Width")), formH = Number(Get(form, "Height"));
            state.MinWidthPx = (int)Math.Round(IsGiven(minWidth) ? outerW + (ToDouble(minWidth) - formW) * pxPerPt : outerW);
            state.MinHeightPx = (int)Math.Round(IsGiven(minHeight) ? outerH + (ToDouble(minHeight) - formH) * pxPerPt : outerH);

            long style = GetWindowLongPtr(hwnd, GWL_STYLE).ToInt64();
            style |= WS_THICKFRAME;
            if (!IsGiven(maxButton) || ToBool(maxButton)) style |= WS_MAXIMIZEBOX; else style &= ~WS_MAXIMIZEBOX;
            if (!IsGiven(minButton) || ToBool(minButton)) style |= WS_MINIMIZEBOX; else style &= ~WS_MINIMIZEBOX;
            SetWindowLongPtr(hwnd, GWL_STYLE, new IntPtr(style));
            SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);

            states[hwnd] = state;
            SetWindowSubclass(hwnd, proc, (UIntPtr)SubclassId, IntPtr.Zero);
        }

        static Anchor Parse(string name, string text)
        {
            Anchor a = new Anchor { Name = name };
            foreach (string token in text.Split(new[] { ' ', ',', '+', '|' }, StringSplitOptions.RemoveEmptyEntries))
            {
                switch (token.ToLowerInvariant())
                {
                    case "left": case "top": break;
                    case "right": a.Right = true; break;
                    case "bottom": a.Bottom = true; break;
                    case "width": a.Width = true; break;
                    case "height": a.Height = true; break;
                    case "fill": a.Width = a.Height = true; break;
                    default:
                        throw Interop.Error(5, "ValueError: âncora '" + token + "' do controle '" + name +
                            "' (use Left, Top, Right, Bottom, Width, Height ou Fill)");
                }
            }
            return a;
        }

        static IntPtr Proc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, IntPtr data)
        {
            State state;
            if (!states.TryGetValue(hwnd, out state)) return DefSubclassProc(hwnd, msg, wParam, lParam);
            switch (msg)
            {
                case WM_GETMINMAXINFO:
                    {
                        IntPtr result = DefSubclassProc(hwnd, msg, wParam, lParam);
                        // MINMAXINFO: ptReserved, ptMaxSize, ptMaxPosition, ptMinTrackSize (x, y)...
                        Marshal.WriteInt32(lParam, 24, Math.Max(Marshal.ReadInt32(lParam, 24), state.MinWidthPx));
                        Marshal.WriteInt32(lParam, 28, Math.Max(Marshal.ReadInt32(lParam, 28), state.MinHeightPx));
                        return result;
                    }
                case WM_SIZE:
                    {
                        IntPtr result = DefSubclassProc(hwnd, msg, wParam, lParam);
                        if ((int)wParam != SIZE_MINIMIZED) Arrange(state);
                        return result;
                    }
                case WM_NCDESTROY:
                    Release(state);
                    return DefSubclassProc(hwnd, msg, wParam, lParam);
            }
            return DefSubclassProc(hwnd, msg, wParam, lParam);
        }

        // Reposiciona os controles ancorados pela diferença do tamanho interno desde o designer
        static void Arrange(State state)
        {
            if (state.Busy || state.Anchors.Count == 0) return;
            state.Busy = true;
            try
            {
                // Diferença desde o designer, de pixels para pontos (as medidas dos controles)
                RECT client;
                GetClientRect(state.Hwnd, out client);
                double ptPerPx = 72.0 / Dpi(state.Hwnd);
                double dx = (client.Right - state.InsideWidth0) * ptPerPx;
                double dy = (client.Bottom - state.InsideHeight0) * ptPerPx;
                object controls = Get(state.Form, "Controls");
                foreach (Anchor a in state.Anchors)
                {
                    object c = Call(controls, "Item", a.Name);
                    if (a.Right) Set(c, "Left", Math.Max(0, a.Left0 + dx));
                    if (a.Bottom) Set(c, "Top", Math.Max(0, a.Top0 + dy));
                    if (a.Width) Set(c, "Width", Math.Max(0, a.Width0 + dx));
                    if (a.Height) Set(c, "Height", Math.Max(0, a.Height0 + dy));
                    Marshal.ReleaseComObject(c);
                }
                Marshal.ReleaseComObject(controls);
            }
            catch (Exception) { } // form fechando, controle removido
            finally { state.Busy = false; }
        }

        static void Release(State state)
        {
            RemoveWindowSubclass(state.Hwnd, proc, (UIntPtr)SubclassId);
            states.Remove(state.Hwnd);
            if (state.Form != null && Marshal.IsComObject(state.Form)) Marshal.ReleaseComObject(state.Form);
            state.Form = null;
        }

        // A janela do form: pela IOleWindow do próprio objeto; senão, na thread do Excel, a janela de
        // UserForm (ThunderDFrame) com o mesmo título
        static IntPtr FindForm(object form)
        {
            try
            {
                IOleWindow ole = form as IOleWindow;
                IntPtr h;
                if (ole != null && ole.GetWindow(out h) == 0 && h != IntPtr.Zero)
                {
                    IntPtr root = GetAncestor(h, GA_ROOT);
                    if (ClassName(root) == "ThunderDFrame") return root;
                }
            }
            catch (Exception) { }
            string caption;
            try { caption = Convert.ToString(Get(form, "Caption"), CultureInfo.InvariantCulture); }
            catch (Exception) { return IntPtr.Zero; }
            IntPtr found = IntPtr.Zero;
            EnumThreadWindows(GetCurrentThreadId(), delegate(IntPtr h, IntPtr l)
            {
                if (ClassName(h) != "ThunderDFrame") return true;
                StringBuilder s = new StringBuilder(512);
                GetWindowText(h, s, s.Capacity);
                if (s.ToString() == caption) { found = h; return false; }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        static string ClassName(IntPtr h)
        {
            StringBuilder s = new StringBuilder(64);
            GetClassName(h, s, s.Capacity);
            return s.ToString();
        }

        static double Dpi(IntPtr hwnd)
        {
            try { int d = GetDpiForWindow(hwnd); if (d > 0) return d; }
            catch (Exception) { }
            return 96;
        }

        // Pelo IDispatch, direto (o "dynamic" do C# falha com o objeto do UserForm do VBA)
        static object Get(object target, string name)
        {
            return target.GetType().InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, target, null);
        }

        static void Set(object target, string name, double value)
        {
            target.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, target, new object[] { (float)value });
        }

        static object Call(object target, string name, object arg)
        {
            return target.GetType().InvokeMember(name, System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.GetProperty,
                null, target, new object[] { arg });
        }

        static double Number(object v) { return Convert.ToDouble(v, CultureInfo.InvariantCulture); }

        static bool IsGiven(object v) { return v != null && !Interop.IsMissing(v) && !(v is DBNull); }
        static double ToDouble(object v) { return Convert.ToDouble(Interop.Unwrap(v), CultureInfo.InvariantCulture); }
        static bool ToBool(object v) { return Convert.ToBoolean(Interop.Unwrap(v), CultureInfo.InvariantCulture); }

        const int GWL_STYLE = -16, GA_ROOT = 2, WM_SIZE = 0x05, WM_GETMINMAXINFO = 0x24, WM_NCDESTROY = 0x82, SIZE_MINIMIZED = 1;
        const long WS_THICKFRAME = 0x40000, WS_MAXIMIZEBOX = 0x10000, WS_MINIMIZEBOX = 0x20000;
        const uint SWP_NOSIZE = 0x1, SWP_NOMOVE = 0x2, SWP_NOZORDER = 0x4, SWP_NOACTIVATE = 0x10, SWP_FRAMECHANGED = 0x20;

        [StructLayout(LayoutKind.Sequential)] struct RECT { public int Left, Top, Right, Bottom; }
        delegate IntPtr SubclassProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, IntPtr data);
        delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);

        [ComImport, Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IOleWindow
        {
            [PreserveSig] int GetWindow(out IntPtr hwnd);
            [PreserveSig] int ContextSensitiveHelp(int enter);
        }

        [DllImport("comctl32.dll")] static extern bool SetWindowSubclass(IntPtr hwnd, SubclassProc proc, UIntPtr id, IntPtr data);
        [DllImport("comctl32.dll")] static extern bool RemoveWindowSubclass(IntPtr hwnd, SubclassProc proc, UIntPtr id);
        [DllImport("comctl32.dll")] static extern IntPtr DefSubclassProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] static extern IntPtr GetWindowLongPtr(IntPtr hwnd, int index);
        [DllImport("user32.dll")] static extern IntPtr SetWindowLongPtr(IntPtr hwnd, int index, IntPtr value);
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT r);
        [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr hwnd, out RECT r);
        [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr hwnd, int flags);
        [DllImport("user32.dll")] static extern int GetDpiForWindow(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool EnumThreadWindows(uint thread, EnumProc proc, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder s, int n);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder s, int n);
        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    }
}
