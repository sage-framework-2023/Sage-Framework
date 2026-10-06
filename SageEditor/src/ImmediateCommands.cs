using System;
using System.Runtime.InteropServices;
using System.Text;

namespace SageEditor
{
    // Comando digitado na Verificação imediata, como num terminal:
    //   Clear + Enter   apaga tudo na Verificação imediata
    // (sem diferenciar maiúsculas; o Enter é interceptado antes de o VBE executar a linha).
    //
    // O VBE não dá acesso ao texto da janela, então a linha é acompanhada pelo que é
    // digitado desde o começo dela (cursor na primeira coluna). Se o cursor for movido
    // (setas, clique...) antes do Enter, a linha fica "desconhecida" e o Enter segue
    // normalmente para o VBE.
    static class ImmediateCommands
    {
        public static dynamic Vbe;
        static IntPtr vbeWindow, window;
        static readonly Native.SubclassProc subclassProc = SubclassProc;
        static readonly UIntPtr SubclassId = (UIntPtr)0x5A71;
        static string typed;                     // texto digitado na linha; null = desconhecido
        static int lineStartX = int.MaxValue;    // posição do cursor na primeira coluna
        static bool swallowReturnChar;
        static int tick;

        const int WM_KEYDOWN = 0x0100, WM_CHAR = 0x0102, WM_LBUTTONDOWN = 0x0201, WM_RUN = 0x8000 + 0x5A71;
        const int VK_BACK = 0x08, VK_RETURN = 0x0D, VK_DELETE = 0x2E;
        const int vbext_wt_Immediate = 5, SelectAllId = 756, ClearId = 478; // Editar > Selecionar tudo / Limpar
        const int CommandClear = 1;

        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        [DllImport("user32.dll")] static extern bool GetCaretPos(out POINT p);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);

        public static void Start(IntPtr vbe) { vbeWindow = vbe; }

        // Encontra a Verificação imediata (encaixada ou flutuante) e a subclassifica
        public static void Poll()
        {
            if (++tick % 7 != 0 || vbeWindow == IntPtr.Zero) return;
            if (window != IntPtr.Zero && Native.IsWindow(window)) return;
            window = IntPtr.Zero;
            string caption = null;
            foreach (dynamic w in Vbe.Windows)
                if ((int)w.Type == vbext_wt_Immediate) { caption = w.Caption; break; }
            if (caption == null) return;

            IntPtr found = IntPtr.Zero;
            Native.EnumWindowsProc match = delegate(IntPtr h, IntPtr l)
            {
                if (Native.ClassName(h) == "VbaWindow" && Title(h) == caption) { found = h; return false; }
                return true;
            };
            Native.EnumChildWindows(vbeWindow, match, IntPtr.Zero);
            if (found == IntPtr.Zero) // flutuante: dentro de uma janela própria, do VBE
                Native.EnumThreadWindows(Native.GetCurrentThreadId(), delegate(IntPtr top, IntPtr l)
                {
                    if (found == IntPtr.Zero && Native.GetAncestor(top, Native.GA_ROOTOWNER) == vbeWindow)
                        Native.EnumChildWindows(top, match, IntPtr.Zero);
                    return found == IntPtr.Zero;
                }, IntPtr.Zero);
            if (found != IntPtr.Zero && Native.SetWindowSubclass(found, subclassProc, SubclassId, UIntPtr.Zero))
            {
                window = found;
                typed = null;
            }
        }

        public static void Shutdown()
        {
            if (window != IntPtr.Zero) Native.RemoveWindowSubclass(window, subclassProc, SubclassId);
            window = IntPtr.Zero;
        }

        static string Title(IntPtr hwnd)
        {
            StringBuilder sb = new StringBuilder(256);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        // ------------------------------------------------------------------

        static IntPtr SubclassProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            try
            {
                switch (msg)
                {
                    case Native.WM_NCDESTROY:
                        Native.RemoveWindowSubclass(hwnd, subclassProc, SubclassId);
                        if (hwnd == window) window = IntPtr.Zero;
                        break;

                    case WM_CHAR:
                    {
                        char c = (char)(int)wParam;
                        if (c == '\r' && swallowReturnChar) { swallowReturnChar = false; return IntPtr.Zero; }
                        if (c == '\b')
                        {
                            typed = string.IsNullOrEmpty(typed) ? null : typed.Substring(0, typed.Length - 1);
                            break;
                        }
                        if (c < ' ') break;
                        // Cursor na primeira coluna: começa uma linha conhecida
                        POINT p;
                        GetCaretPos(out p);
                        if (p.X <= lineStartX) { lineStartX = p.X; typed = ""; }
                        if (typed != null) typed += c;
                        break;
                    }

                    case WM_KEYDOWN:
                    {
                        int key = (int)wParam;
                        if (key == VK_RETURN)
                        {
                            int command = Command(typed), length = typed == null ? 0 : typed.Length;
                            typed = null;
                            if (command != 0)
                            {
                                swallowReturnChar = true;
                                PostMessage(hwnd, WM_RUN, (IntPtr)command, (IntPtr)length); // depois desta tecla
                                return IntPtr.Zero;
                            }
                        }
                        // Setas, Home, End, PgUp/PgDn, Delete: o cursor sai do fim da linha conhecida
                        else if ((key >= 0x21 && key <= 0x28) || key == VK_DELETE) typed = null;
                        break;
                    }

                    case WM_LBUTTONDOWN:
                        typed = null;
                        break;

                    case WM_RUN:
                        Run(hwnd, (int)wParam, (int)lParam);
                        return IntPtr.Zero;
                }
            }
            catch (Exception ex) { Log.Error(ex); }
            return Native.DefSubclassProc(hwnd, msg, wParam, lParam);
        }

        static int Command(string line)
        {
            if (line == null) return 0;
            switch (line.Trim().ToLowerInvariant())
            {
                case "clear": return CommandClear;
            }
            return 0;
        }

        static void Run(IntPtr hwnd, int command, int typedLength)
        {
            // Tudo, inclusive o comando digitado: Editar > Selecionar tudo e Limpar, com a
            // Verificação imediata ativa
            Execute(SelectAllId);
            Execute(ClearId);
        }

        static void Execute(int id)
        {
            dynamic control = Vbe.CommandBars.FindControl(Type.Missing, id);
            if (control != null && (bool)control.Enabled) control.Execute();
        }
    }
}
