using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace SageEditor
{
    // Aplica o tema a todas as janelas do VBE (não só ao editor de código).
    //
    // 1. GetSysColor/GetSysColorBrush são interceptados na tabela de importação do
    //    VBE7.DLL, das DLLs mso* (menus e barras) e do comctl32 (árvore do Projeto).
    //    O desvio só vale enquanto uma janela do VBE está tratando uma mensagem
    //    (contador "depth", mantido pelo subclassing), então o Excel não muda.
    //    SetTextColor/SetBkColor do VBE7.DLL também: as 16 cores padrão do
    //    "Formato do editor" viram as cores do tema (palavra-chave, comentário...).
    //    No VBEUI.DLL (barras de menu e de ferramentas do VBE, com a paleta do Office),
    //    SetTextColor/SetBkColor/CreateSolidBrush/CreatePen/SetDC*Color têm a
    //    luminosidade invertida nos temas escuros (ícones são bitmaps e ficam iguais).
    // 2. Todas as janelas do VBE são subclassificadas: as existentes e as novas
    //    (hook WH_CBT). Ali também são tratados fundos e WM_CTLCOLOR*.
    // 3. Barras de rolagem e árvore usam o tema escuro do Windows (DarkMode_Explorer);
    //    as barras de título usam o DWM.
    //
    // Tudo roda na thread de interface do Excel.
    static class ThemeEngine
    {
        static Theme current;
        static IntPtr vbeWindow;
        static bool initialized;

        [ThreadStatic] static int depth;

        // Subclassing
        static readonly Native.SubclassProc subclassProc = SubclassProc;
        static readonly Dictionary<IntPtr, string> subclassed = new Dictionary<IntPtr, string>();
        const int WM_NCCALCSIZE = 0x0083, WM_VSCROLL = 0x0115, WM_MOUSEWHEEL = 0x020A, WM_KEYDOWN = 0x0100, WM_LBUTTONDOWN = 0x0201;
        const int WM_NCACTIVATE = 0x0086, WM_TIMER = 0x0113, WM_MOUSEMOVE = 0x0200, WM_LBUTTONUP = 0x0202, WM_MOUSELEAVE = 0x02A3;
        static readonly UIntPtr SubclassId = (UIntPtr)0x5A6E;

        // Hook CBT para janelas criadas depois
        static readonly Native.HookProc cbtProc = CbtProc;
        static IntPtr cbtHook;

        // Desvios na tabela de importação
        static readonly Native.GetSysColorProc hookColor = HookGetSysColor;
        static readonly Native.GetSysColorBrushProc hookBrush = HookGetSysColorBrush;
        static readonly Native.SetColorProc hookText = HookSetTextColor;
        static readonly Native.SetColorProc hookBk = HookSetBkColor;
        static readonly Native.SetColorProc hookOfficeText = HookOfficeSetTextColor;
        static readonly Native.SetColorProc hookOfficeBk = HookOfficeSetBkColor;
        static readonly Native.CreateSolidBrushProc hookOfficeBrush = HookOfficeCreateSolidBrush;
        static readonly Native.CreatePenProc hookOfficePen = HookOfficeCreatePen;
        static readonly Native.SetColorProc hookOfficeDCBrush = HookOfficeSetDCBrushColor;
        static readonly Native.SetColorProc hookOfficeDCPen = HookOfficeSetDCPenColor;
        static readonly Native.GetStockObjectProc hookOfficeStock = HookOfficeGetStockObject;
        static readonly Native.FillRectProc hookOfficeFillRect = HookOfficeFillRect;
        static Native.GetSysColorProc origColor;
        static Native.GetSysColorBrushProc origBrush;
        static Native.SetColorProc origText, origBk;
        static Native.CreateSolidBrushProc origSolidBrush;
        static Native.CreatePenProc origPen;
        static Native.SetColorProc origDCBrush, origDCPen;
        static Native.GetStockObjectProc origStock;
        static Native.FillRectProc origFillRect;
        static Dictionary<string, IntPtr> user32Hooks, gdi32Hooks, officeGdiHooks;
        static readonly List<KeyValuePair<IntPtr, IntPtr>> iatPatches = new List<KeyValuePair<IntPtr, IntPtr>>();
        static readonly HashSet<IntPtr> patchedModules = new HashSet<IntPtr>();

        public static Theme Current { get { return current; } }

        // Tema escuro ativo e pintura de uma janela do VBE em andamento
        public static Theme DarkPainting
        {
            get
            {
                Theme t = current;
                return depth > 0 && t != null && t.IsDark ? t : null;
            }
        }

        public static void Initialize(IntPtr vbe)
        {
            if (initialized) return;
            vbeWindow = vbe;

            IntPtr user32 = Native.GetModuleHandle("user32.dll");
            IntPtr gdi32 = Native.GetModuleHandle("gdi32.dll");
            origColor = (Native.GetSysColorProc)Original(user32, "GetSysColor", typeof(Native.GetSysColorProc));
            origBrush = (Native.GetSysColorBrushProc)Original(user32, "GetSysColorBrush", typeof(Native.GetSysColorBrushProc));
            origText = (Native.SetColorProc)Original(gdi32, "SetTextColor", typeof(Native.SetColorProc));
            origBk = (Native.SetColorProc)Original(gdi32, "SetBkColor", typeof(Native.SetColorProc));

            user32Hooks = new Dictionary<string, IntPtr>();
            user32Hooks["GetSysColor"] = Marshal.GetFunctionPointerForDelegate(hookColor);
            user32Hooks["GetSysColorBrush"] = Marshal.GetFunctionPointerForDelegate(hookBrush);
            Painters.HookEdges(user32, user32Hooks);
            gdi32Hooks = new Dictionary<string, IntPtr>();
            gdi32Hooks["SetTextColor"] = Marshal.GetFunctionPointerForDelegate(hookText);
            gdi32Hooks["SetBkColor"] = Marshal.GetFunctionPointerForDelegate(hookBk);
            Syntax.Hooks(gdi32, gdi32Hooks);

            origSolidBrush = (Native.CreateSolidBrushProc)Original(gdi32, "CreateSolidBrush", typeof(Native.CreateSolidBrushProc));
            origPen = (Native.CreatePenProc)Original(gdi32, "CreatePen", typeof(Native.CreatePenProc));
            officeGdiHooks = new Dictionary<string, IntPtr>();
            officeGdiHooks["SetTextColor"] = Marshal.GetFunctionPointerForDelegate(hookOfficeText);
            officeGdiHooks["SetBkColor"] = Marshal.GetFunctionPointerForDelegate(hookOfficeBk);
            officeGdiHooks["CreateSolidBrush"] = Marshal.GetFunctionPointerForDelegate(hookOfficeBrush);
            officeGdiHooks["CreatePen"] = Marshal.GetFunctionPointerForDelegate(hookOfficePen);
            origDCBrush = (Native.SetColorProc)Original(gdi32, "SetDCBrushColor", typeof(Native.SetColorProc));
            origDCPen = (Native.SetColorProc)Original(gdi32, "SetDCPenColor", typeof(Native.SetColorProc));
            officeGdiHooks["SetDCBrushColor"] = Marshal.GetFunctionPointerForDelegate(hookOfficeDCBrush);
            officeGdiHooks["SetDCPenColor"] = Marshal.GetFunctionPointerForDelegate(hookOfficeDCPen);
            origStock = (Native.GetStockObjectProc)Original(gdi32, "GetStockObject", typeof(Native.GetStockObjectProc));
            officeGdiHooks["GetStockObject"] = Marshal.GetFunctionPointerForDelegate(hookOfficeStock);

            origFillRect = (Native.FillRectProc)Original(user32, "FillRect", typeof(Native.FillRectProc));
            user32Hooks["FillRect"] = Marshal.GetFunctionPointerForDelegate(hookOfficeFillRect);

            initialized = true;
        }

        public static void Apply(Theme theme)
        {
            if (!initialized) return;
            Theme previous = current;
            current = theme.IsDefault ? null : theme;
            if (previous == null && current == null) return; // tema padrão: não toca em nada

            if (current != null)
            {
                PatchModules();
                if (cbtHook == IntPtr.Zero)
                    cbtHook = Native.SetWindowsHookEx(Native.WH_CBT, cbtProc, IntPtr.Zero, Native.GetCurrentThreadId());
            }
            SubclassExisting();

            foreach (IntPtr hwnd in new List<IntPtr>(subclassed.Keys))
                StyleWindow(hwnd);

            LineNumbers.RefreshFrames(new List<IntPtr>(subclassed.Keys));
            if (previous != null || current != null)
                Refresh();
        }

        public static void Shutdown()
        {
            if (!initialized) return;
            current = null;
            LineNumbers.RefreshFrames(new List<IntPtr>(subclassed.Keys));
            foreach (IntPtr hwnd in new List<IntPtr>(subclassed.Keys))
            {
                StyleWindow(hwnd);
                Native.RemoveWindowSubclass(hwnd, subclassProc, SubclassId);
            }
            subclassed.Clear();
            if (cbtHook != IntPtr.Zero) { Native.UnhookWindowsHookEx(cbtHook); cbtHook = IntPtr.Zero; }
            RestoreModules();
            Refresh();
            initialized = false;
        }

        // ------------------------------------------------------------------
        // Janelas
        // ------------------------------------------------------------------

        static bool BelongsToVbe(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return false;
            if (hwnd == vbeWindow) return true;
            // Diálogos (Opções, Referências, Localizar...) ficam como estão
            IntPtr root = Native.GetAncestor(hwnd, Native.GA_ROOT);
            if (root != IntPtr.Zero && Native.ClassName(root) == "#32770") return false;
            return Native.GetAncestor(hwnd, Native.GA_ROOTOWNER) == vbeWindow;
        }

        static void SubclassExisting()
        {
            List<IntPtr> found = new List<IntPtr>();
            found.Add(vbeWindow);
            Native.EnumChildWindows(vbeWindow, delegate(IntPtr h, IntPtr l) { found.Add(h); return true; }, IntPtr.Zero);
            Native.EnumThreadWindows(Native.GetCurrentThreadId(), delegate(IntPtr h, IntPtr l)
            {
                if (h != vbeWindow && BelongsToVbe(h))
                {
                    found.Add(h);
                    Native.EnumChildWindows(h, delegate(IntPtr c, IntPtr l2) { found.Add(c); return true; }, IntPtr.Zero);
                }
                return true;
            }, IntPtr.Zero);

            foreach (IntPtr h in found) Subclass(h);
        }

        static void Subclass(IntPtr hwnd)
        {
            if (subclassed.ContainsKey(hwnd)) return;
            if (Native.SetWindowSubclass(hwnd, subclassProc, SubclassId, UIntPtr.Zero))
                subclassed[hwnd] = Native.ClassName(hwnd);
        }

        static IntPtr CbtProc(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code == Native.HCBT_CREATEWND && current != null)
            {
                try
                {
                    // CBT_CREATEWND { CREATESTRUCT* lpcs; ... }; CREATESTRUCT.hwndParent é o 4º campo
                    IntPtr cs = Marshal.ReadIntPtr(lParam);
                    IntPtr parent = Marshal.ReadIntPtr(cs, 3 * IntPtr.Size);
                    if (BelongsToVbe(parent))
                    {
                        Subclass(wParam);
                        pendingStyle.Add(wParam);
                    }
                }
                catch (Exception ex) { Log.Error(ex); }
            }
            return Native.CallNextHookEx(cbtHook, code, wParam, lParam);
        }

        // Janelas novas recebem o estilo (tema de rolagem, DWM) no primeiro WM_ERASEBKGND
        static readonly HashSet<IntPtr> pendingStyle = new HashSet<IntPtr>();

        static IntPtr SubclassProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            depth++;
            try
            {
                Theme t = current;
                if (t != null)
                {
                    string cls;
                    subclassed.TryGetValue(hwnd, out cls);

                    // Caixa de ferramentas: recolorida depois de cada desenho do FM20
                    // (inclusive os feitos fora do WM_PAINT, ao passar o mouse e clicar)
                    if (t.IsDark && cls != null && cls.StartsWith("F3 ") && IsToolbox(hwnd))
                    {
                        switch (msg)
                        {
                            case Native.WM_PAINT:
                            case Native.WM_NCPAINT:
                            case WM_NCACTIVATE:
                            case WM_MOUSEMOVE:
                            case WM_MOUSELEAVE:
                            case WM_LBUTTONDOWN:
                            case WM_LBUTTONUP:
                            case WM_TIMER:
                            {
                                IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                                Painters.InvertGrays(hwnd, t); // com as bordas (área não-cliente)
                                return result;
                            }
                        }
                    }

                    switch (msg)
                    {
                        case Native.WM_ERASEBKGND:
                            if (pendingStyle.Remove(hwnd)) StyleWindow(hwnd);
                            int background;
                            if (OwnBackground(hwnd, cls, t, out background))
                            {
                                Native.RECT r;
                                Native.GetClientRect(hwnd, out r);
                                Native.FillRect(wParam, ref r, t.Brush(background));
                                return (IntPtr)1;
                            }
                            break;

                        case Native.WM_CTLCOLOREDIT:
                        case Native.WM_CTLCOLORLISTBOX:
                        case Native.WM_CTLCOLORSTATIC:
                        {
                            Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                            int bg = msg == Native.WM_CTLCOLORSTATIC ? t.Face : t.Window;
                            Native.SetTextColor(wParam, t.WindowText);
                            Native.SetBkColor(wParam, bg);
                            return t.Brush(bg);
                        }

                        case Native.WM_NCPAINT:
                        {
                            IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                            Painters.PaintBorder(hwnd, t);
                            if (cls == "VbaWindow") LineNumbers.Paint(hwnd, t);
                            return result;
                        }

                        case WM_NCCALCSIZE:
                            if (cls == "VbaWindow" && LineNumbers.IsCodePane(hwnd))
                            {
                                IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                                LineNumbers.AdjustClient(hwnd, lParam);
                                return result;
                            }
                            break;

                        // Rolagem, edição e cliques: atualiza os números sem esperar a verificação periódica
                        case WM_VSCROLL:
                        case WM_MOUSEWHEEL:
                        case WM_KEYDOWN:
                        case WM_LBUTTONDOWN:
                            if (cls == "VbaWindow")
                            {
                                IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                                LineNumbers.Paint(hwnd, t);
                                return result;
                            }
                            break;

                        case Native.WM_PAINT:
                            if (cls == "SysTabControl32")
                            {
                                Painters.PaintTabs(hwnd, t);
                                return IntPtr.Zero;
                            }
                            if (cls == "VbaWindow")
                            {
                                IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                                LineNumbers.Paint(hwnd, t);
                                return result;
                            }
                            if (cls == "ObtbarWndClass")
                            {
                                IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                                Painters.Recolor(hwnd, t);
                                return result;
                            }
                            if (cls == "MsoCommandBar" && t.IsDark)
                            {
                                // O VBEUI só pinta a área dos botões; o resto da barra (à
                                // direita) ficaria com o que havia antes.
                                // Só a região que vai ser repintada (o VBEUI não redesenha o resto).
                                IntPtr region = Native.CreateRectRgn(0, 0, 0, 0);
                                if (Native.GetUpdateRgn(hwnd, region, false) > 1)
                                {
                                    IntPtr dc = Native.GetDC(hwnd);
                                    Native.FillRgn(dc, region, t.Brush(t.Face));
                                    Native.ReleaseDC(hwnd, dc);
                                }
                                Native.DeleteObject(region);

                                // campo "Ln, Col" da barra Padrão sai com o cinza fixo do Windows
                                IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                                Painters.ReplaceColors(hwnd, t, 0xF0F0F0, 0xF1F1F1);
                                return result;
                            }
                            break;
                    }
                }
                if (msg == Native.WM_NCDESTROY)
                {
                    Native.RemoveWindowSubclass(hwnd, subclassProc, SubclassId);
                    subclassed.Remove(hwnd);
                    pendingStyle.Remove(hwnd);
                }
                return Native.DefSubclassProc(hwnd, msg, wParam, lParam);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                return Native.DefSubclassProc(hwnd, msg, wParam, lParam);
            }
            finally
            {
                depth--;
            }
        }

        // Fundos pintados pelo user32 com o pincel da classe (sem passar pelo
        // GetSysColorBrush desviado) ou com cor fixa pelo VBE.
        static bool OwnBackground(IntPtr hwnd, string cls, Theme t, out int color)
        {
            switch (cls)
            {
                case "MDIClient":
                case "DesignerWindow": // fundo em volta do UserForm no designer
                    color = t.Window;
                    return true;
                case "VBSlider":
                case "MsoCommandBarDock":
                case "MsoCommandBar":
                case "PROJECT":
                case "wndclass_pbrs":
                case "wndclass_desked_gsk":
                case "DockingView":
                    color = t.Face;
                    return true;
                case "Edit":
                    // caixa de edição da janela Propriedades (o VBE a apaga de branco)
                    color = t.Window;
                    return Native.ClassName(Native.GetParent(hwnd)) == "wndclass_pbrs";
            }
            color = 0;
            return false;
        }

        static void StyleWindow(IntPtr hwnd)
        {
            Theme t = current;
            string cls = Native.ClassName(hwnd);
            bool dark = t != null && t.IsDark;

            if (cls == "ScrollBar" || cls == "SysTreeView32" || cls == "ListBox" || cls == "VbaWindow")
                Native.SetWindowTheme(hwnd, dark ? "DarkMode_Explorer" : null, null);
            if (cls == "Edit" || cls == "ComboBox")
                Native.SetWindowTheme(hwnd, dark ? "DarkMode_CFD" : null, null);

            if (cls == "SysTreeView32")
            {
                // -1 = cor do sistema (que já vem do tema pelo GetSysColor)
                Native.SendMessage(hwnd, Native.TVM_SETBKCOLOR, IntPtr.Zero, (IntPtr)(t == null ? -1 : t.Face));
                Native.SendMessage(hwnd, Native.TVM_SETTEXTCOLOR, IntPtr.Zero, (IntPtr)(t == null ? -1 : t.WindowText));
                Native.SendMessage(hwnd, Native.TVM_SETLINECOLOR, IntPtr.Zero, (IntPtr)(t == null ? -1 : t.FrameBorder));
            }

            // Barra de título escura do DWM. A Caixa de ferramentas tem dono (GetParent devolve o dono).
            if (Native.GetParent(hwnd) == IntPtr.Zero || hwnd == vbeWindow || cls.StartsWith("F3 MinFrame"))
            {
                int on = dark ? 1 : 0;
                Native.DwmSetWindowAttribute(hwnd, Native.DWMWA_USE_IMMERSIVE_DARK_MODE, ref on, 4);
                int caption = t == null ? Native.DWMWA_COLOR_DEFAULT : t.Caption;
                int text = t == null ? Native.DWMWA_COLOR_DEFAULT : t.CaptionText;
                int border = t == null ? Native.DWMWA_COLOR_DEFAULT : t.FrameBorder;
                Native.DwmSetWindowAttribute(hwnd, Native.DWMWA_CAPTION_COLOR, ref caption, 4);
                Native.DwmSetWindowAttribute(hwnd, Native.DWMWA_TEXT_COLOR, ref text, 4);
                Native.DwmSetWindowAttribute(hwnd, Native.DWMWA_BORDER_COLOR, ref border, 4);
            }
        }

        static void Refresh()
        {
            foreach (IntPtr hwnd in new List<IntPtr>(subclassed.Keys))
            {
                if (!Native.IsWindow(hwnd)) { subclassed.Remove(hwnd); continue; }
                // Faz controles que guardam cores em cache (árvore, barras do Office) relerem
                Native.SendMessage(hwnd, Native.WM_SYSCOLORCHANGE, IntPtr.Zero, IntPtr.Zero);
                Native.SendMessage(hwnd, Native.WM_THEMECHANGED, IntPtr.Zero, IntPtr.Zero);
            }
            foreach (IntPtr hwnd in new List<IntPtr>(subclassed.Keys))
                if (Native.GetParent(hwnd) == IntPtr.Zero || hwnd == vbeWindow)
                    Native.RedrawWindow(hwnd, IntPtr.Zero, IntPtr.Zero,
                        Native.RDW_INVALIDATE | Native.RDW_ERASE | Native.RDW_FRAME | Native.RDW_ALLCHILDREN | Native.RDW_UPDATENOW);
        }

        // ------------------------------------------------------------------
        // GetSysColor / GetSysColorBrush
        // ------------------------------------------------------------------

        static uint HookGetSysColor(int index)
        {
            if (depth > 0)
            {
                Theme t = current;
                int color;
                if (t != null && t.SysColors.TryGetValue(index, out color)) return (uint)color;
            }
            return origColor(index);
        }

        static IntPtr HookGetSysColorBrush(int index)
        {
            if (depth > 0)
            {
                Theme t = current;
                int color;
                if (t != null && t.SysColors.TryGetValue(index, out color)) return t.Brush(color);
            }
            return origBrush(index);
        }

        static int HookSetTextColor(IntPtr hdc, int color)
        {
            Syntax.NoteTextColor(color);
            Theme t = current;
            return origText(hdc, t == null ? color : MapColor(hdc, color, t.Palette));
        }

        static int HookSetBkColor(IntPtr hdc, int color)
        {
            Theme t = current;
            return origBk(hdc, t == null ? color : MapColor(hdc, color, t.BackPalette));
        }

        // Cores padrão da paleta do editor -> cores do tema. Além da pintura normal
        // (depth > 0), vale quando o VBE desenha direto numa janela dele: ao editar
        // uma linha (digitando ou via CodeModule.ReplaceLine) ele repinta a linha
        // na hora, fora do WM_PAINT.
        static int MapColor(IntPtr hdc, int color, int[] palette)
        {
            if (palette != null && (color & unchecked((int)0xFF000000)) == 0 && (depth > 0 || IsVbeSurface(hdc)))
            {
                int i = Array.IndexOf(StandardPalette, color);
                if (i >= 0) return palette[i];
            }
            return color;
        }

        // Tema para o realce de sintaxe: pintura normal ou desenho direto numa janela do VBE
        public static Theme SyntaxTheme(IntPtr hdc)
        {
            Theme t = current;
            return t != null && (depth > 0 || IsVbeSurface(hdc)) ? t : null;
        }

        static readonly Dictionary<IntPtr, bool> vbeSurfaces = new Dictionary<IntPtr, bool>();

        static bool IsVbeSurface(IntPtr hdc)
        {
            IntPtr hwnd = Native.WindowFromDC(hdc);
            if (hwnd == IntPtr.Zero) return false;
            bool vbe;
            lock (vbeSurfaces)
            {
                if (!vbeSurfaces.TryGetValue(hwnd, out vbe))
                {
                    vbe = BelongsToVbe(hwnd);
                    if (vbeSurfaces.Count > 500) vbeSurfaces.Clear();
                    vbeSurfaces[hwnd] = vbe;
                }
            }
            return vbe;
        }

        // --- Microsoft Forms (FM20.DLL): Caixa de ferramentas ---
        //
        // O FM20 não lê as cores pelo GetSysColor importado (o desvio na tabela de
        // importação não o alcança), então a Caixa de ferramentas é recolorida depois de
        // desenhada (Painters.InvertGrays). Os UserForms, no designer ou rodando, ficam
        // com as cores que o usuário escolheu.

        static readonly Dictionary<IntPtr, bool> toolboxWindows = new Dictionary<IntPtr, bool>();

        static bool IsToolbox(IntPtr hwnd)
        {
            bool toolbox;
            if (!toolboxWindows.TryGetValue(hwnd, out toolbox))
            {
                toolbox = Native.ClassName(Native.GetAncestor(hwnd, Native.GA_ROOT)).StartsWith("F3 MinFrame");
                if (toolboxWindows.Count > 200) toolboxWindows.Clear();
                toolboxWindows[hwnd] = toolbox;
            }
            return toolbox;
        }

        // Chamado pela verificação periódica. A Caixa de ferramentas é criada sem dono
        // (só depois passa a pertencer ao VBE), então o hook CBT não a pega.
        public static void PollForms()
        {
            if (!initialized || current == null || Native.GetModuleHandle("fm20.dll") == IntPtr.Zero) return;

            List<IntPtr> boxes = new List<IntPtr>();
            Native.EnumThreadWindows(Native.GetCurrentThreadId(), delegate(IntPtr h, IntPtr l)
            {
                if (!subclassed.ContainsKey(h) && Native.ClassName(h).StartsWith("F3 MinFrame") && BelongsToVbe(h))
                    boxes.Add(h);
                return true;
            }, IntPtr.Zero);

            foreach (IntPtr box in boxes)
            {
                Subclass(box);
                Native.EnumChildWindows(box, delegate(IntPtr c, IntPtr l) { Subclass(c); return true; }, IntPtr.Zero);
                StyleWindow(box);
                Native.RedrawWindow(box, IntPtr.Zero, IntPtr.Zero,
                    Native.RDW_INVALIDATE | Native.RDW_ERASE | Native.RDW_FRAME | Native.RDW_ALLCHILDREN);
            }
        }

        // --- Office (barras de menu e de ferramentas) ---

        static int HookOfficeSetTextColor(IntPtr hdc, int color) { return origText(hdc, OfficeTextColor(color)); }
        static int HookOfficeSetBkColor(IntPtr hdc, int color) { return origBk(hdc, OfficeColor(color)); }
        static IntPtr HookOfficeCreateSolidBrush(int color) { return origSolidBrush(OfficeColor(color)); }
        static IntPtr HookOfficeCreatePen(int style, int width, int color) { return origPen(style, width, OfficeColor(color)); }

        static int HookOfficeSetDCBrushColor(IntPtr hdc, int color) { return origDCBrush(hdc, OfficeColor(color)); }
        static int HookOfficeSetDCPenColor(IntPtr hdc, int color) { return origDCPen(hdc, OfficeColor(color)); }

        // FillRect com pincel "COLOR_* + 1" é resolvido dentro do user32
        static int HookOfficeFillRect(IntPtr hdc, IntPtr rect, IntPtr brush)
        {
            long value = brush.ToInt64();
            Theme t = current;
            int color;
            if (depth > 0 && t != null && value > 0 && value <= 31 && t.SysColors.TryGetValue((int)value - 1, out color))
                brush = t.Brush(color);
            return origFillRect(hdc, rect, brush);
        }

        // Pincéis de estoque branco e cinza-claro (fundos de campos e barras)
        static IntPtr HookOfficeGetStockObject(int index)
        {
            const int WHITE_BRUSH = 0, LTGRAY_BRUSH = 1;
            Theme t = current;
            if (depth > 0 && t != null && t.IsDark)
            {
                if (index == WHITE_BRUSH) return t.Brush(t.Window);
                if (index == LTGRAY_BRUSH) return t.Brush(t.Face);
            }
            return origStock(index);
        }

        // Tema escuro: inverte a luminosidade. Cinzas vão do fundo (branco) ao texto
        // (preto) do tema; cores mantêm o matiz.
        static readonly int[] officeCalls = new int[2];
        public static string Diagnostics()
        {
            return "cores do Office fora=" + officeCalls[0] + " dentro=" + officeCalls[1] +
                " janelas=" + subclassed.Count + " modulos=" + patchedModules.Count + " desvios=" + iatPatches.Count;
        }

        // Texto: preto vira o texto do tema; cinzas claros (desabilitado, janela
        // inativa) viram um cinza ainda legível.
        static int OfficeTextColor(int color)
        {
            officeCalls[depth > 0 ? 1 : 0]++;
            Theme t = current;
            if (depth == 0 || t == null || !t.IsDark || (color & unchecked((int)0xFF000000)) != 0 || t.OwnsColor(color))
                return color;
            int r = color & 0xFF, g = (color >> 8) & 0xFF, b = (color >> 16) & 0xFF;
            if (Math.Max(r, Math.Max(g, b)) - Math.Min(r, Math.Min(g, b)) >= 30) return OfficeColor(color);
            double light = (r + g + b) / 765.0;
            // O Office escreve itens habilitados em preto ou cinza-escuro, desabilitados
            // em cinza-claro e o item pressionado/selecionado em quase branco.
            if (light <= 0.55 || light > 0.85) return t.WindowText;
            return Lerp(t.WindowText, t.Face, 0.55); // desabilitado
        }

        static int OfficeColor(int color)
        {
            officeCalls[depth > 0 ? 1 : 0]++;
            if (depth == 0) return color;
            Theme t = current;
            if (t == null || !t.IsDark || (color & unchecked((int)0xFF000000)) != 0) return color;
            if (t.OwnsColor(color)) return color; // já veio do tema (GetSysColor desviado)

            double r = (color & 0xFF) / 255.0, g = ((color >> 8) & 0xFF) / 255.0, b = ((color >> 16) & 0xFF) / 255.0;
            double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b));
            double light = (max + min) / 2;

            if (max - min < 0.12)
            {
                // Cinza já escuro: mistura de cores do tema feita pelo próprio VBEUI
                if (light > 0.04 && light < 0.30) return color;
                double k = 1 - light; // 0 = fundo, 1 = texto
                return Lerp(t.Face, t.WindowText, k);
            }

            double sat = light > 0.5 ? (max - min) / (2 - max - min) : (max - min) / (max + min);
            double hue;
            if (max == r) hue = (g - b) / (max - min) + (g < b ? 6 : 0);
            else if (max == g) hue = (b - r) / (max - min) + 2;
            else hue = (r - g) / (max - min) + 4;
            hue /= 6;

            double l2 = Math.Max(0.12, Math.Min(0.88, 1 - light));
            return FromHsl(hue, Math.Min(sat, 0.7), l2);
        }

        static int Lerp(int a, int b, double k)
        {
            int r = (int)((a & 0xFF) + ((b & 0xFF) - (a & 0xFF)) * k);
            int g = (int)(((a >> 8) & 0xFF) + (((b >> 8) & 0xFF) - ((a >> 8) & 0xFF)) * k);
            int bl = (int)(((a >> 16) & 0xFF) + (((b >> 16) & 0xFF) - ((a >> 16) & 0xFF)) * k);
            return r | (g << 8) | (bl << 16);
        }

        static int FromHsl(double h, double s, double l)
        {
            double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
            int r = (int)Math.Round(255 * Hue(p, q, h + 1.0 / 3));
            int g = (int)Math.Round(255 * Hue(p, q, h));
            int b = (int)Math.Round(255 * Hue(p, q, h - 1.0 / 3));
            return r | (g << 8) | (b << 16);
        }

        static double Hue(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1.0 / 6) return p + (q - p) * 6 * t;
            if (t < 1.0 / 2) return q;
            if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
            return p;
        }

        static Delegate Original(IntPtr module, string name, Type type)
        {
            return Marshal.GetDelegateForFunctionPointer(Native.GetProcAddress(module, name), type);
        }

        static void PatchModules()
        {
            foreach (ProcessModule module in Process.GetCurrentProcess().Modules)
            {
                if (patchedModules.Contains(module.BaseAddress)) continue;
                string name = module.ModuleName.ToLowerInvariant();
                bool vbe = name == "vbe7.dll", bars = name == "vbeui.dll";
                if (!vbe && !bars && name != "comctl32.dll") continue;
                try
                {
                    PatchImports(module.BaseAddress, "user32.dll", user32Hooks);
                    if (vbe) PatchImports(module.BaseAddress, "gdi32.dll", gdi32Hooks);
                    if (bars) PatchImports(module.BaseAddress, "gdi32.dll", officeGdiHooks);
                    patchedModules.Add(module.BaseAddress);
                }
                catch (Exception ex) { Log.Error(new Exception("Import de " + module.ModuleName, ex)); }
            }
        }

        static void RestoreModules()
        {
            foreach (KeyValuePair<IntPtr, IntPtr> patch in iatPatches)
                WritePointer(patch.Key, patch.Value);
            iatPatches.Clear();
            patchedModules.Clear();
        }

        static void PatchImports(IntPtr module, string dll, Dictionary<string, IntPtr> hooks)
        {
            IntPtr nt = module + Marshal.ReadInt32(module, 0x3C);
            bool pe64 = Marshal.ReadInt16(nt, 24) == 0x20B;
            int dirs = 24 + (pe64 ? 112 : 96);

            int imports = Marshal.ReadInt32(nt, dirs + 8 * 1);
            if (imports != 0)
            {
                // IMAGE_IMPORT_DESCRIPTOR (20 bytes): OriginalFirstThunk, TimeDateStamp, ForwarderChain, Name, FirstThunk
                for (IntPtr d = module + imports; ; d += 20)
                {
                    int name = Marshal.ReadInt32(d, 12);
                    if (name == 0) break;
                    if (!IsDll(module, name, dll)) continue;
                    int names = Marshal.ReadInt32(d, 0), iat = Marshal.ReadInt32(d, 16);
                    PatchThunks(module, names != 0 ? names : iat, iat, hooks);
                }
            }

            int delayed = Marshal.ReadInt32(nt, dirs + 8 * 13);
            if (delayed != 0)
            {
                // IMAGE_DELAYLOAD_DESCRIPTOR (32 bytes): Attributes, DllNameRVA, ModuleHandleRVA, IAT, INT, ...
                for (IntPtr d = module + delayed; ; d += 32)
                {
                    int name = Marshal.ReadInt32(d, 4);
                    if (name == 0) break;
                    if (!IsDll(module, name, dll)) continue;
                    PatchThunks(module, Marshal.ReadInt32(d, 16), Marshal.ReadInt32(d, 12), hooks);
                }
            }
        }

        static bool IsDll(IntPtr module, int nameRva, string dll)
        {
            return string.Equals(Marshal.PtrToStringAnsi(module + nameRva), dll, StringComparison.OrdinalIgnoreCase);
        }

        static void PatchThunks(IntPtr module, int namesRva, int iatRva, Dictionary<string, IntPtr> hooks)
        {
            int size = IntPtr.Size;
            for (int i = 0; ; i++)
            {
                long entry = size == 8 ? Marshal.ReadInt64(module, namesRva + i * size) : (uint)Marshal.ReadInt32(module, namesRva + i * size);
                if (entry == 0) break;
                bool byOrdinal = size == 8 ? entry < 0 : (entry & 0x80000000L) != 0;
                if (byOrdinal) continue;

                // IMAGE_IMPORT_BY_NAME: Hint (2 bytes) + nome
                string fn = Marshal.PtrToStringAnsi(module + (int)(entry & 0x7FFFFFFF) + 2);
                IntPtr replacement;
                if (!hooks.TryGetValue(fn, out replacement)) continue;

                IntPtr slot = module + iatRva + i * size;
                IntPtr previous = Marshal.ReadIntPtr(slot);
                if (previous == replacement) continue;
                WritePointer(slot, replacement);
                iatPatches.Add(new KeyValuePair<IntPtr, IntPtr>(slot, previous));
            }
        }

        static void WritePointer(IntPtr slot, IntPtr value)
        {
            uint old;
            Native.VirtualProtect(slot, (UIntPtr)IntPtr.Size, Native.PAGE_READWRITE, out old);
            Marshal.WriteIntPtr(slot, value);
            Native.VirtualProtect(slot, (UIntPtr)IntPtr.Size, old, out old);
        }

        // Paleta padrão do "Formato do editor" (COLORREF), na ordem de Theme.Palette
        static readonly int[] StandardPalette =
        {
            0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xC0C0C0,
            0x808080, 0xFF0000, 0x00FF00, 0xFFFF00, 0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF,
        };
    }

    static class Log
    {
        static readonly string path = Path.Combine(Settings.Folder, "SageEditor.log");

        public static void Info(string message) { Write("INFO  " + message); }
        public static void Error(Exception ex) { Write("ERRO  " + ex); }

        static void Write(string line)
        {
            try
            {
                Directory.CreateDirectory(Settings.Folder);
                File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + line + Environment.NewLine);
            }
            catch { }
        }
    }
}
