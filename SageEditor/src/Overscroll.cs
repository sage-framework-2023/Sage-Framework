using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageEditor
{
    // Rolar além da última linha, como o "editor.scrollBeyondLastLine" do VS Code. O VBE limita a
    // rolagem: o TopLine para quando a última linha chega embaixo da área de código. Quando a roda
    // do mouse pede mais do que o VBE rolou (já no fim), o Sage guarda um espaço extra para a
    // janela e o VBE passa a ver a área cliente mais baixa por esse tanto (desvio do GetClientRect
    // do VBE7.DLL): ele conta menos linhas visíveis, o limite do TopLine desce e o código sobe.
    // As barras de rolagem e os botões de modo de exibição são janelas filhas que o VBE posiciona
    // pela área que vê: depois de cada arranjo do VBE, o Sage os leva de volta ao fundo real, e o
    // vão entre o código e a barra horizontal é pintado com o fundo do código. Rolando para cima,
    // o espaço diminui antes; saindo do fim por outro caminho (barra de rolagem, Ctrl+Home, busca),
    // ele some. O clique no espaço vazio põe o cursor no fim do código.
    static class Overscroll
    {
        static readonly Dictionary<IntPtr, int> extra = new Dictionary<IntPtr, int>(); // pixels por janela
        static readonly Dictionary<IntPtr, int> lastTop = new Dictionary<IntPtr, int>(); // o TopLine que o Sage pôs
        static readonly Dictionary<IntPtr, int> realHeight = new Dictionary<IntPtr, int>(); // altura da área cliente real
        static bool fixing;
        static IntPtr arranging; // a janela que o VBE está arrumando a pedido do Sage

        const int WM_SIZE = 0x0005, WM_PAINT = 0x000F, WM_WINDOWPOSCHANGED = 0x0047, WM_MOUSEWHEEL = 0x020A,
            WM_LBUTTONDOWN = 0x0201, WM_LBUTTONDBLCLK = 0x0203;
        const uint SWP_NOSIZE = 0x1, SWP_NOMOVE = 0x2, SWP_NOZORDER = 0x4, SWP_NOACTIVATE = 0x10;

        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        [DllImport("user32.dll")] static extern bool ScreenToClient(IntPtr hwnd, ref POINT p);
        [DllImport("user32.dll")] static extern IntPtr SetFocus(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] static extern bool RedrawWindow(IntPtr hwnd, ref Native.RECT rect, IntPtr region, uint flags);
        [DllImport("user32.dll")] static extern bool InvalidateRect(IntPtr hwnd, ref Native.RECT rect, bool erase);
        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr hwnd, int index);
        const int GWL_STYLE = -16, SBS_VERT = 0x1, SBS_SIZEBOX = 0x8;
        const uint SWP_HIDEWINDOW = 0x80, SWP_SHOWWINDOW = 0x40, SWP_FRAMECHANGED = 0x20;
        [DllImport("gdi32.dll")] static extern bool StretchBlt(IntPtr dest, int x, int y, int w, int h, IntPtr src, int sx, int sy, int sw, int sh, uint rop);
        const uint SRCCOPY = 0x00CC0020;
        const uint RDW_INVALIDATE = 0x1, RDW_NOCHILDREN = 0x40, RDW_ALLCHILDREN = 0x80, RDW_UPDATENOW = 0x100;

        // ------------------------------------------------------------------
        // Desvio do GetClientRect (VBE7.DLL)
        // ------------------------------------------------------------------

        delegate bool GetClientRectProc(IntPtr hwnd, IntPtr rect);
        static GetClientRectProc origClientRect;
        static readonly GetClientRectProc hookClientRect = HookGetClientRect;

        public static void Hooks(IntPtr user32, Dictionary<string, IntPtr> hooks)
        {
            origClientRect = (GetClientRectProc)Marshal.GetDelegateForFunctionPointer(Native.GetProcAddress(user32, "GetClientRect"), typeof(GetClientRectProc));
            hooks["GetClientRect"] = Marshal.GetFunctionPointerForDelegate(hookClientRect);
            origSetPos = (SetWindowPosProc)Marshal.GetDelegateForFunctionPointer(Native.GetProcAddress(user32, "SetWindowPos"), typeof(SetWindowPosProc));
            hooks["SetWindowPos"] = Marshal.GetFunctionPointerForDelegate(hookSetPos);
            origMove = (MoveWindowProc)Marshal.GetDelegateForFunctionPointer(Native.GetProcAddress(user32, "MoveWindow"), typeof(MoveWindowProc));
            hooks["MoveWindow"] = Marshal.GetFunctionPointerForDelegate(hookMove);
            origDefer = (DeferWindowPosProc)Marshal.GetDelegateForFunctionPointer(Native.GetProcAddress(user32, "DeferWindowPos"), typeof(DeferWindowPosProc));
            hooks["DeferWindowPos"] = Marshal.GetFunctionPointerForDelegate(hookDefer);
            origInvalidate = (InvalidateRectProc)Marshal.GetDelegateForFunctionPointer(Native.GetProcAddress(user32, "InvalidateRect"), typeof(InvalidateRectProc));
            hooks["InvalidateRect"] = Marshal.GetFunctionPointerForDelegate(hookInvalidate);
            origRedraw = (RedrawWindowProc)Marshal.GetDelegateForFunctionPointer(Native.GetProcAddress(user32, "RedrawWindow"), typeof(RedrawWindowProc));
            hooks["RedrawWindow"] = Marshal.GetFunctionPointerForDelegate(hookRedraw);
        }

        // ------------------------------------------------------------------
        // Desvios de InvalidateRect e RedrawWindow (VBE7.DLL). A janela de código não tem
        // WS_CLIPCHILDREN: invalidá-la inteira invalida também os filhos, e as caixas de cima e os
        // botões de baixo piscavam a cada passo da roda. Durante um passo, só a área do texto.
        // ------------------------------------------------------------------

        delegate bool InvalidateRectProc(IntPtr hwnd, IntPtr rect, bool erase);
        delegate bool RedrawWindowProc(IntPtr hwnd, IntPtr rect, IntPtr region, uint flags);
        static InvalidateRectProc origInvalidate;
        static RedrawWindowProc origRedraw;
        static readonly InvalidateRectProc hookInvalidate = HookInvalidateRect;
        static readonly RedrawWindowProc hookRedraw = HookRedrawWindow;

        static bool HookInvalidateRect(IntPtr hwnd, IntPtr rect, bool erase)
        {
            if (hwnd != arranging || arranging == IntPtr.Zero) return origInvalidate(hwnd, rect, erase);
            Native.RECT text = Clip(hwnd, rect);
            return text.Bottom <= text.Top || text.Right <= text.Left || InvalidateRect(hwnd, ref text, erase);
        }

        static bool HookRedrawWindow(IntPtr hwnd, IntPtr rect, IntPtr region, uint flags)
        {
            if (hwnd != arranging || arranging == IntPtr.Zero || region != IntPtr.Zero) return origRedraw(hwnd, rect, region, flags);
            Native.RECT text = Clip(hwnd, rect);
            return text.Bottom <= text.Top || text.Right <= text.Left || RedrawWindow(hwnd, ref text, IntPtr.Zero, (flags & ~RDW_ALLCHILDREN) | RDW_NOCHILDREN);
        }

        // Para o ColorDecorators, que redesenha o código a cada rolagem
        public static void InvalidateText(IntPtr hwnd)
        {
            Native.RECT text = TextArea(hwnd);
            InvalidateRect(hwnd, ref text, false);
        }

        // O retângulo pedido (NULL: a janela toda) recortado à área do texto
        static Native.RECT Clip(IntPtr hwnd, IntPtr rect)
        {
            Native.RECT text = TextArea(hwnd);
            if (rect == IntPtr.Zero) return text;
            Native.RECT r = (Native.RECT)Marshal.PtrToStructure(rect, typeof(Native.RECT));
            text.Left = Math.Max(text.Left, r.Left);
            text.Top = Math.Max(text.Top, r.Top);
            text.Right = Math.Min(text.Right, r.Right);
            text.Bottom = Math.Min(text.Bottom, r.Bottom);
            return text;
        }

        // A área do texto e do espaço vazio: abaixo das caixas, acima da fileira de baixo e à
        // esquerda da barra vertical (17 px + 2 de margem), sem encostar em nenhum filho
        static Native.RECT TextArea(IntPtr hwnd)
        {
            Native.RECT c;
            Native.GetClientRect(hwnd, out c);
            int top = LineNumbers.TextTopOf(hwnd);
            return new Native.RECT { Left = 0, Top = top < 0 ? 32 : top, Right = c.Right - 19, Bottom = c.Bottom - 19 };
        }

        // ------------------------------------------------------------------
        // Desvios de SetWindowPos, MoveWindow e DeferWindowPos (VBE7.DLL): o VBE posiciona as
        // barras e os botões pela área que ele vê; com espaço extra, o pedido já sai com o fundo
        // real. Sem isso, o VBE os punha no lugar dele e o Sage os trazia de volta (piscava).
        // ------------------------------------------------------------------

        delegate bool SetWindowPosProc(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        delegate bool MoveWindowProc(IntPtr hwnd, int x, int y, int w, int h, bool repaint);
        delegate IntPtr DeferWindowPosProc(IntPtr info, IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        static SetWindowPosProc origSetPos;
        static MoveWindowProc origMove;
        static DeferWindowPosProc origDefer;
        static readonly SetWindowPosProc hookSetPos = HookSetWindowPos;
        static readonly MoveWindowProc hookMove = HookMoveWindow;
        static readonly DeferWindowPosProc hookDefer = HookDeferWindowPos;

        static bool HookSetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags)
        {
            if (extra.Count > 0 && Adjust(hwnd, ref y, ref cy, (flags & SWP_NOMOVE) == 0, (flags & SWP_NOSIZE) == 0)) flags &= ~SWP_HIDEWINDOW;
            if (AlreadyThere(hwnd, x, y, cx, cy, flags)) return true;
            return origSetPos(hwnd, after, x, y, cx, cy, flags);
        }

        static bool HookMoveWindow(IntPtr hwnd, int x, int y, int w, int h, bool repaint)
        {
            if (extra.Count > 0) Adjust(hwnd, ref y, ref h, true, true);
            if (AlreadyThere(hwnd, x, y, w, h, 0)) return true;
            return origMove(hwnd, x, y, w, h, repaint);
        }

        static IntPtr HookDeferWindowPos(IntPtr info, IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags)
        {
            if (extra.Count > 0 && Adjust(hwnd, ref y, ref cy, (flags & SWP_NOMOVE) == 0, (flags & SWP_NOSIZE) == 0)) flags &= ~SWP_HIDEWINDOW;
            if (AlreadyThere(hwnd, x, y, cx, cy, flags)) return info;
            return origDefer(info, hwnd, after, x, y, cx, cy, flags);
        }

        // No arranjo que o Sage pede (SetExtra), o VBE reposiciona todos os filhos, quase sempre no
        // mesmo lugar, e cada um é repintado: as caixas de cima e os botões de baixo piscavam a
        // cada passo da roda. Pedido que não muda nada é ignorado.
        static bool AlreadyThere(IntPtr hwnd, int x, int y, int cx, int cy, uint flags)
        {
            if (arranging == IntPtr.Zero || Native.GetParent(hwnd) != arranging) return false;
            if ((flags & (SWP_HIDEWINDOW | SWP_FRAMECHANGED)) != 0) return false;
            if ((flags & SWP_SHOWWINDOW) != 0 && !Native.IsWindowVisible(hwnd)) return false;
            Native.RECT r = ClientRectOf(arranging, hwnd);
            // A caixa de cima: o VBE a encolhe para 16 px e a devolve para 22 (a altura não conta)
            bool sameHeight = Native.ClassName(hwnd) == "ComboBox" || r.Bottom - r.Top == cy;
            return ((flags & SWP_NOMOVE) != 0 || (r.Left == x && r.Top == y)) &&
                ((flags & SWP_NOSIZE) != 0 || (r.Right - r.Left == cx && sameHeight));
        }

        // Uma barra ou os botões de uma janela de código com espaço extra: a fileira de baixo a
        // 2 px do fundo real, a barra vertical até o canto (como o VBE os põe com a área inteira).
        // true: é um deles (quem chamou não deixa o VBE escondê-lo)
        static bool Adjust(IntPtr child, ref int y, ref int cy, bool move, bool size)
        {
            try
            {
                IntPtr parent = Native.GetParent(child);
                int e;
                if (parent == IntPtr.Zero || !extra.TryGetValue(parent, out e) || e <= 0) return false;
                string cls = Native.ClassName(child);
                if (cls != "ScrollBar" && cls != "ObtbarWndClass") return false;
                Native.RECT c, now;
                Native.GetClientRect(parent, out c); // a área real (o desvio é só do VBE)
                Native.GetWindowRect(child, out now);
                int h = size ? cy : now.Bottom - now.Top;
                if (cls == "ScrollBar" && IsVertical(child))
                {
                    if (!size) return true;
                    int top = move ? y : ClientRectOf(parent, child).Top;
                    cy = Math.Max(20, c.Bottom - 18 - top);
                }
                else if (move) y = c.Bottom - 2 - h;
                return true;
            }
            catch (Exception) { return false; } // nunca atrapalhar o VBE
        }

        // Pelo estilo, e não pela forma: com a área pequena, o VBE pede a vertical com 14 px de
        // altura, do tamanho do canto (SBS_SIZEBOX)
        static bool IsVertical(IntPtr scrollBar)
        {
            int style = GetWindowLong(scrollBar, GWL_STYLE);
            return (style & SBS_SIZEBOX) == 0 && (style & SBS_VERT) != 0;
        }

        static bool HookGetClientRect(IntPtr hwnd, IntPtr rect)
        {
            bool ok = origClientRect(hwnd, rect);
            int e;
            if (ok && extra.Count > 0 && rect != IntPtr.Zero && extra.TryGetValue(hwnd, out e) && e > 0)
            {
                int top = Marshal.ReadInt32(rect, 4), bottom = Marshal.ReadInt32(rect, 12);
                Marshal.WriteInt32(rect, 12, Math.Max(top, bottom - e));
            }
            return ok;
        }

        // ------------------------------------------------------------------

        static int Extra(IntPtr hwnd)
        {
            int e;
            return extra.TryGetValue(hwnd, out e) ? e : 0;
        }

        // O maior espaço com a última linha ainda no topo: a área que o VBE vê tem de caber o
        // topo do texto (abaixo das caixas), uma linha e a fileira de baixo (17 px + 2 de margem).
        // Menor que isso, o VBE esconde a última linha e as barras de rolagem.
        static int MaxExtra(IntPtr hwnd, int clientHeight, int lineHeight)
        {
            int textTop = LineNumbers.TextTopOf(hwnd);
            if (textTop < 0) textTop = 32;
            // Duas linhas: a última e a vazia depois dela, onde o VBE deixa o cursor
            return Math.Max(0, clientHeight - (textTop + 2 * lineHeight + 19));
        }

        // Muda o espaço e faz o VBE arrumar a janela de novo (ele lê a área pelo GetClientRect)
        static void SetExtra(IntPtr hwnd, int pixels)
        {
            if (pixels <= 0) { extra.Remove(hwnd); lastTop.Remove(hwnd); } else extra[hwnd] = pixels;
            if (extra.Count > 50) foreach (IntPtr h in new List<IntPtr>(extra.Keys)) if (!Native.IsWindow(h)) extra.Remove(h);
            Native.RECT c;
            Native.GetClientRect(hwnd, out c); // o real (o desvio só vale para o VBE)
            realHeight[hwnd] = c.Bottom;
            IntPtr previous = arranging; // um passo da roda já pode estar em andamento
            arranging = hwnd;
            try { Native.SendMessage(hwnd, WM_SIZE, IntPtr.Zero, (IntPtr)((c.Bottom << 16) | (c.Right & 0xFFFF))); }
            finally { arranging = previous; }
            FixChildren(hwnd);
        }

        // Os filhos que o VBE pôs embaixo da área que ele vê voltam ao fundo real; a barra
        // vertical estica até lá
        static void FixChildren(IntPtr hwnd)
        {
            int e = Extra(hwnd);
            if (e == 0 || fixing) return;
            fixing = true;
            try
            {
                Native.RECT c;
                Native.GetClientRect(hwnd, out c);
                List<IntPtr> children = Children(hwnd);
                // 1. A fileira de baixo (barra horizontal, canto, botões) a 2 px do fundo real, como o
                // VBE a põe com a área inteira. Pela forma, e não pela posição: com a janela
                // redimensionada, o VBE não a move (a área que ele vê é a mesma)
                int target = c.Bottom - 18;
                foreach (IntPtr child in children)
                {
                    string cls = Native.ClassName(child);
                    if (cls != "ScrollBar" && cls != "ObtbarWndClass") continue;
                    Native.RECT r = ClientRectOf(hwnd, child);
                    int w = r.Right - r.Left, h = r.Bottom - r.Top;
                    if (cls == "ScrollBar" && IsVertical(child)) continue;
                    int top = c.Bottom - 2 - h;
                    if (r.Top != top) SetWindowPos(child, IntPtr.Zero, r.Left, top, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
                    if (cls == "ScrollBar" && (GetWindowLong(child, GWL_STYLE) & SBS_SIZEBOX) != 0 && Native.IsWindowVisible(child)) target = top; // o canto
                }
                // 2. A barra vertical vai até o canto (o VBE a encurta de novo quando o TopLine muda)
                foreach (IntPtr child in children)
                {
                    if (Native.ClassName(child) != "ScrollBar" || !IsVertical(child)) continue;
                    Native.RECT r = ClientRectOf(hwnd, child);
                    int w = r.Right - r.Left, h = r.Bottom - r.Top;
                    if (r.Bottom != target && target > r.Top + 20)
                        SetWindowPos(child, IntPtr.Zero, 0, 0, w, target - r.Top, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                }
            }
            finally { fixing = false; }
        }

        // Filhos no lugar e só a área do texto redesenhada (sem os filhos: as caixas de cima piscavam)
        static void Refresh(IntPtr hwnd)
        {
            FixChildren(hwnd);
            Native.RECT text = TextArea(hwnd);
            RedrawWindow(hwnd, ref text, IntPtr.Zero, RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOCHILDREN);
        }

        static List<IntPtr> Children(IntPtr hwnd)
        {
            List<IntPtr> list = new List<IntPtr>();
            Native.EnumChildWindows(hwnd, delegate(IntPtr h, IntPtr l)
            {
                if (Native.GetParent(h) == hwnd) list.Add(h);
                return true;
            }, IntPtr.Zero);
            return list;
        }

        static Native.RECT ClientRectOf(IntPtr parent, IntPtr child)
        {
            Native.RECT r;
            Native.GetWindowRect(child, out r);
            POINT a = new POINT { X = r.Left, Y = r.Top }, b = new POINT { X = r.Right, Y = r.Bottom };
            ScreenToClient(parent, ref a);
            ScreenToClient(parent, ref b);
            return new Native.RECT { Left = a.X, Top = a.Y, Right = b.X, Bottom = b.Y };
        }

        // O vão entre o código e a barra horizontal (coordenadas da área cliente real): do lugar
        // onde o VBE pôs a fileira de baixo (e a linha acima dela) até onde ela está de verdade
        static bool Gap(IntPtr hwnd, out Native.RECT gap)
        {
            gap = new Native.RECT();
            int e = Extra(hwnd);
            if (e == 0) return false;
            Native.RECT c;
            Native.GetClientRect(hwnd, out c);
            int right = c.Right, rowTop = c.Bottom;
            foreach (IntPtr child in Children(hwnd))
            {
                string cls = Native.ClassName(child);
                if ((cls != "ScrollBar" && cls != "ObtbarWndClass") || !Native.IsWindowVisible(child)) continue;
                Native.RECT r = ClientRectOf(hwnd, child);
                int w = r.Right - r.Left, h = r.Bottom - r.Top;
                if (cls == "ScrollBar" && h > w && h > 20) right = Math.Min(right, r.Left);   // vertical
                else if (r.Top > c.Bottom - 40) rowTop = Math.Min(rowTop, r.Top);              // fileira de baixo
            }
            gap.Left = 0;
            gap.Right = right;
            gap.Bottom = rowTop;
            gap.Top = Math.Max(0, rowTop - e - 2);
            return gap.Bottom > gap.Top && gap.Right > 0;
        }

        static void PaintGap(IntPtr hwnd)
        {
            Native.RECT gap;
            if (!Gap(hwnd, out gap)) return;
            Theme t = ThemeEngine.Effective;
            IntPtr dc = Native.GetDC(hwnd);
            try
            {
                Native.FillRect(dc, ref gap, t.Brush(t.Window));
                // A margem de indicadores (pontos de interrupção) continua pelo espaço vazio: uma
                // linha de pixels do código, logo acima, esticada. Os primeiros 30 px são margem e
                // recuo, sem texto.
                int source = gap.Top - 4, margin = Math.Min(30, gap.Right);
                if (source > 0 && margin > 0) StretchBlt(dc, 0, gap.Top, margin, gap.Bottom - gap.Top, dc, 0, source, margin, 1, SRCCOPY);
            }
            finally { Native.ReleaseDC(hwnd, dc); }
        }

        // ------------------------------------------------------------------
        // Mensagens da janela de código (ThemeEngine). true: tratada (result é a resposta)
        // ------------------------------------------------------------------

        public static bool Handle(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, out IntPtr result)
        {
            result = IntPtr.Zero;
            if (msg == WM_MOUSEWHEEL) return LineNumbers.IsCodePane(hwnd) && Wheel(hwnd, msg, wParam, lParam, out result);
            if (Extra(hwnd) == 0) return false;
            switch (msg)
            {
                case WM_SIZE:
                {
                    // A janela mudou de altura: a diferença vai para o espaço, e o VBE continua vendo
                    // a mesma área (o código fica onde está, como no VS Code), de zero até a última
                    // linha no topo
                    int now = (int)(((long)lParam >> 16) & 0xFFFF), before;
                    int height = LineNumbers.LineHeightOf(hwnd);
                    if (realHeight.TryGetValue(hwnd, out before) && now != before && height > 0)
                    {
                        int limit = MaxExtra(hwnd, now, height);
                        int e = Math.Min(limit, Math.Max(0, Extra(hwnd) + now - before));
                        if (e > 0) extra[hwnd] = e; else { extra.Remove(hwnd); lastTop.Remove(hwnd); }
                    }
                    realHeight[hwnd] = now;
                    result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                    FixChildren(hwnd);
                    return true;
                }
                case WM_WINDOWPOSCHANGED:
                    // O VBE arruma os filhos pela área que ele vê; eles voltam ao fundo real
                    result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                    FixChildren(hwnd);
                    return true;
                case WM_PAINT:
                    result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                    PaintGap(hwnd);
                    LineNumbers.Paint(hwnd, ThemeEngine.Effective);
                    return true;
                case WM_LBUTTONDOWN:
                case WM_LBUTTONDBLCLK:
                {
                    Native.RECT gap;
                    int y = (short)(((long)lParam >> 16) & 0xFFFF), x = (short)((long)lParam & 0xFFFF);
                    if (!Gap(hwnd, out gap) || y < gap.Top || y >= gap.Bottom || x >= gap.Right) return false;
                    GoToEnd(hwnd);
                    return true;
                }
            }
            return false;
        }

        static bool Wheel(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, out IntPtr result)
        {
            result = IntPtr.Zero;
            int height = LineNumbers.LineHeightOf(hwnd);
            if (height <= 0 || ((long)wParam & 0x0004 /* MK_SHIFT */) != 0) return false;
            int delta = (short)(((long)wParam >> 16) & 0xFFFF);
            int perNotch = SystemInformation.MouseWheelScrollLines;
            if (delta == 0 || perNotch <= 0) return false; // "uma tela por vez": fica com o VBE
            int lines = Math.Max(1, Math.Abs(delta) * perNotch / 120);
            dynamic pane = LineNumbers.PaneOf(hwnd);
            if (pane == null) return false;
            int e = Extra(hwnd);
            try
            {
                if (delta > 0)
                {
                    // Para cima: primeiro o espaço diminui (o código desce junto)
                    if (e == 0) return false;
                    arranging = hwnd;
                    int before = pane.TopLine;
                    int remove = Math.Min(e, lines * height);
                    SetExtra(hwnd, e - remove);
                    pane.TopLine = Math.Max(1, before - remove / height);
                    Refresh(hwnd);
                    if (Extra(hwnd) > 0) lastTop[hwnd] = pane.TopLine;
                    return true;
                }

                // Para baixo: o VBE rola o que puder; o que faltar vira espaço vazio
                int top = pane.TopLine;
                int total = pane.CodeModule.CountOfLines;
                if (e > 0) arranging = hwnd;
                result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                // Nunca além da última linha (o VBE iria até a linha vazia depois dela)
                if (e > 0 && (int)pane.TopLine > total) { pane.TopLine = total; lastTop[hwnd] = total; return true; }
                int rolled = (int)pane.TopLine - top;
                int missing = lines - rolled;
                if (missing <= 0) return true;
                Native.RECT client;
                Native.GetClientRect(hwnd, out client);
                // Até a última linha ficar no topo (e não além: o VBE esconderia as barras)
                int limit = MaxExtra(hwnd, client.Bottom, height);
                int grow = Math.Min(missing, (limit - e) / height) * height; // linhas inteiras
                if (grow <= 0) return true;
                int topBefore = pane.TopLine;
                arranging = hwnd;
                SetExtra(hwnd, e + grow);
                pane.TopLine = Math.Min(total, topBefore + grow / height);
                Refresh(hwnd);
                lastTop[hwnd] = pane.TopLine;
                return true;
            }
            catch (Exception ex) { Log.Error(ex); return false; } // VBE ocupado
            finally { arranging = IntPtr.Zero; }
        }

        // Clique no espaço vazio: o cursor vai para o fim da última linha, sem rolar
        static void GoToEnd(IntPtr hwnd)
        {
            try
            {
                dynamic pane = LineNumbers.PaneOf(hwnd);
                if (pane == null) return;
                int total = pane.CodeModule.CountOfLines;
                if (total < 1) return;
                int column = ((string)pane.CodeModule.Lines(total, 1)).Length + 1;
                int top = pane.TopLine;
                pane.SetSelection(total, column, total, column);
                pane.TopLine = top;
                SetFocus(hwnd);
            }
            catch (Exception ex) { Log.Error(ex); }
        }

        // Verificação periódica: a última linha saiu da vista. Com o mesmo TopLine que o Sage pôs,
        // foi a janela que diminuiu: o espaço encolhe só o bastante para a última linha voltar, sem
        // rolar (como no VS Code). Com outro TopLine (barra de rolagem, Ctrl+Home, busca), o espaço some.
        public static void Poll(dynamic pane, IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || Extra(hwnd) == 0) return;
            try
            {
                int top = pane.TopLine, visible = pane.CountOfVisibleLines, total = pane.CodeModule.CountOfLines;
                int height = LineNumbers.LineHeightOf(hwnd), mine;
                if (top + visible - 1 >= total) { FixChildren(hwnd); return; } // o VBE pode ter encurtado a barra vertical de novo
                if (height > 0 && lastTop.TryGetValue(hwnd, out mine) && mine == top)
                {
                    int shrink = (total - (top + visible - 1)) * height;
                    SetExtra(hwnd, Extra(hwnd) - shrink);
                    pane.TopLine = top;
                    Refresh(hwnd);
                    if (Extra(hwnd) > 0) lastTop[hwnd] = top;
                    return;
                }
                SetExtra(hwnd, 0);
                Refresh(hwnd);
            }
            catch (Exception) { } // VBE ocupado
        }
    }
}
