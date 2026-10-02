using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace SageEditor
{
    // Vários cursores na janela de código, como no VS Code:
    //   Alt+Clique                  acrescenta um cursor (ou tira, se já houver um ali)
    //   Ctrl+Alt+Seta Cima/Baixo    acrescenta um cursor na linha de cima/de baixo
    //   Shift+Alt+arrastar          seleção em coluna (um cursor por linha)
    //   Esc                         volta a um cursor
    // Com mais de um cursor, digitar, Backspace, Delete, Tab, setas (com Shift, seleção na
    // linha), Home e End valem para todos. Outras teclas (Enter, Ctrl+...) e cliques sem
    // Alt voltam a um cursor e seguem para o VBE.
    //
    // As edições são feitas pelo CodeModule (ReplaceLine; sem Desfazer). O VBE reformata a
    // linha gravada ("x=1" vira "x = 1"), então a coluna de cada cursor é recalculada pelos
    // caracteres que não são espaço antes dele. O VBE só mostra um cursor (o principal, o
    // último criado); os outros são desenhados aqui invertendo os pixels, só na área que o
    // VBE acabou de repintar (inverter duas vezes apagaria). A posição na tela (e a linha e
    // a coluna do mouse no Shift+Alt+arrastar) vem do cursor do VBE (GetCaretPos) mais a
    // largura do caractere e a altura da linha, medidas movendo o cursor dele ao entrar no
    // modo. Fonte de largura fixa, como a padrão.
    static class MultiCursor
    {
        sealed class Caret
        {
            public int Line, Col, Anchor; // colunas base 1; Anchor = Col: sem seleção
            public int Start { get { return Math.Min(Col, Anchor); } }
            public int End { get { return Math.Max(Col, Anchor); } }
            public Caret(int line, int col, int anchor) { Line = line; Col = col; Anchor = anchor; }
        }

        public static dynamic Vbe;
        static IntPtr vbeWindow;
        static readonly Native.SubclassProc subclassProc = SubclassProc;
        static readonly UIntPtr SubclassId = (UIntPtr)0x5A70;
        static readonly HashSet<IntPtr> attached = new HashSet<IntPtr>();

        static IntPtr window;                             // janela de código com vários cursores
        static readonly List<Caret> carets = new List<Caret>(); // o último é o principal
        static int charWidth, lineHeight;
        static bool dragging;
        static Caret dragAnchor;
        static int tick;

        const int WM_KEYDOWN = 0x0100, WM_SYSKEYDOWN = 0x0104, WM_CHAR = 0x0102, WM_SYSCHAR = 0x0106,
            WM_KILLFOCUS = 0x0008, WM_MOUSEMOVE = 0x0200, WM_LBUTTONDOWN = 0x0201, WM_LBUTTONUP = 0x0202;
        const int VK_BACK = 0x08, VK_TAB = 0x09, VK_SHIFT = 0x10, VK_CONTROL = 0x11, VK_MENU = 0x12, VK_ESCAPE = 0x1B,
            VK_END = 0x23, VK_HOME = 0x24, VK_LEFT = 0x25, VK_UP = 0x26, VK_RIGHT = 0x27, VK_DOWN = 0x28, VK_DELETE = 0x2E;
        const uint DSTINVERT = 0x00550009, PATINVERT = 0x005A0049;
        const int COLOR_WINDOW = 5, COLOR_HIGHLIGHT = 13;
        [DllImport("user32.dll")] static extern uint GetSysColor(int index);
        [DllImport("gdi32.dll")] static extern IntPtr SelectObject(IntPtr dc, IntPtr obj);
        const int TabWidth = 4; // Ferramentas > Opções > Largura da tabulação (padrão)

        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] struct MSG { public IntPtr hwnd; public int message; public IntPtr wParam, lParam; public uint time; public POINT pt; }
        const int SM_SWAPBUTTON = 23;
        [DllImport("user32.dll")] static extern short GetAsyncKeyState(int vk);
        [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
        [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT p);
        [DllImport("user32.dll")] static extern bool ScreenToClient(IntPtr hwnd, ref POINT p);
        [DllImport("user32.dll")] static extern bool PeekMessage(out MSG msg, IntPtr hwnd, uint min, uint max, uint remove);
        [DllImport("user32.dll")] static extern bool TranslateMessage(ref MSG msg);
        [DllImport("user32.dll")] static extern IntPtr DispatchMessage(ref MSG msg);
        [DllImport("user32.dll")] static extern void PostQuitMessage(int code);
        [DllImport("user32.dll")] static extern short GetKeyState(int vk);
        [DllImport("user32.dll")] static extern bool GetCaretPos(out POINT p);
        [DllImport("user32.dll")] static extern IntPtr SetCapture(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool ReleaseCapture();
        [DllImport("user32.dll")] static extern bool InvalidateRect(IntPtr hwnd, IntPtr rect, bool erase);
        [DllImport("user32.dll")] static extern bool UpdateWindow(IntPtr hwnd);
        [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
        [DllImport("gdi32.dll")] static extern bool PatBlt(IntPtr dc, int x, int y, int w, int h, uint rop);
        [DllImport("gdi32.dll")] static extern int SelectClipRgn(IntPtr dc, IntPtr region);

        static bool Active { get { return window != IntPtr.Zero && carets.Count > 1; } }

        // ------------------------------------------------------------------
        // Ciclo de vida (thread de interface do Excel)
        // ------------------------------------------------------------------

        public static void Start(IntPtr vbe) { vbeWindow = vbe; }

        // Subclassifica as janelas de código novas (verificação periódica)
        public static void Poll()
        {
            if (++tick % 7 != 0 || vbeWindow == IntPtr.Zero) return;
            Native.EnumChildWindows(vbeWindow, delegate(IntPtr h, IntPtr l)
            {
                if (!attached.Contains(h) && LineNumbers.IsCodePane(h) &&
                    Native.SetWindowSubclass(h, subclassProc, SubclassId, UIntPtr.Zero))
                    attached.Add(h);
                return true;
            }, IntPtr.Zero);
        }

        public static void Shutdown()
        {
            Exit();
            foreach (IntPtr h in attached) Native.RemoveWindowSubclass(h, subclassProc, SubclassId);
            attached.Clear();
        }

        // ------------------------------------------------------------------
        // Mensagens da janela de código
        // ------------------------------------------------------------------

        static IntPtr SubclassProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            try
            {
                if (msg == Native.WM_NCDESTROY)
                {
                    Native.RemoveWindowSubclass(hwnd, subclassProc, SubclassId);
                    attached.Remove(hwnd);
                    if (hwnd == window) { window = IntPtr.Zero; carets.Clear(); dragging = false; }
                }
                else if (Settings.MultiCursor)
                {
                    bool mine = hwnd == window;
                    switch (msg)
                    {
                        case Native.WM_PAINT:
                            if (mine && (Active || dragging))
                            {
                                IntPtr region = Native.CreateRectRgn(0, 0, 0, 0);
                                Native.GetUpdateRgn(hwnd, region, false);
                                IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
                                Draw(hwnd, region);
                                Native.DeleteObject(region);
                                return result;
                            }
                            break;

                        case WM_LBUTTONDOWN:
                            if (IsDown(VK_MENU) && IsDown(VK_SHIFT)) { StartColumn(hwnd, lParam); return IntPtr.Zero; }
                            if (IsDown(VK_MENU)) return AltClick(hwnd, msg, wParam, lParam);
                            if (Active) Exit();
                            break;

                        case WM_MOUSEMOVE:
                            if (mine && dragging) { UpdateColumn(lParam); return IntPtr.Zero; }
                            break;

                        case WM_LBUTTONUP:
                            if (mine && dragging) { EndColumn(); return IntPtr.Zero; }
                            break;

                        case WM_KEYDOWN:
                        case WM_SYSKEYDOWN:
                        {
                            int key = (int)wParam;
                            bool ctrl = IsDown(VK_CONTROL), alt = IsDown(VK_MENU), shift = IsDown(VK_SHIFT);
                            if (ctrl && alt && !shift && (key == VK_UP || key == VK_DOWN))
                            {
                                AddVertical(hwnd, key == VK_UP);
                                return IntPtr.Zero;
                            }
                            if (!mine || !Active) break;
                            if (key == VK_SHIFT || key == VK_CONTROL || key == VK_MENU || key == VK_MASK) break;
                            if (Key(key, ctrl, alt, shift)) return IntPtr.Zero;
                            Exit(); // outra tecla: volta a um cursor e segue para o VBE
                            break;
                        }

                        case WM_CHAR:
                            if (mine && Active)
                            {
                                char c = (char)(int)wParam;
                                if (c == '\t') { Insert(null); return IntPtr.Zero; }
                                if (c == '\b' || c == (char)27) return IntPtr.Zero; // tratados no WM_KEYDOWN
                                if (c >= ' ' && c != (char)127) { Insert(c.ToString()); return IntPtr.Zero; }
                            }
                            break;

                        case WM_SYSCHAR:
                            if (mine && Active) return IntPtr.Zero;
                            break;

                        case WM_KILLFOCUS:
                            if (mine && !dragging) Exit();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                Exit();
            }
            return Native.DefSubclassProc(hwnd, msg, wParam, lParam);
        }

        static bool IsDown(int vk) { return GetKeyState(vk) < 0; }

        // Com a tecla ou o clique consumidos, o VBE veria só "Alt pressionado e solto" e
        // ativaria a barra de menus. Uma tecla neutra no meio evita isso.
        // (Ela chega aqui como WM_KEYDOWN e não pode encerrar os vários cursores.)
        const int VK_MASK = 0xE8; // código de tecla sem uso

        static void MaskAlt()
        {
            keybd_event(VK_MASK, 0, 0, UIntPtr.Zero);
            keybd_event(VK_MASK, 0, 2 /* KEYEVENTF_KEYUP */, UIntPtr.Zero);
        }

        // ------------------------------------------------------------------
        // Criar cursores
        // ------------------------------------------------------------------

        // O clique segue para o VBE (que põe o cursor dele ali); o novo cursor é esse.
        static IntPtr AltClick(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam)
        {
            if (window != hwnd || carets.Count == 0) Begin(hwnd);
            IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
            MaskAlt();
            dynamic pane = Vbe.ActiveCodePane;
            int line, col;
            Selection(pane, out line, out col);
            Caret existing = carets.FirstOrDefault(c => c.Line == line && c.Col == col);
            if (existing != null && carets.Count > 1) carets.Remove(existing); // como o VS Code: tira
            else if (existing == null) carets.Add(new Caret(line, col, col));
            Calibrate(pane);
            Refresh(pane);
            return result;
        }

        static void AddVertical(IntPtr hwnd, bool up)
        {
            if (window != hwnd || carets.Count == 0) Begin(hwnd);
            MaskAlt();
            dynamic pane = Vbe.ActiveCodePane;
            dynamic module = pane.CodeModule;
            Caret primary = carets[carets.Count - 1];
            int line = up ? carets.Min(c => c.Line) - 1 : carets.Max(c => c.Line) + 1;
            if (line < 1 || line > (int)module.CountOfLines) return;
            int col = Math.Min(primary.Col, LineText(module, line).Length + 1);
            if (!carets.Any(c => c.Line == line && c.Col == col)) carets.Add(new Caret(line, col, col));
            Calibrate(pane);
            Refresh(pane);
        }

        // Começa com o cursor (e a seleção, se for numa linha só) que o VBE tem agora
        static void Begin(IntPtr hwnd)
        {
            if (window != IntPtr.Zero && window != hwnd) Exit();
            window = hwnd;
            carets.Clear();
            dynamic pane = Vbe.ActiveCodePane;
            int sl = 0, sc = 0, el = 0, ec = 0;
            pane.GetSelection(ref sl, ref sc, ref el, ref ec);
            carets.Add(sl == el ? new Caret(sl, ec, sc) : new Caret(el, ec, ec));
        }

        // Shift+Alt+arrastar. A linha e a coluna do mouse vêm das medidas (largura do
        // caractere, altura da linha) a partir do cursor do VBE: um clique sintético não
        // serve, porque com o botão pressionado o VBE entraria no arraste dele.
        static void StartColumn(IntPtr hwnd, IntPtr lParam)
        {
            if (window != IntPtr.Zero && window != hwnd) Exit();
            window = hwnd;
            MaskAlt();
            dynamic pane = Vbe.ActiveCodePane;
            Calibrate(pane);
            int line, col;
            PointToPosition(pane, lParam, out line, out col);
            dragAnchor = new Caret(line, col, col);
            dragging = true;
            carets.Clear();
            carets.Add(new Caret(line, col, col));
            Refresh(pane);
            DragLoop(hwnd);
        }

        // Laço próprio enquanto o botão estiver pressionado, como no arraste das abas: com a
        // captura ativa, o loop de mensagens do VBE não entrega WM_MOUSEMOVE. Pintura,
        // teclado e timers continuam sendo despachados; o mouse é descartado.
        static void DragLoop(IntPtr hwnd)
        {
            int button = GetSystemMetrics(SM_SWAPBUTTON) != 0 ? 0x02 : 0x01; // VK_RBUTTON / VK_LBUTTON
            SetCapture(hwnd);
            try
            {
                POINT last = new POINT { X = int.MinValue };
                MSG msg;
                while ((GetAsyncKeyState(button) & 0x8000) != 0 && dragging)
                {
                    while (PeekMessage(out msg, IntPtr.Zero, 0, 0, 1 /* PM_REMOVE */))
                    {
                        if (msg.message == 0x0012 /* WM_QUIT */) { PostQuitMessage((int)msg.wParam); return; }
                        if (msg.message >= 0x0200 && msg.message <= 0x020E) continue; // mouse
                        TranslateMessage(ref msg);
                        DispatchMessage(ref msg);
                    }
                    POINT p;
                    GetCursorPos(out p);
                    ScreenToClient(hwnd, ref p);
                    if (p.X != last.X || p.Y != last.Y)
                    {
                        last = p;
                        UpdateColumn((IntPtr)((p.Y << 16) | (p.X & 0xFFFF)));
                    }
                    System.Threading.Thread.Sleep(10);
                }
            }
            finally { EndColumn(); }
        }

        static void UpdateColumn(IntPtr lParam)
        {
            if (!dragging) return;
            dynamic pane = Vbe.ActiveCodePane;
            dynamic module = pane.CodeModule;
            int line, col;
            PointToPosition(pane, lParam, out line, out col);
            carets.Clear();
            int step = line >= dragAnchor.Line ? 1 : -1;
            for (int l = dragAnchor.Line; ; l += step)
            {
                int length = LineText(module, l).Length;
                carets.Add(new Caret(l, Math.Min(col, length + 1), Math.Min(dragAnchor.Col, length + 1)));
                if (l == line) break;
            }
            Refresh(pane);
        }

        static void EndColumn()
        {
            if (!dragging) return;
            dragging = false;
            ReleaseCapture();
            if (carets.Count <= 1) Exit(); // uma linha só: fica a seleção normal do VBE
        }

        // Ponto da área cliente -> linha e coluna, pela referência (ver SetReference)
        static void PointToPosition(dynamic pane, IntPtr lParam, out int line, out int col)
        {
            int x = (short)((long)lParam & 0xFFFF), y = (short)(((long)lParam >> 16) & 0xFFFF);
            int top = pane.TopLine;
            dynamic module = pane.CodeModule;
            line = Clamp(refLine + (int)Math.Floor((y - refY) / (double)Math.Max(1, lineHeight)) + (top - refTop),
                1, Math.Max(1, (int)module.CountOfLines));
            col = Clamp(refCol + (int)Math.Round((x - refX) / (double)Math.Max(1, charWidth)), 1, LineText(module, line).Length + 1);
        }

        static void Selection(dynamic pane, out int line, out int col)
        {
            int sl = 0, sc = 0, el = 0, ec = 0;
            pane.GetSelection(ref sl, ref sc, ref el, ref ec);
            line = el;
            col = ec;
        }

        static void Exit()
        {
            IntPtr old = window;
            bool wasActive = carets.Count > 1;
            if (dragging) { dragging = false; ReleaseCapture(); }
            carets.Clear();
            window = IntPtr.Zero;
            if (old != IntPtr.Zero && wasActive) InvalidateRect(old, IntPtr.Zero, false); // apaga os cursores desenhados
        }

        // ------------------------------------------------------------------
        // Teclas com vários cursores
        // ------------------------------------------------------------------

        // Tecla que produz um caractere (o WM_CHAR que vem depois é que insere)
        static bool IsTypingKey(int key)
        {
            return key == 0x20 || (key >= 0x30 && key <= 0x39) || (key >= 0x41 && key <= 0x5A) || // espaço, dígitos, letras
                (key >= 0x60 && key <= 0x6F) ||                                                   // teclado numérico
                (key >= 0xBA && key <= 0xC2) || (key >= 0xDB && key <= 0xDF) || key == 0xE2;      // pontuação (OEM)
        }

        static bool Key(int key, bool ctrl, bool alt, bool shift)
        {
            if (key == VK_ESCAPE) { Exit(); return true; }
            // Digitação (inclusive AltGr, que chega como Ctrl+Alt): fica com vários cursores
            if (IsTypingKey(key) && ctrl == alt) return true;
            if (ctrl || alt) return false;
            dynamic pane = Vbe.ActiveCodePane;
            dynamic module = pane.CodeModule;
            switch (key)
            {
                case VK_BACK: Edit(pane, Op.Backspace, null); return true;
                case VK_DELETE: Edit(pane, Op.Delete, null); return true;
                case VK_TAB: return true; // o WM_CHAR '\t' insere
                case VK_LEFT:
                case VK_RIGHT:
                    foreach (Caret c in carets)
                    {
                        int length = LineText(module, c.Line).Length;
                        if (shift) c.Col = Clamp(c.Col + (key == VK_LEFT ? -1 : 1), 1, length + 1);
                        else if (c.Anchor != c.Col) c.Col = c.Anchor = key == VK_LEFT ? c.Start : c.End;
                        else c.Col = c.Anchor = Clamp(c.Col + (key == VK_LEFT ? -1 : 1), 1, length + 1);
                    }
                    break;
                case VK_HOME:
                case VK_END:
                    foreach (Caret c in carets)
                    {
                        string text = LineText(module, c.Line);
                        int first = text.Length - text.TrimStart(' ').Length + 1;
                        int target = key == VK_END ? text.Length + 1 : (c.Col == first ? 1 : first);
                        c.Col = target;
                        if (!shift) c.Anchor = target;
                    }
                    break;
                case VK_UP:
                case VK_DOWN:
                    if (shift) return false; // seleção de várias linhas: volta a um cursor
                    int total = module.CountOfLines;
                    foreach (Caret c in carets)
                    {
                        int line = c.Line + (key == VK_UP ? -1 : 1);
                        if (line < 1 || line > total) continue;
                        c.Line = line;
                        c.Col = c.Anchor = Math.Min(c.Col, LineText(module, line).Length + 1);
                    }
                    break;
                default:
                    return false;
            }
            Dedupe();
            Refresh(pane);
            return true;
        }

        enum Op { Insert, Backspace, Delete }

        // text null: Tab (espaços até a próxima parada de tabulação, por cursor)
        static void Insert(string text)
        {
            Edit(Vbe.ActiveCodePane, Op.Insert, text);
        }

        static void Edit(dynamic pane, Op op, string text)
        {
            dynamic module = pane.CodeModule;
            foreach (IGrouping<int, Caret> group in carets.GroupBy(c => c.Line).ToList())
            {
                string original = LineText(module, group.Key);
                string s = original;
                List<Caret> done = new List<Caret>();
                // Da direita para a esquerda: cada edição só desloca os cursores já tratados
                foreach (Caret c in group.OrderByDescending(c => c.End))
                {
                    int start = Clamp(c.Start, 1, s.Length + 1), end = Clamp(c.End, 1, s.Length + 1);
                    int pos = start, removed = end - start;
                    string inserted = "";
                    if (op == Op.Insert) inserted = text ?? new string(' ', TabWidth - (start - 1) % TabWidth);
                    else if (removed == 0)
                    {
                        if (op == Op.Backspace) { if (start <= 1) continue; pos = start - 1; removed = 1; }
                        else { if (start > s.Length) continue; removed = 1; }
                    }
                    s = s.Substring(0, pos - 1) + inserted + s.Substring(pos - 1 + removed);
                    int delta = inserted.Length - removed;
                    foreach (Caret d in done) { d.Col += delta; d.Anchor += delta; }
                    c.Col = c.Anchor = pos + inserted.Length;
                    done.Add(c);
                }
                if (s != original) Write(module, group.Key, s, group.ToList());
            }
            Dedupe();
            Refresh(pane);
        }

        // Grava a linha e reposiciona os cursores dela na versão reformatada pelo VBE
        static void Write(dynamic module, int line, string raw, List<Caret> lineCarets)
        {
            var colKeys = lineCarets.Select(c => Tuple.Create(c, Key(raw, c.Col), Key(raw, c.Anchor))).ToList();
            module.ReplaceLine(line, raw);
            string formatted = LineText(module, line);
            foreach (var k in colKeys)
            {
                k.Item1.Col = Locate(formatted, k.Item2);
                k.Item1.Anchor = Locate(formatted, k.Item3);
            }
        }

        // Posição pela quantidade de caracteres que não são espaço antes dela, mais os
        // espaços logo antes
        static Tuple<int, int> Key(string s, int col)
        {
            int before = Math.Min(col - 1, s.Length), solid = 0, spaces = 0;
            for (int i = 0; i < before; i++)
            {
                if (s[i] == ' ') spaces++;
                else { solid++; spaces = 0; }
            }
            return Tuple.Create(solid, spaces);
        }

        static int Locate(string s, Tuple<int, int> key)
        {
            int i = 0, solid = 0;
            while (solid < key.Item1 && i < s.Length) { if (s[i] != ' ') solid++; i++; }
            int spaces = 0;
            while (spaces < key.Item2 && i < s.Length && s[i] == ' ') { spaces++; i++; }
            return i + 1;
        }

        static void Dedupe()
        {
            for (int i = carets.Count - 1; i >= 0; i--)
                for (int j = 0; j < i; j++)
                    if (carets[j].Line == carets[i].Line && carets[j].Col == carets[i].Col && carets[j].Anchor == carets[i].Anchor)
                    {
                        carets.RemoveAt(j); // fica o mais recente (o principal continua no fim)
                        i--;
                        break;
                    }
            if (carets.Count <= 1) Exit();
        }

        // ------------------------------------------------------------------
        // Tela
        // ------------------------------------------------------------------

        // Cursor do VBE no principal e repintura completa (os outros são desenhados no WM_PAINT)
        static void Refresh(dynamic pane)
        {
            if (carets.Count == 0 || window == IntPtr.Zero) return;
            Caret p = carets[carets.Count - 1];
            SetReference(pane, p.Line, p.Col);
            if (p.Anchor != p.Col) pane.SetSelection(p.Line, p.Start, p.Line, p.End);
            InvalidateRect(window, IntPtr.Zero, false);
            UpdateWindow(window);
        }

        // Posição na tela de uma linha e coluna conhecidas. Medida com o cursor do VBE sem
        // seleção: com seleção, o VBE esconde o cursor (GetCaretPos fica fora da tela).
        // A primeira linha visível compensa a rolagem vertical feita depois.
        static int refLine, refCol, refX, refY, refTop;

        static void SetReference(dynamic pane, int line, int col)
        {
            POINT p = CaretAt(pane, line, col);
            refLine = line; refCol = col; refX = p.X; refY = p.Y;
            refTop = pane.TopLine;
        }

        static Point ToPixel(int line, int col, int top)
        {
            return new Point(refX + (col - refCol) * charWidth, refY + (line - refLine - (top - refTop)) * lineHeight);
        }

        // Largura do caractere e altura da linha, movendo o cursor do VBE (perto de onde ele
        // está, para não rolar a tela)
        static void Calibrate(dynamic pane)
        {
            dynamic module = pane.CodeModule;
            int line, col;
            Selection(pane, out line, out col);
            int length = LineText(module, line).Length;
            POINT a = CaretAt(pane, line, col);
            if (length > 0)
            {
                int other = col > 1 ? col - 1 : col + 1;
                POINT b = CaretAt(pane, line, other);
                if (b.X != a.X) charWidth = Math.Abs(b.X - a.X);
            }
            int top = pane.TopLine, total = module.CountOfLines;
            int neighbor = line > top ? line - 1 : line + 1;
            if (neighbor >= 1 && neighbor <= total)
            {
                POINT c = CaretAt(pane, neighbor, Math.Min(col, LineText(module, neighbor).Length + 1));
                if (c.Y != a.Y) lineHeight = Math.Abs(c.Y - a.Y);
            }
            SetReference(pane, line, col);
        }

        static POINT CaretAt(dynamic pane, int line, int col)
        {
            pane.SetSelection(line, col, line, col);
            POINT p;
            GetCaretPos(out p);
            return p;
        }

        // Os outros cursores e as seleções deles, invertendo os pixels só na área repintada.
        // O principal com seleção também: o VBE desenha a seleção, mas esconde o cursor.
        static void Draw(IntPtr hwnd, IntPtr region)
        {
            if (carets.Count < 2 && !dragging) return;
            if (charWidth == 0 || lineHeight == 0) return;
            Caret p = carets[carets.Count - 1];
            int top = Vbe.ActiveCodePane.TopLine;
            // Seleção: XOR com (fundo ^ cor de seleção) deixa o fundo na cor de seleção do
            // tema, como a do VBE, e continua se desfazendo se for aplicado duas vezes
            Theme theme = ThemeEngine.Current;
            int back = theme != null ? theme.Window : (int)GetSysColor(COLOR_WINDOW);
            int highlight = theme != null ? theme.SysColors[COLOR_HIGHLIGHT] : (int)GetSysColor(COLOR_HIGHLIGHT);
            IntPtr brush = Native.CreateSolidBrush(back ^ highlight);
            IntPtr dc = Native.GetDC(hwnd);
            IntPtr oldBrush = SelectObject(dc, brush);
            try
            {
                SelectClipRgn(dc, region);
                foreach (Caret c in carets)
                {
                    bool selected = c.Anchor != c.Col;
                    Point at = ToPixel(c.Line, c.Col, top);
                    if (c != p || selected) PatBlt(dc, at.X, at.Y, 2, lineHeight, DSTINVERT);
                    if (c != p && selected)
                    {
                        // sem o trecho de 2 px do cursor (inverter duas vezes o apagaria)
                        Point left = ToPixel(c.Line, c.Start, top);
                        int skip = c.Col == c.Start ? 2 : 0;
                        PatBlt(dc, left.X + skip, left.Y, (c.End - c.Start) * charWidth - skip, lineHeight, PATINVERT);
                    }
                }
            }
            finally
            {
                SelectObject(dc, oldBrush);
                Native.ReleaseDC(hwnd, dc);
                Native.DeleteObject(brush);
            }
        }

        static string LineText(dynamic module, int line)
        {
            return (string)module.Lines(line, 1) ?? "";
        }

        static int Clamp(int v, int low, int high) { return Math.Max(low, Math.Min(high, v)); }
    }
}
