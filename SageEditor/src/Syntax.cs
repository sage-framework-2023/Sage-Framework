using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace SageEditor
{
    // Cores de sintaxe além das três categorias do VBE (palavra-chave, comentário,
    // texto normal), no estilo do VS Code.
    sealed class SyntaxColors
    {
        public int Function, Variable, String, Number, Type, Control;

        public static SyntaxColors Create(string function, string variable, string text, string number, string type, string control)
        {
            SyntaxColors s = new SyntaxColors();
            s.Function = Theme.Ref(function);
            s.Variable = Theme.Ref(variable);
            s.String = Theme.Ref(text);
            s.Number = Theme.Ref(number);
            s.Type = Theme.Ref(type);
            s.Control = Theme.Ref(control);
            return s;
        }
    }

    // O VBE desenha cada linha do código com ExtTextOutA/TextOutA, um trecho por cor
    // (o texto normal inteiro num trecho só). Aqui cada trecho de texto normal ou de
    // palavra-chave é dividido em tokens e redesenhado com a cor de cada um.
    static class Syntax
    {
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        delegate bool ExtTextOutProc(IntPtr hdc, int x, int y, uint options, IntPtr rect, IntPtr text, uint count, IntPtr dx);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        delegate bool TextOutProc(IntPtr hdc, int x, int y, IntPtr text, int count);

        [DllImport("gdi32.dll")] static extern uint GetTextAlign(IntPtr hdc);
        [DllImport("gdi32.dll")] static extern int GetTextColor(IntPtr hdc);
        [DllImport("gdi32.dll")] static extern bool GetTextExtentPoint32A(IntPtr hdc, IntPtr text, int count, out Size size);
        [StructLayout(LayoutKind.Sequential)] struct Size { public int Width, Height; }
        [StructLayout(LayoutKind.Sequential)] struct Point { public int X, Y; }
        [DllImport("gdi32.dll")] static extern bool GetCurrentPositionEx(IntPtr hdc, out Point point);

        const uint ETO_OPAQUE = 0x2, ETO_CLIPPED = 0x4, ETO_GLYPH_INDEX = 0x10, ETO_PDY = 0x2000;
        const uint TA_UPDATECP = 0x1, TA_RIGHT = 0x2, TA_CENTER = 0x6;

        static ExtTextOutProc origExt;
        static TextOutProc origTextOut;
        static readonly ExtTextOutProc hookExt = HookExtTextOut;
        static readonly TextOutProc hookTextOut = HookTextOut;

        public static void Hooks(IntPtr gdi32, Dictionary<string, IntPtr> hooks)
        {
            origExt = (ExtTextOutProc)Marshal.GetDelegateForFunctionPointer(Native.GetProcAddress(gdi32, "ExtTextOutA"), typeof(ExtTextOutProc));
            origTextOut = (TextOutProc)Marshal.GetDelegateForFunctionPointer(Native.GetProcAddress(gdi32, "TextOutA"), typeof(TextOutProc));
            hooks["ExtTextOutA"] = Marshal.GetFunctionPointerForDelegate(hookExt);
            hooks["TextOutA"] = Marshal.GetFunctionPointerForDelegate(hookTextOut);
        }

        // Cor original pedida pelo VBE no último SetTextColor (antes de virar cor do tema)
        [ThreadStatic] static int requestedColor;
        public static void NoteTextColor(int original) { requestedColor = original; }

        // ------------------------------------------------------------------
        // Desenho
        // ------------------------------------------------------------------

        static bool HookTextOut(IntPtr hdc, int x, int y, IntPtr text, int count)
        {
            if (count > 0 && Draw(hdc, x, y, 0, IntPtr.Zero, text, count, IntPtr.Zero)) return true;
            return origTextOut(hdc, x, y, text, count);
        }

        static bool HookExtTextOut(IntPtr hdc, int x, int y, uint options, IntPtr rect, IntPtr text, uint count, IntPtr dx)
        {
            if (count > 0 && (options & (ETO_GLYPH_INDEX | ETO_PDY)) == 0 &&
                Draw(hdc, x, y, options, rect, text, (int)count, dx)) return true;
            return origExt(hdc, x, y, options, rect, text, count, dx);
        }

        enum Kind { Normal, Keyword }

        // Contexto da linha que está sendo desenhada (os trechos vêm da esquerda para a direita)
        [ThreadStatic] static IntPtr lineWindow;
        [ThreadStatic] static int lineY, lineX;
        [ThreadStatic] static string prev1, prev2;
        [ThreadStatic] static bool typeChain;

        static bool Draw(IntPtr hdc, int x, int y, uint options, IntPtr rect, IntPtr text, int count, IntPtr dx)
        {
            try
            {
                Theme t = ThemeEngine.SyntaxTheme(hdc);
                if (t == null) return false;
                IntPtr hwnd = Native.WindowFromDC(hdc);
                if (!IsCodePane(hwnd)) return false;

                // Com TA_UPDATECP o VBE passa x = y = 0; a posição real é a corrente do DC
                int px = x, py = y;
                if ((GetTextAlign(hdc) & TA_UPDATECP) != 0)
                {
                    Point cp;
                    if (GetCurrentPositionEx(hdc, out cp)) { px = cp.X; py = cp.Y; }
                }
                LineNumbers.Observe(hwnd, px, py);
                if (t.Syntax == null) return false;

                Kind kind;
                int requested = requestedColor;
                if (requested == 0x800000) kind = Kind.Keyword;                          // Azul-marinho padrão
                else if (requested == 0 || requested == t.WindowText) kind = Kind.Normal; // Automático / preto
                else { ResetLine(hwnd, py, px); return false; }                          // comentário, seleção...

                if (hwnd != lineWindow || py != lineY || px < lineX) ResetLine(hwnd, py, px);
                lineX = px;

                byte[] bytes = new byte[count];
                Marshal.Copy(text, bytes, 0, count);
                List<Segment> segments = kind == Kind.Keyword ? Keywords(bytes, t.Syntax) : Tokens(bytes, t.Syntax);
                if (segments.Count == 0) return false;

                uint align = GetTextAlign(hdc);
                if ((align & (TA_RIGHT | TA_CENTER)) != 0) return false;
                // Com TA_UPDATECP (o VBE usa) cada chamada continua de onde a anterior parou
                bool sequential = (align & TA_UPDATECP) != 0;

                // Fundo do trecho inteiro uma vez; depois cada pedaço na sua cor
                if ((options & ETO_OPAQUE) != 0 && rect != IntPtr.Zero)
                    origExt(hdc, x, y, ETO_OPAQUE | (options & ETO_CLIPPED), rect, IntPtr.Zero, 0, IntPtr.Zero);

                int baseColor = GetTextColor(hdc);
                uint segOptions = options & ~ETO_OPAQUE;
                int pos = 0;
                foreach (Segment s in segments)
                {
                    if (s.Start > pos) Piece(hdc, x, y, segOptions, rect, text, pos, s.Start - pos, dx, baseColor, sequential);
                    Piece(hdc, x, y, segOptions, rect, text, s.Start, s.Length, dx, s.Color, sequential);
                    pos = s.Start + s.Length;
                }
                if (pos < count) Piece(hdc, x, y, segOptions, rect, text, pos, count - pos, dx, baseColor, sequential);
                Native.SetTextColor(hdc, baseColor);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                return false;
            }
        }

        static void Piece(IntPtr hdc, int x, int y, uint options, IntPtr rect, IntPtr text, int start, int length, IntPtr dx, int color, bool sequential)
        {
            int offset = 0;
            if (start > 0 && !sequential)
            {
                if (dx != IntPtr.Zero)
                    for (int i = 0; i < start; i++) offset += Marshal.ReadInt32(dx, i * 4);
                else
                {
                    Size size;
                    GetTextExtentPoint32A(hdc, text, start, out size);
                    offset = size.Width;
                }
            }
            Native.SetTextColor(hdc, color);
            origExt(hdc, x + offset, y, options, rect, text + start, (uint)length,
                dx == IntPtr.Zero ? IntPtr.Zero : dx + start * 4);
        }

        static void ResetLine(IntPtr hwnd, int y, int x)
        {
            lineWindow = hwnd;
            lineY = y;
            lineX = x;
            prev1 = prev2 = null;
            typeChain = false;
        }

        // Só janelas de código (VbaWindow dentro do MDIClient), não a Verificação imediata
        static readonly Dictionary<IntPtr, bool> codePanes = new Dictionary<IntPtr, bool>();

        static bool IsCodePane(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return false;
            bool result;
            lock (codePanes)
            {
                if (!codePanes.TryGetValue(hwnd, out result))
                {
                    result = Native.ClassName(hwnd) == "VbaWindow" && Native.ClassName(Native.GetParent(hwnd)) == "MDIClient";
                    if (codePanes.Count > 200) codePanes.Clear();
                    codePanes[hwnd] = result;
                }
            }
            return result;
        }

        // ------------------------------------------------------------------
        // Tokens
        // ------------------------------------------------------------------

        struct Segment
        {
            public int Start, Length, Color;
            public Segment(int start, int length, int color) { Start = start; Length = length; Color = color; }
        }

        static bool IsLetter(byte c) { return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_' || c >= 0xC0; }
        static bool IsDigit(byte c) { return c >= '0' && c <= '9'; }

        static List<Segment> Tokens(byte[] b, SyntaxColors colors)
        {
            List<Segment> list = new List<Segment>();
            int n = b.Length, i = 0;
            while (i < n)
            {
                byte c = b[i];
                if (c == '"')
                {
                    int j = i + 1;
                    while (j < n)
                    {
                        if (b[j] == '"') { if (j + 1 < n && b[j + 1] == '"') { j += 2; continue; } j++; break; }
                        j++;
                    }
                    list.Add(new Segment(i, j - i, colors.String));
                    Push(null);
                    i = j;
                }
                else if (c == '\'')
                {
                    break; // comentário no fim da linha vem com a cor própria
                }
                else if (IsDigit(c) || (c == '&' && i + 1 < n && (b[i + 1] == 'H' || b[i + 1] == 'h' || b[i + 1] == 'O' || b[i + 1] == 'o')))
                {
                    int j = i + 1;
                    bool hex = c == '&';
                    if (hex) j++;
                    while (j < n && (IsDigit(b[j]) || b[j] == '.' || (hex && ((b[j] >= 'A' && b[j] <= 'F') || (b[j] >= 'a' && b[j] <= 'f'))))) j++;
                    while (j < n && (b[j] == '&' || b[j] == '#' || b[j] == '!' || b[j] == '@' || b[j] == '%' || b[j] == '^')) j++;
                    list.Add(new Segment(i, j - i, colors.Number));
                    Push(null);
                    i = j;
                }
                else if (IsLetter(c))
                {
                    int j = i + 1;
                    while (j < n && (IsLetter(b[j]) || IsDigit(b[j]))) j++;
                    int end = j;
                    if (j < n && (b[j] == '$' || b[j] == '%' || b[j] == '&' || b[j] == '#' || b[j] == '!' || b[j] == '@') &&
                        (j + 1 >= n || !IsLetter(b[j + 1]))) j++;
                    string word = Encoding.Default.GetString(b, i, end - i);
                    bool afterDot = i > 0 && b[i - 1] == '.';
                    int k = j;
                    while (k < n && b[k] == ' ') k++;
                    bool call = k < n && b[k] == '(';
                    list.Add(new Segment(i, j - i, Classify(word, afterDot, call, colors)));
                    i = j;
                }
                else
                {
                    if (c != ' ' && c != '\t')
                    {
                        if (c != '.') typeChain = false;
                        if (c != '.') Push(null);
                    }
                    i++;
                }
            }
            return list;
        }

        static int Classify(string word, bool afterDot, bool call, SyntaxColors colors)
        {
            string p1 = prev1, p2 = prev2;
            Push(word);

            if (p1 != null && !afterDot)
            {
                if (Eq(p1, "Sub") || Eq(p1, "Function") || Eq(p1, "Enum") || Eq(p1, "Event") ||
                    ((Eq(p1, "Get") || Eq(p1, "Let") || Eq(p1, "Set")) && Eq(p2, "Property")))
                {
                    Learn(word);
                    return colors.Function;
                }
                if (Eq(p1, "As") || Eq(p1, "New") || Eq(p1, "Implements") || Eq(p1, "Type"))
                {
                    typeChain = true;
                    return colors.Type;
                }
            }
            if (afterDot && typeChain) return colors.Type;
            typeChain = false;

            if (KnownProcedure(word) && !afterDot) return colors.Function;
            // Funções nativas só como chamada: "str", "left", "day"... são nomes comuns de variáveis
            if (call && !afterDot && Builtins.Contains(word)) return colors.Function;
            if (afterDot && call) return colors.Function;
            return colors.Variable;
        }

        static void Push(string word)
        {
            prev2 = prev1;
            prev1 = word;
        }

        static bool Eq(string a, string b) { return a != null && string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }

        static readonly HashSet<string> ControlWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "If", "Then", "Else", "ElseIf", "For", "Each", "In", "To", "Step", "Next", "Do", "Loop", "While", "Wend",
            "Until", "Select", "Case", "Exit", "GoTo", "GoSub", "Return", "Resume", "With", "Call", "Stop", "On", "Error"
        };

        static readonly HashSet<string> TypeWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "String", "Long", "Integer", "Boolean", "Double", "Single", "Currency", "Date", "Byte", "Variant",
            "Object", "LongPtr", "LongLong", "Decimal", "Any", "Collection"
        };

        // Trecho de palavras-chave: controle de fluxo em roxo, tipos em verde-água
        static List<Segment> Keywords(byte[] b, SyntaxColors colors)
        {
            List<Segment> list = new List<Segment>();
            List<KeyValuePair<int, int>> words = new List<KeyValuePair<int, int>>();
            int n = b.Length, i = 0;
            while (i < n)
            {
                if (IsLetter(b[i]))
                {
                    int j = i + 1;
                    while (j < n && (IsLetter(b[j]) || IsDigit(b[j]))) j++;
                    words.Add(new KeyValuePair<int, int>(i, j - i));
                    i = j;
                }
                else i++;
            }

            for (int w = 0; w < words.Count; w++)
            {
                string word = Encoding.Default.GetString(b, words[w].Key, words[w].Value);
                string next = w + 1 < words.Count ? Encoding.Default.GetString(b, words[w + 1].Key, words[w + 1].Value) : null;
                bool control = ControlWords.Contains(word) ||
                    (Eq(word, "End") && (next == null || Eq(next, "If") || Eq(next, "Select") || Eq(next, "With")));
                if (Eq(word, "Case") && Eq(next, "Else")) control = true;
                if (control)
                    list.Add(new Segment(words[w].Key, words[w].Value, colors.Control));
                else if (TypeWords.Contains(word) && Eq(prev1, "As"))
                    list.Add(new Segment(words[w].Key, words[w].Value, colors.Type));
                Push(word);
                typeChain = false;
            }
            return list;
        }

        // ------------------------------------------------------------------
        // Procedimentos conhecidos (declarados nos módulos abertos)
        // ------------------------------------------------------------------

        static readonly HashSet<string> procedures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        static readonly Regex Declaration = new Regex(
            @"^[ \t]*(?:(?:Public|Private|Friend|Static)[ \t]+)*(?:Sub|Function|Property[ \t]+(?:Get|Let|Set)|Declare[ \t]+(?:PtrSafe[ \t]+)?(?:Sub|Function))[ \t]+([A-Za-z_À-ÿ][A-Za-z0-9_À-ÿ]*)",
            RegexOptions.Multiline | RegexOptions.IgnoreCase);

        static bool KnownProcedure(string word)
        {
            lock (procedures) return procedures.Contains(word);
        }

        static void Learn(string word)
        {
            lock (procedures) procedures.Add(word);
        }

        // Lê os módulos de todos os projetos que não estão protegidos
        public static void Scan(dynamic vbe)
        {
            int found = 0;
            foreach (dynamic project in vbe.VBProjects)
            {
                try
                {
                    if (project.Protection != 0) continue;
                    foreach (dynamic component in project.VBComponents)
                    {
                        dynamic module = component.CodeModule;
                        int lines = module.CountOfLines;
                        if (lines == 0) continue;
                        string code = module.Lines(1, lines);
                        foreach (Match m in Declaration.Matches(code)) { Learn(m.Groups[1].Value); found++; }
                    }
                }
                catch (Exception) { }
            }
            Log.Info("Sintaxe: " + found + " procedimento(s) conhecido(s)");
        }

        static readonly HashSet<string> Builtins = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Abs", "Array", "Asc", "AscW", "Atn", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "Chr", "ChrW",
            "CInt", "CLng", "CLngPtr", "Choose", "Cos", "CreateObject", "CSng", "CStr", "CurDir", "CVar", "DateAdd",
            "DateDiff", "DatePart", "DateSerial", "DateValue", "Day", "Dir", "DoEvents", "Environ", "Exp", "FileLen",
            "Filter", "Fix", "Format", "FormatNumber", "FormatPercent", "FreeFile", "GetObject", "Hex", "Hour", "IIf",
            "InputBox", "InStr", "InStrRev", "Int", "IsArray", "IsDate", "IsEmpty", "IsError", "IsMissing", "IsNull",
            "IsNumeric", "IsObject", "Join", "LCase", "Left", "Len", "Log", "LTrim", "Mid", "Minute", "Month",
            "MonthName", "MsgBox", "Now", "Oct", "Replace", "RGB", "Right", "Rnd", "Round", "RTrim", "Second", "Sgn",
            "Shell", "Sin", "Space", "Split", "Sqr", "Str", "StrComp", "StrConv", "StrReverse", "Switch", "Tan", "Time",
            "Timer", "TimeSerial", "TimeValue", "Trim", "TypeName", "UCase", "Val", "VarType", "Weekday", "WeekdayName", "Year"
        };
    }
}
