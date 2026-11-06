using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace SageTypes
{
    // Console: escreve na aba Terminal da janela do Painel (Ctrl+J). Global (sem Dim):
    //
    //   Console.PrintF "Linha %d de %d: %s\n", i, total, nome
    //   Console.PrintF "Total: %.2f (%5.1f%%)\n", soma, taxa
    //
    // Como o printf do C, a quebra de linha só vem com o \n. Marcadores: %d %i %u (inteiro),
    // %f %e %g (decimal), %s (qualquer valor, como o CStr do VBA), %c (caractere ou código),
    // %x %X %o (hexadecimal, octal) e %%, com as opções %-5d, %05d, %+d, %.2f. Os números saem
    // sempre com ponto decimal, em qualquer idioma; %s segue o Windows, como o CStr.
    //
    // Acrescente membros só no fim, com o próximo DispId (ver StringS). Não é dual, ao contrário
    // das outras: uma dual responderia ao Dim o As Object com o IDispatch dela, que não sabe
    // juntar o ParamArray (ver ConsoleLateBound). Acrescente cada membro novo lá também.
    [ComVisible(true), Guid("AD3E917D-B6FF-4B86-85D9-89845062DB7D"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface _Console
    {
        // ParamArray de verdade: array por referência aqui e [VarArg] no .tlb (install.ps1)
        [DispId(1), VarArg] void PrintF(object Format,
            [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] Args);
    }

    // Só para chamadas por Object (ver Console.GetInterface); não é criada pelo VBA
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDispatch), TypeLibType(TypeLibTypeFlags.FHidden)]
    public sealed class ConsoleLateBound
    {
        readonly Console owner;

        internal ConsoleLateBound(Console owner) { this.owner = owner; }

        public void PrintF(object Format, params object[] Args) { owner.PrintF(Format, ref Args); }
    }

    // O último parâmetro do método é um ParamArray no VBA (cParamsOpt = -1, aplicado pelo install.ps1)
    [AttributeUsage(AttributeTargets.Method)]
    sealed class VarArgAttribute : Attribute { }

    [ComVisible(true), Guid("48E1D247-7261-4DBB-9B9A-FA601658FC2F"), ProgId("Sage.Console")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_Console))]
    public sealed class Console : _Console, ICustomQueryInterface
    {
        // O SageEditor deixa no AppDomain uma função que recebe o texto (TerminalWindow), como a
        // do df.Show: os dois rodam no processo do Excel, sem referência um ao outro. Sem o
        // SageEditor não há Terminal, e o texto é descartado: a macro segue normalmente.
        public const string EditorSlot = "Sage.ConsolePrint";

        static readonly Guid IDispatch = new Guid("00020400-0000-0000-C000-000000000046");
        readonly ConsoleLateBound lateBound;

        public Console() { lateBound = new ConsoleLateBound(this); }

        // Dim o As Object: pelo IDispatch, o VBA manda os argumentos soltos, e o .NET não os
        // junta no array por referência da interface. Nesse caminho responde uma classe com params.
        CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            if (iid != IDispatch) return CustomQueryInterfaceResult.NotHandled;
            ppv = Marshal.GetIDispatchForObject(lateBound);
            return CustomQueryInterfaceResult.Handled;
        }

        public void PrintF(object Format, ref object[] Args)
        {
            List<object> args = new List<object>();
            if (Args != null)
                foreach (object arg in Args)
                    if (!Interop.IsMissing(arg)) args.Add(DefaultValue(arg));
            string text = Render(Interop.Text(DefaultValue(Format)), args);
            Func<string, bool> print = AppDomain.CurrentDomain.GetData(EditorSlot) as Func<string, bool>;
            if (print != null) print(text);
        }

        // Num ParamArray o VBA não lê a propriedade padrão: Range("A1") chega como objeto, e
        // aqui vira o valor da célula. Objeto sem propriedade padrão fica como está.
        static object DefaultValue(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return value;
            try { return value.GetType().InvokeMember("", BindingFlags.GetProperty | BindingFlags.InvokeMethod, null, value, null); }
            catch (Exception) { return value; }
        }

        // ------------------------------------------------------------------
        // Formatação
        // ------------------------------------------------------------------

        sealed class Spec
        {
            public bool Left, Plus, Space, Zero, Alternate;
            public int Width;
            public int Precision = -1; // -1: não informada
            public char Type;
        }

        static string Render(string format, List<object> args)
        {
            StringBuilder sb = new StringBuilder();
            int next = 0, i = 0;
            while (i < format.Length)
            {
                char c = format[i];
                if (c == '\\' && i + 1 < format.Length)
                {
                    string escape = Escape(format[i + 1]);
                    if (escape == null) { sb.Append(c); i++; continue; } // C:\dados: fica como está
                    sb.Append(escape);
                    i += 2;
                    continue;
                }
                if (c != '%') { sb.Append(c); i++; continue; }
                int start = i++;
                Spec spec = Parse(format, ref i, start);
                if (spec.Type == '%') { sb.Append('%'); continue; }
                if (next >= args.Count) throw Interop.Error(450, "TypeError: not enough arguments for format string");
                sb.Append(Convert(spec, args[next++]));
            }
            if (next < args.Count) throw Interop.Error(450, "TypeError: not all arguments converted during string formatting");
            return sb.ToString();
        }

        static string Escape(char c)
        {
            switch (c)
            {
                case 'n': return "\n";
                case 't': return "\t";
                case 'r': return "\r";
                case '\\': return "\\";
                default: return null;
            }
        }

        // Depois do %: opções, largura, .precisão e o tipo (ignora os "l" e "h" do C: %ld)
        static Spec Parse(string format, ref int i, int start)
        {
            Spec spec = new Spec();
            for (; i < format.Length; i++)
            {
                char f = format[i];
                if (f == '-') spec.Left = true;
                else if (f == '+') spec.Plus = true;
                else if (f == ' ') spec.Space = true;
                else if (f == '0') spec.Zero = true;
                else if (f == '#') spec.Alternate = true;
                else break;
            }
            spec.Width = Digits(format, ref i);
            if (i < format.Length && format[i] == '.')
            {
                i++;
                spec.Precision = Digits(format, ref i);
            }
            while (i < format.Length && (format[i] == 'l' || format[i] == 'h')) i++;
            if (i >= format.Length) throw Interop.Error(5, "ValueError: incomplete format");
            spec.Type = format[i++];
            if ("diufFeEgGxXoscC%".IndexOf(spec.Type) < 0)
                throw Interop.Error(5, "ValueError: unsupported format character '" + spec.Type + "' at index " + (i - 1));
            return spec;
        }

        static int Digits(string format, ref int i)
        {
            int value = 0;
            while (i < format.Length && char.IsDigit(format[i]))
            {
                value = Math.Min(value * 10 + (format[i] - '0'), 4096);
                i++;
            }
            return value;
        }

        static string Convert(Spec spec, object value)
        {
            switch (spec.Type)
            {
                case 's':
                    string s = value is DateTime ? DateText((DateTime)value) : Interop.Text(value);
                    if (spec.Precision >= 0 && spec.Precision < s.Length) s = s.Substring(0, spec.Precision);
                    return Pad(spec, "", s, false);
                case 'c':
                case 'C':
                    return Pad(spec, "", Character(value), false);
                case 'd':
                case 'i':
                    long n = Integer(spec, value);
                    string digits = n == long.MinValue ? "9223372036854775808" : Math.Abs(n).ToString(CultureInfo.InvariantCulture);
                    return Whole(spec, Sign(spec, n < 0), digits);
                case 'u':
                    return Whole(spec, "", Unsigned(Integer(spec, value), value).ToString(CultureInfo.InvariantCulture));
                case 'x':
                case 'X':
                case 'o':
                    ulong bits = Unsigned(Integer(spec, value), value);
                    string text = System.Convert.ToString(unchecked((long)bits), spec.Type == 'o' ? 8 : 16);
                    string prefix = spec.Alternate && bits != 0 ? (spec.Type == 'o' ? "0" : "0x") : "";
                    if (spec.Type == 'X') { text = text.ToUpperInvariant(); prefix = prefix.ToUpperInvariant(); }
                    return Whole(spec, prefix, text);
                default:
                    return Real(spec, Number(spec, value));
            }
        }

        // Como o CStr: sem a hora quando ela é 00:00:00, e só a hora quando não há data (Time)
        static string DateText(DateTime d)
        {
            if (d.Year == 1899 && d.Month == 12 && d.Day == 30) return d.ToLongTimeString();
            return d.TimeOfDay == TimeSpan.Zero ? d.ToShortDateString() : d.ToShortDateString() + " " + d.ToLongTimeString();
        }

        static string Sign(Spec spec, bool negative)
        {
            return negative ? "-" : spec.Plus ? "+" : spec.Space ? " " : "";
        }

        // Inteiro: a precisão é o mínimo de dígitos e, como no C, cancela o preenchimento com zeros
        static string Whole(Spec spec, string prefix, string digits)
        {
            if (spec.Precision >= 0) digits = digits.PadLeft(spec.Precision, '0');
            return Pad(spec, prefix, digits, spec.Zero && spec.Precision < 0);
        }

        static string Pad(Spec spec, string prefix, string body, bool zeros)
        {
            int fill = spec.Width - prefix.Length - body.Length;
            if (fill <= 0) return prefix + body;
            if (spec.Left) return prefix + body + new string(' ', fill);
            if (zeros) return prefix + new string('0', fill) + body;
            return new string(' ', fill) + prefix + body;
        }

        static string Real(Spec spec, double value)
        {
            bool upper = char.IsUpper(spec.Type);
            if (double.IsNaN(value)) return Pad(spec, "", upper ? "NAN" : "nan", false);
            bool negative = value < 0 || (value == 0 && 1 / value < 0);
            string sign = Sign(spec, negative);
            double abs = Math.Abs(value);
            if (double.IsInfinity(value)) return Pad(spec, sign, upper ? "INF" : "inf", false);
            int precision = spec.Precision >= 0 ? spec.Precision : 6;
            string body;
            switch (char.ToLowerInvariant(spec.Type))
            {
                case 'f': body = Fixed(abs, precision); break;
                case 'e': body = Exponent(abs, precision); break;
                default: body = General(abs, precision, spec.Alternate); break;
            }
            return Pad(spec, sign, upper ? body.ToUpperInvariant() : body, spec.Zero);
        }

        static string Fixed(double value, int precision)
        {
            return value.ToString("F" + precision, CultureInfo.InvariantCulture);
        }

        // Com dois dígitos no expoente, como no C: 1.500000e+03
        static string Exponent(double value, int precision)
        {
            string mantissa = precision > 0 ? "0." + new string('0', precision) : "0";
            return value.ToString(mantissa + "e+00", CultureInfo.InvariantCulture);
        }

        // %g: o mais curto entre %f e %e, com a precisão em dígitos significativos
        static string General(double value, int precision, bool keepZeros)
        {
            if (value == 0) return "0";
            int digits = precision == 0 ? 1 : precision;
            string probe = value.ToString("e" + (digits - 1), CultureInfo.InvariantCulture);
            int exponent = int.Parse(probe.Substring(probe.IndexOf('e') + 1), CultureInfo.InvariantCulture);
            bool plain = exponent >= -4 && exponent < digits;
            string body = plain ? Fixed(value, digits - 1 - exponent) : Exponent(value, digits - 1);
            if (keepZeros) return body;
            int e = body.IndexOf('e');
            string mantissa = e >= 0 ? body.Substring(0, e) : body;
            if (mantissa.IndexOf('.') >= 0) mantissa = mantissa.TrimEnd('0').TrimEnd('.');
            return e >= 0 ? mantissa + body.Substring(e) : mantissa;
        }

        // %d de um número com fração corta, como no Python e no Fix do VBA. Empty vale 0, como no VBA.
        static long Integer(Spec spec, object value)
        {
            object v = Interop.Unwrap(value);
            if (v == null) return 0;
            if (v is bool) return (bool)v ? 1 : 0;
            RequireNumber(spec, v);
            try
            {
                if (v is decimal) return System.Convert.ToInt64(decimal.Truncate((decimal)v));
                if (v is double || v is float)
                {
                    double d = System.Convert.ToDouble(v, CultureInfo.InvariantCulture);
                    if (double.IsNaN(d) || double.IsInfinity(d)) throw new OverflowException();
                    return System.Convert.ToInt64(Math.Truncate(d));
                }
                return System.Convert.ToInt64(v, CultureInfo.InvariantCulture);
            }
            catch (OverflowException)
            {
                throw Interop.Error(6, "OverflowError: value out of range for %" + spec.Type);
            }
        }

        // Negativo em %u %x %o: o complemento de dois no tamanho do tipo (Long: 32 bits)
        static ulong Unsigned(long n, object value)
        {
            object v = Interop.Unwrap(value);
            if (n < 0 && (v is int || v is short || v is sbyte)) return unchecked((uint)(int)n);
            return unchecked((ulong)n);
        }

        static double Number(Spec spec, object value)
        {
            object v = Interop.Unwrap(value);
            if (v == null) return 0;
            if (v is bool) return (bool)v ? 1 : 0;
            RequireNumber(spec, v);
            return System.Convert.ToDouble(v, CultureInfo.InvariantCulture);
        }

        static void RequireNumber(Spec spec, object v)
        {
            if (!PyOrder.IsNumber(v))
                throw Interop.Error(13, "TypeError: %" + spec.Type + " format: a number is required, not " + Interop.TypeLabel(v));
        }

        static string Character(object value)
        {
            object v = Interop.Unwrap(value);
            string s = v as string;
            if (s != null)
            {
                bool one = s.Length == 1 || (s.Length == 2 && char.IsSurrogatePair(s, 0));
                if (one) return s;
            }
            else if (v != null && !(v is bool) && PyOrder.IsNumber(v))
            {
                double code = System.Convert.ToDouble(v, CultureInfo.InvariantCulture);
                bool valid = code >= 0 && code <= 0x10FFFF && code == Math.Floor(code) && !(code >= 0xD800 && code <= 0xDFFF);
                if (valid) return char.ConvertFromUtf32((int)code);
            }
            throw Interop.Error(13, "TypeError: %c requires a character or a character code");
        }
    }
}
