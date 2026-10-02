using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace SageTypes
{
    // StringS: texto com métodos, no lugar da classe StringS do Sage.xlam (mesma API).
    //
    //   Dim s As Sage.StringS: Set s = New Sage.StringS   ' só StringS seria o módulo VBA.Strings
    //   s = "  Olá\tmundo  "            ' membro padrão (Value); \n e \t viram quebra e tabulação
    //   Debug.Print s.Upper.Replace("MUNDO", "VBA")
    //
    // A interface é dual: o VBA chama pela vtable. Nunca reordene nem remova membros
    // de _StringS (o código VBA compilado contra a versão anterior quebraria);
    // acrescente novos sempre no fim, com o próximo DispId.
    [ComVisible(true), Guid("6E2B8F3A-41C7-4D5E-9A1B-3C8D7E6F5A21"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _StringS
    {
        // String (e não object): com object o .NET exporta Property Set, e "s = texto" não compila
        [DispId(0)] string Value { get; set; }
        [DispId(1)] StringS Upper();
        [DispId(2)] StringS Lower();
        [DispId(3)] StringS Proper();
        [DispId(4)] StringS Trim([Optional] object RemoveText);
        [DispId(5)] StringS Replace(object OldValue, object NewValue);
        [DispId(6)] StringS Right(object Length);
        [DispId(7)] StringS Left(object Length);
        // ParamArray do VBA é um array por referência; o "params" do C# sai por valor e o
        // VBA recusa ("tipo de automação não suportado"). Por isso: até 30 opcionais.
        [DispId(8)] StringS FString([Optional] object A1, [Optional] object A2, [Optional] object A3, [Optional] object A4, [Optional] object A5,
            [Optional] object A6, [Optional] object A7, [Optional] object A8, [Optional] object A9, [Optional] object A10,
            [Optional] object A11, [Optional] object A12, [Optional] object A13, [Optional] object A14, [Optional] object A15,
            [Optional] object A16, [Optional] object A17, [Optional] object A18, [Optional] object A19, [Optional] object A20,
            [Optional] object A21, [Optional] object A22, [Optional] object A23, [Optional] object A24, [Optional] object A25,
            [Optional] object A26, [Optional] object A27, [Optional] object A28, [Optional] object A29, [Optional] object A30);
        [DispId(9)] ListS Split(object Delimiter);
        [DispId(10)] StringS Reverse();
        [DispId(11)] int Length();
        [DispId(12)] bool IsNone();
        [DispId(13)] StringS FixUtf8();
        [DispId(14)] StringS Join([Optional] object A1, [Optional] object A2, [Optional] object A3, [Optional] object A4, [Optional] object A5,
            [Optional] object A6, [Optional] object A7, [Optional] object A8, [Optional] object A9, [Optional] object A10,
            [Optional] object A11, [Optional] object A12, [Optional] object A13, [Optional] object A14, [Optional] object A15,
            [Optional] object A16, [Optional] object A17, [Optional] object A18, [Optional] object A19, [Optional] object A20,
            [Optional] object A21, [Optional] object A22, [Optional] object A23, [Optional] object A24, [Optional] object A25,
            [Optional] object A26, [Optional] object A27, [Optional] object A28, [Optional] object A29, [Optional] object A30);
        [DispId(15)] int Count(object ArgChar);
        [DispId(16)] StringS NormalizeAccent();
        [DispId(17)] StringS Mid(int Start, [Optional] object Length);
        [DispId(18)] bool IsTextual();
        [DispId(19)] bool IsDigit();
        [DispId(20)] bool IsAlpha();
        [DispId(21)] bool IsNumeric();
    }

    [ComVisible(true), Guid("9A4C1E7B-2D58-4F3A-B6C9-0E1F2A3B4C5D"), ProgId("Sage.StringS")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_StringS))]
    public sealed class StringS : _StringS
    {
        string text = "";

        public StringS() { }

        static StringS From(string value)
        {
            StringS s = new StringS();
            s.text = value;
            return s;
        }

        public override string ToString() { return text; }

        // ------------------------------------------------------------------
        // Valor
        // ------------------------------------------------------------------

        // Atribuir: "\n" vira quebra de linha (vbNewLine), "\t" vira espaços até a
        // próxima coluna múltipla de 4, e o resultado perde espaços e quebras das pontas.
        public string Value
        {
            get { return text; }
            set { text = TrimEnter(ExpandTabs((value ?? "").Replace("\\n", "\r\n"))); }
        }

        static string ExpandTabs(string value)
        {
            if (value.IndexOf("\\t", StringComparison.Ordinal) < 0) return value;
            string[] lines = value.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            for (int n = 0; n < lines.Length; n++)
            {
                string[] parts = lines[n].Split(new string[] { "\\t" }, StringSplitOptions.None);
                if (parts.Length == 1) continue;
                StringBuilder sb = new StringBuilder();
                for (int t = 0; t < parts.Length - 1; t++)
                    sb.Append(parts[t]).Append(' ', 4 - parts[t].Length % 4);
                lines[n] = sb.Append(parts[parts.Length - 1]).ToString();
            }
            return string.Join("\r\n", lines);
        }

        // ------------------------------------------------------------------
        // Texto
        // ------------------------------------------------------------------

        public StringS Upper() { return From(text.ToUpper(CultureInfo.CurrentCulture)); }
        public StringS Lower() { return From(text.ToLower(CultureInfo.CurrentCulture)); }

        // Como PROPER do Excel: maiúscula depois de qualquer caractere que não é letra
        public StringS Proper()
        {
            StringBuilder sb = new StringBuilder(text.Length);
            bool afterLetter = false;
            foreach (char c in text)
            {
                sb.Append(afterLetter ? char.ToLower(c, CultureInfo.CurrentCulture) : char.ToUpper(c, CultureInfo.CurrentCulture));
                afterLetter = char.IsLetter(c);
            }
            return From(sb.ToString());
        }

        // Sem argumento: espaços e quebras de linha das pontas. Com RemoveText: tira
        // também esse texto das pontas, repetidamente ("--a--".Trim("-") = "a").
        public StringS Trim(object RemoveText)
        {
            string result = TrimEnter(text);
            string remove = Escapes(Optional(RemoveText, ""));
            if (remove.Length > 0)
            {
                bool changed = true;
                while (changed && result.Length > 0)
                {
                    changed = false;
                    if (result.StartsWith(remove, StringComparison.Ordinal)) { result = result.Substring(remove.Length); changed = true; }
                    if (result.EndsWith(remove, StringComparison.Ordinal)) { result = result.Substring(0, result.Length - remove.Length); changed = true; }
                    result = TrimEnter(result);
                }
            }
            return From(result);
        }

        // Texto por texto, ou listas (arrays) de antigos e novos, aos pares
        public StringS Replace(object OldValue, object NewValue)
        {
            string result = text;
            List<object> olds = Flatten(OldValue), news = Flatten(NewValue);
            if (olds != null && news != null)
            {
                for (int i = 0; i < Math.Min(olds.Count, news.Count); i++)
                    result = ReplaceText(result, Escapes(Text(olds[i])), Escapes(Text(news[i])));
            }
            else if (olds == null && news == null)
                result = ReplaceText(result, Escapes(Text(OldValue)), Escapes(Text(NewValue)));
            return From(result);
        }

        static string ReplaceText(string value, string old, string replacement)
        {
            return old.Length == 0 ? value : value.Replace(old, replacement);
        }

        // Positivo: os últimos N caracteres. Negativo: tudo menos os primeiros N.
        public StringS Right(object Length)
        {
            int n;
            if (!Number(Length, out n)) return From(text);
            if (n > 0) return From(n >= text.Length ? text : text.Substring(text.Length - n));
            if (n < 0) return From(text.Length + n <= 0 ? "" : text.Substring(0, text.Length + n));
            return From("");
        }

        // Positivo: os primeiros N caracteres. Negativo: tudo menos os últimos N.
        public StringS Left(object Length)
        {
            int n;
            if (!Number(Length, out n)) return From(text);
            if (n > 0) return From(n >= text.Length ? text : text.Substring(0, n));
            if (n < 0) return From(text.Length + n <= 0 ? "" : text.Substring(-n));
            return From("");
        }

        // Junta os argumentos (não usa o próprio texto)
        public StringS FString(object A1, object A2, object A3, object A4, object A5, object A6, object A7, object A8, object A9, object A10,
            object A11, object A12, object A13, object A14, object A15, object A16, object A17, object A18, object A19, object A20,
            object A21, object A22, object A23, object A24, object A25, object A26, object A27, object A28, object A29, object A30)
        {
            object[] args = { A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, A12, A13, A14, A15, A16, A17, A18, A19, A20,
                A21, A22, A23, A24, A25, A26, A27, A28, A29, A30 };
            StringBuilder sb = new StringBuilder();
            foreach (object arg in Items(args)) sb.Append(Text(arg));
            return From(sb.ToString());
        }

        // ListS, como no Sage.xlam: .Split(".")(-1) é o último pedaço
        public ListS Split(object Delimiter)
        {
            string delimiter = Text(Delimiter);
            string[] parts = delimiter.Length == 0 ? new string[] { text } : text.Split(new string[] { delimiter }, StringSplitOptions.None);
            return ListS.From(parts);
        }

        // Texto sem as conversões de Value (\n, \t, corte das pontas): ListS.Join
        internal void SetRaw(string value) { text = value ?? ""; }

        public StringS Reverse()
        {
            char[] chars = text.ToCharArray();
            Array.Reverse(chars);
            return From(new string(chars));
        }

        public int Length() { return text.Length; }

        public bool IsNone() { return text.Length == 0 || text == "\0"; }

        // Texto UTF-8 lido como ISO-8859-1 ("Ã§" -> "ç")
        public StringS FixUtf8()
        {
            string result = text;
            for (int i = 0; i < Utf8Wrong.Length; i++) result = result.Replace(Utf8Wrong[i], Utf8Right[i]);
            return From(result);
        }

        // O próprio texto é o separador: StringS(", ").Join("a", "b") = "a, b".
        // Um único array também serve: .Join(lista).
        public StringS Join(object A1, object A2, object A3, object A4, object A5, object A6, object A7, object A8, object A9, object A10,
            object A11, object A12, object A13, object A14, object A15, object A16, object A17, object A18, object A19, object A20,
            object A21, object A22, object A23, object A24, object A25, object A26, object A27, object A28, object A29, object A30)
        {
            object[] args = { A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, A12, A13, A14, A15, A16, A17, A18, A19, A20,
                A21, A22, A23, A24, A25, A26, A27, A28, A29, A30 };
            List<string> parts = new List<string>();
            foreach (object arg in Items(args)) parts.Add(Text(arg));
            return From(string.Join(text, parts.ToArray()));
        }

        // Quantas vezes o texto aparece (sem sobreposição)
        public int Count(object ArgChar)
        {
            string find = Text(ArgChar);
            if (find.Length == 0) return 0;
            int count = 0;
            for (int i = text.IndexOf(find, StringComparison.Ordinal); i >= 0; i = text.IndexOf(find, i + find.Length, StringComparison.Ordinal))
                count++;
            return count;
        }

        public StringS NormalizeAccent() { return From(RemoveAccents(text)); }

        // Como Mid do VBA (início em 1). Sem Length (ou "End"/"Max"): até o fim.
        public StringS Mid(int Start, object Length)
        {
            if (Start < 1) throw new COMException("Start deve ser maior ou igual a 1.", unchecked((int)0x800A0005));
            if (Start > text.Length) return From("");
            int n;
            string rest = text.Substring(Start - 1);
            if (!Number(Length, out n)) return From(rest);
            if (n < 0) throw new COMException("Length não pode ser negativo.", unchecked((int)0x800A0005));
            return From(n >= rest.Length ? rest : rest.Substring(0, n));
        }

        // Tem alguma letra (A-Z, com ou sem acento)
        public bool IsTextual()
        {
            foreach (char c in RemoveAccents(text).ToUpperInvariant())
                if (c >= 'A' && c <= 'Z') return true;
            return false;
        }

        // Como IsNumeric do VBA: é um número no formato local ("1,5", "1.000", "-2")
        public bool IsDigit()
        {
            double d;
            string s = text.Trim();
            if (s.Length == 0) return false;
            if (s.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
            {
                long h;
                return long.TryParse(s.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out h);
            }
            return double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d);
        }

        // Só letras (A-Z, com ou sem acento). Texto vazio: verdadeiro, como no Sage.xlam.
        public bool IsAlpha()
        {
            foreach (char c in RemoveAccents(text).ToUpperInvariant())
                if (c < 'A' || c > 'Z') return false;
            return true;
        }

        // Só dígitos 0-9. Texto vazio: verdadeiro, como no Sage.xlam.
        public bool IsNumeric()
        {
            foreach (char c in text)
                if (c < '0' || c > '9') return false;
            return true;
        }

        // ------------------------------------------------------------------
        // Conversões (valores vindos do VBA)
        // ------------------------------------------------------------------

        static string Text(object value) { return Interop.Text(value); }

        static string Optional(object value, string fallback)
        {
            return value == null || value is System.Reflection.Missing ? fallback : Text(value);
        }

        static string Escapes(string value)
        {
            return value.Replace("\\n", "\r\n").Replace("\\t", "\t");
        }

        static bool Number(object value, out int n)
        {
            n = 0;
            if (value == null || value is System.Reflection.Missing || value is DBNull) return false;
            if (value is string)
            {
                double d;
                if (!double.TryParse((string)value, NumberStyles.Any, CultureInfo.CurrentCulture, out d)) return false;
                n = (int)Math.Round(d, MidpointRounding.ToEven);
                return true;
            }
            try { n = Convert.ToInt32(value, CultureInfo.CurrentCulture); return true; }
            catch (Exception) { return false; }
        }

        // Argumentos opcionais passados (os omitidos chegam como Missing); um único
        // array ou ListS é expandido
        static IEnumerable Items(object[] args)
        {
            List<object> given = new List<object>();
            foreach (object arg in args)
                if (!(arg is System.Reflection.Missing)) given.Add(arg);
            if (given.Count == 1 && (given[0] is Array || given[0] is ListS)) return Flatten(given[0]);
            return given;
        }

        // Array do VBA (todas as células, mesmo 2D, como no Sage.xlam) ou ListS -> lista;
        // outro valor: null
        static List<object> Flatten(object value)
        {
            Array array = value as Array;
            if (array != null) return array.Cast<object>().ToList();
            ListS list = value as ListS;
            return list != null ? new List<object>(list.Items) : null;
        }

        static string TrimEnter(string value)
        {
            return value.Trim(' ', '\r', '\n');
        }

        const string Accented = "áàâãäéèêëíìîïóòôõöúùûüçñÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇÑ";
        const string Plain = "aaaaaeeeeiiiiooooouuuucnAAAAAEEEEIIIIOOOOOUUUUCN";

        static string RemoveAccents(string value)
        {
            StringBuilder sb = new StringBuilder(value.Length);
            foreach (char c in value)
            {
                int i = Accented.IndexOf(c);
                sb.Append(i >= 0 ? Plain[i] : c);
            }
            return sb.ToString();
        }

        static readonly string[] Utf8Wrong =
        {
            "Ã¡", "Ã¢", "Ã£", "Ã§", "Ã¨", "Ã©", "Ãª", "Ã«", "Ã­", "Ã³", "Ã´", "Ãµ", "Ã¶", "Ã¹", "Ãº", "Ã¼", "Ã±",
            "Ã\u0081", "Ã‚", "Ãƒ", "Ã‡", "Ã€", "Ã‰", "ÃŠ", "Ã‹", "Ã\u008D", "Ã“", "Ã”", "Ã•", "Ã–", "Ã™", "Ãš", "Ãœ", "Ã‘",
            "â€œ", "â€\u009D", "â€˜", "â€™", "â€“", "â€”",
            "Â¡", "Â¢", "Â£", "Â¤", "Â¥", "Â¦", "Â§", "Â¨", "Â©", "Âª", "Â«", "Â¬", "Â­", "Â®", "Â¯", "Â°", "Â±", "Â²", "Â³", "Â´",
            "Âµ", "Â¶", "Â·", "Â¸", "Â¹", "Âº", "Â»", "Â¼", "Â½", "Â¾", "Â¿", "Ã·", "Ã¸",
        };

        static readonly string[] Utf8Right =
        {
            "á", "â", "ã", "ç", "è", "é", "ê", "ë", "í", "ó", "ô", "õ", "ö", "ù", "ú", "ü", "ñ",
            "Á", "Â", "Ã", "Ç", "À", "É", "Ê", "Ë", "Í", "Ó", "Ô", "Õ", "Ö", "Ù", "Ú", "Ü", "Ñ",
            "“", "”", "‘", "’", "–", "—",
            "¡", "¢", "£", "¤", "¥", "¦", "§", "¨", "©", "ª", "«", "¬", "­", "®", "¯", "°", "±", "²", "³", "´",
            "µ", "¶", "·", "¸", "¹", "º", "»", "¼", "½", "¾", "¿", "÷", "ø",
        };
    }
}
