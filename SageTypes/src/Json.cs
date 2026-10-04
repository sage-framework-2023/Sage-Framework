using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace SageTypes
{
    // Json: o módulo json do Python. Global, como no Python (sem Dim):
    //
    //   Dim d As Sage.DictionaryS
    //   Set d = Json.Loads("{""nome"": ""Ana"", ""itens"": [1, 2.5, null]}")
    //   Debug.Print d("itens")(1)                     ' 2.5
    //   Debug.Print Json.Dumps(d, Indent:=2)
    //
    // Objeto -> DictionaryS, lista -> ListS, texto -> StringS, número -> Long/LongLong/Double,
    // true/false -> Boolean, null -> Empty (None). Na volta (Dumps), também Scripting.Dictionary,
    // arrays (2D: lista de linhas), Collection, Date/DateTimeS (texto ISO) e DataFrame (lista de linhas).
    //
    // Interface dual: acrescente membros só no fim, com o próximo DispId (ver StringS).
    [ComVisible(true), Guid("F9E2F586-295B-4A25-86CB-745578464E70"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _Json
    {
        [DispId(1)] object Loads(object Text);
        [DispId(2)] StringS Dumps(object Value, [Optional] object Indent, [Optional] object SortKeys, [Optional] object EnsureAscii);
        [DispId(3)] object Load(string Path, [Optional] object Encoding);
        [DispId(4)] void Dump(object Value, string Path, [Optional] object Indent, [Optional] object SortKeys, [Optional] object EnsureAscii);
    }

    [ComVisible(true), Guid("F255D235-A52E-4362-9484-F344705B9DE7"), ProgId("Sage.Json")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_Json))]
    public sealed class Json : _Json
    {
        public Json() { }

        public object Loads(object Text) { return Interop.Wrap(Parse(Interop.Text(Text))); }

        public StringS Dumps(object Value, object Indent, object SortKeys, object EnsureAscii)
        {
            return (StringS)Interop.Wrap(Serialize(Value, Indent, SortKeys, EnsureAscii));
        }

        // Arquivo em UTF-8 (padrão), com ou sem BOM
        public object Load(string Path, object Encoding)
        {
            Encoding encoding = Interop.IsMissing(Encoding) ? new UTF8Encoding(false) : System.Text.Encoding.GetEncoding(Interop.Text(Encoding));
            return Loads(File.ReadAllText(Path, encoding));
        }

        // Grava em UTF-8 sem BOM
        public void Dump(object Value, string Path, object Indent, object SortKeys, object EnsureAscii)
        {
            File.WriteAllText(Path, Serialize(Value, Indent, SortKeys, EnsureAscii), new UTF8Encoding(false));
        }

        // ------------------------------------------------------------------
        // Leitura
        // ------------------------------------------------------------------

        internal static object Parse(string text)
        {
            Parser p = new Parser(text ?? "");
            p.SkipSpace();
            object result = p.Value(0);
            p.SkipSpace();
            if (p.Position < p.Text.Length) throw p.Fail("Extra data");
            return result;
        }

        sealed class Parser
        {
            public readonly string Text;
            public int Position;

            public Parser(string text)
            {
                Text = text;
                if (Text.Length > 0 && Text[0] == '﻿') Position = 1; // BOM
            }

            // JSONDecodeError do Python: mensagem com linha e coluna
            public Exception Fail(string message)
            {
                int line = 1, column = 1;
                for (int i = 0; i < Position && i < Text.Length; i++)
                {
                    if (Text[i] == '\n') { line++; column = 1; }
                    else column++;
                }
                return Interop.Error(5, "JSONDecodeError: " + message + ": line " + line + " column " + column + " (char " + Position + ")");
            }

            public void SkipSpace()
            {
                while (Position < Text.Length && (Text[Position] == ' ' || Text[Position] == '\t' || Text[Position] == '\n' || Text[Position] == '\r'))
                    Position++;
            }

            public object Value(int depth)
            {
                if (depth > 1000) throw Fail("Too deeply nested");
                if (Position >= Text.Length) throw Fail("Expecting value");
                char c = Text[Position];
                switch (c)
                {
                    case '{': return Object(depth);
                    case '[': return Array(depth);
                    case '"': return String();
                }
                if (Word("true")) return true;
                if (Word("false")) return false;
                if (Word("null")) return null;
                if (Word("NaN")) return double.NaN;
                if (Word("Infinity")) return double.PositiveInfinity;
                if (Word("-Infinity")) return double.NegativeInfinity;
                if (c == '-' || (c >= '0' && c <= '9')) return Number();
                throw Fail("Expecting value");
            }

            bool Word(string word)
            {
                if (string.CompareOrdinal(Text, Position, word, 0, word.Length) != 0) return false;
                Position += word.Length;
                return true;
            }

            DictionaryS Object(int depth)
            {
                DictionaryS result = new DictionaryS();
                Position++; // {
                SkipSpace();
                if (Position < Text.Length && Text[Position] == '}') { Position++; return result; }
                while (true)
                {
                    SkipSpace();
                    if (Position >= Text.Length || Text[Position] != '"') throw Fail("Expecting property name enclosed in double quotes");
                    string key = String();
                    SkipSpace();
                    if (Position >= Text.Length || Text[Position] != ':') throw Fail("Expecting ':' delimiter");
                    Position++;
                    SkipSpace();
                    result.LetValue(key, Value(depth + 1)); // chave repetida: fica a última, como no Python
                    SkipSpace();
                    if (Position < Text.Length && Text[Position] == ',') { Position++; continue; }
                    if (Position < Text.Length && Text[Position] == '}') { Position++; return result; }
                    throw Fail("Expecting ',' delimiter");
                }
            }

            ListS Array(int depth)
            {
                List<object> items = new List<object>();
                Position++; // [
                SkipSpace();
                if (Position < Text.Length && Text[Position] == ']') { Position++; return ListS.From(items); }
                while (true)
                {
                    SkipSpace();
                    items.Add(Value(depth + 1));
                    SkipSpace();
                    if (Position < Text.Length && Text[Position] == ',') { Position++; continue; }
                    if (Position < Text.Length && Text[Position] == ']') { Position++; return ListS.From(items); }
                    throw Fail("Expecting ',' delimiter");
                }
            }

            string String()
            {
                int start = Position;
                Position++; // "
                StringBuilder sb = new StringBuilder();
                while (true)
                {
                    if (Position >= Text.Length) { Position = start; throw Fail("Unterminated string starting at"); }
                    char c = Text[Position++];
                    if (c == '"') return sb.ToString();
                    if (c < ' ') { Position--; throw Fail("Invalid control character at"); }
                    if (c != '\\') { sb.Append(c); continue; }
                    if (Position >= Text.Length) { Position = start; throw Fail("Unterminated string starting at"); }
                    char e = Text[Position++];
                    switch (e)
                    {
                        case '"': sb.Append('"'); break;
                        case '\\': sb.Append('\\'); break;
                        case '/': sb.Append('/'); break;
                        case 'b': sb.Append('\b'); break;
                        case 'f': sb.Append('\f'); break;
                        case 'n': sb.Append('\n'); break;
                        case 'r': sb.Append('\r'); break;
                        case 't': sb.Append('\t'); break;
                        case 'u':
                            int code;
                            if (Position + 4 > Text.Length || !int.TryParse(Text.Substring(Position, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
                            {
                                Position -= 2;
                                throw Fail("Invalid \\uXXXX escape");
                            }
                            sb.Append((char)code);
                            Position += 4;
                            break;
                        default:
                            Position -= 2;
                            throw Fail("Invalid \\escape");
                    }
                }
            }

            object Number()
            {
                int start = Position;
                if (Text[Position] == '-') Position++;
                if (Position >= Text.Length || !char.IsDigit(Text[Position])) { Position = start; throw Fail("Expecting value"); }
                if (Text[Position] == '0') Position++;
                else while (Position < Text.Length && char.IsDigit(Text[Position])) Position++;
                bool integral = true;
                if (Position < Text.Length && Text[Position] == '.' && Position + 1 < Text.Length && char.IsDigit(Text[Position + 1]))
                {
                    integral = false;
                    Position++;
                    while (Position < Text.Length && char.IsDigit(Text[Position])) Position++;
                }
                if (Position < Text.Length && (Text[Position] == 'e' || Text[Position] == 'E'))
                {
                    int mark = Position++;
                    if (Position < Text.Length && (Text[Position] == '+' || Text[Position] == '-')) Position++;
                    if (Position < Text.Length && char.IsDigit(Text[Position]))
                    {
                        integral = false;
                        while (Position < Text.Length && char.IsDigit(Text[Position])) Position++;
                    }
                    else Position = mark;
                }
                string number = Text.Substring(start, Position - start);
                if (integral)
                {
                    // Long se couber; senão LongLong; senão Double
                    int i;
                    if (int.TryParse(number, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out i)) return i;
                    long l;
                    if (long.TryParse(number, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out l)) return l;
                }
                return double.Parse(number, NumberStyles.Float, CultureInfo.InvariantCulture);
            }
        }

        // ------------------------------------------------------------------
        // Escrita
        // ------------------------------------------------------------------

        internal static string Serialize(object value, object indent, object sortKeys, object ensureAscii)
        {
            Writer w = new Writer();
            if (!Interop.IsMissing(indent) && indent != null)
            {
                object i = Interop.Unwrap(indent);
                w.Indent = i is string ? (string)i : new string(' ', Interop.Integer(i));
            }
            w.SortKeys = !Interop.IsMissing(sortKeys) && Convert.ToBoolean(sortKeys, CultureInfo.InvariantCulture);
            w.EnsureAscii = !Interop.IsMissing(ensureAscii) && Convert.ToBoolean(ensureAscii, CultureInfo.InvariantCulture);
            w.Write(value, 0);
            return w.Output.ToString();
        }

        sealed class Writer
        {
            public readonly StringBuilder Output = new StringBuilder();
            public string Indent;          // null: tudo numa linha
            public bool SortKeys, EnsureAscii;
            readonly HashSet<object> open = new HashSet<object>(ReferenceComparer.Instance);

            public void Write(object value, int depth)
            {
                if (depth > 1000) throw Interop.Error(5, "ValueError: estrutura aninhada demais.");
                if (value == null || value is DBNull || Interop.IsMissing(value) || value == KeyComparer.Empty) { Output.Append("null"); return; }
                StringS s = value as StringS;
                if (s != null) { Quote(s.Value); return; }
                string text = value as string;
                if (text != null) { Quote(text); return; }
                if (value is bool) { Output.Append((bool)value ? "true" : "false"); return; }
                if (value is double || value is float) { Output.Append(Float(Convert.ToDouble(value, CultureInfo.InvariantCulture))); return; }
                if (value is decimal || (value.GetType().IsPrimitive && !(value is char))) { Output.Append(((IFormattable)value).ToString(null, CultureInfo.InvariantCulture)); return; }
                if (value is DateTime) { Quote(DateTimeS.From((DateTime)value, null).Iso8601("T", "auto")); return; }
                DateTimeS date = value as DateTimeS;
                if (date != null) { Quote(date.Iso8601("T", "auto")); return; }

                if (!open.Add(value)) throw Interop.Error(5, "ValueError: Circular reference detected");
                try
                {
                    DictionaryS dict = value as DictionaryS;
                    if (dict != null) { WriteObject(dict.Pairs(), depth); return; }
                    ListS list = value as ListS;
                    if (list != null) { WriteArray(list.Items, depth); return; }
                    Array array = value as Array;
                    if (array != null) { WriteArray(Interop.Items(array), depth); return; }
                    DataFrame frame = value as DataFrame;
                    if (frame != null) { WriteArray(frame.Cast<object>(), depth); return; }
                    if (Marshal.IsComObject(value))
                    {
                        // Scripting.Dictionary (tem Keys(), reconhecido por isso: o nome do tipo
                        // visto pelo .NET é IDictionary) ou Collection; outros objetos não são JSON
                        IEnumerable<KeyValuePair<object, object>> pairs = null;
                        try { pairs = Interop.PairsOf(value); }
                        catch (COMException) { }
                        if (pairs != null) { WriteObject(pairs, depth); return; }
                        string type = Microsoft.VisualBasic.Information.TypeName(value);
                        if (type != "Range" && value is IEnumerable) { WriteArray(Interop.Items(value), depth); return; }
                    }
                }
                finally { open.Remove(value); }
                throw Interop.Error(13, "TypeError: Object of type " + Interop.TypeLabel(value) + " is not JSON serializable");
            }

            void WriteObject(IEnumerable<KeyValuePair<object, object>> pairs, int depth)
            {
                List<KeyValuePair<string, object>> items = new List<KeyValuePair<string, object>>();
                foreach (KeyValuePair<object, object> pair in pairs) items.Add(new KeyValuePair<string, object>(Key(pair.Key), pair.Value));
                if (SortKeys) items.Sort((a, b) => string.CompareOrdinal(a.Key, b.Key));
                if (items.Count == 0) { Output.Append("{}"); return; }
                Output.Append('{');
                for (int i = 0; i < items.Count; i++)
                {
                    if (i > 0) Output.Append(Indent == null ? ", " : ",");
                    NewLine(depth + 1);
                    Quote(items[i].Key);
                    Output.Append(": ");
                    Write(items[i].Value, depth + 1);
                }
                NewLine(depth);
                Output.Append('}');
            }

            void WriteArray(IEnumerable<object> items, int depth)
            {
                bool first = true;
                Output.Append('[');
                foreach (object item in items)
                {
                    if (!first) Output.Append(Indent == null ? ", " : ",");
                    NewLine(depth + 1);
                    first = false;
                    Write(item, depth + 1);
                }
                if (first) { Output.Append(']'); return; }
                NewLine(depth);
                Output.Append(']');
            }

            void NewLine(int depth)
            {
                if (Indent == null) return;
                Output.Append('\n');
                for (int i = 0; i < depth; i++) Output.Append(Indent);
            }

            // Chaves como no Python: texto; números e True/False/None viram texto
            static string Key(object key)
            {
                if (key == null || key == KeyComparer.Empty) return "null";
                if (key is string) return (string)key;
                if (key is bool) return (bool)key ? "true" : "false";
                if (key is double || key is float) return Float(Convert.ToDouble(key, CultureInfo.InvariantCulture));
                if (key is decimal || key.GetType().IsPrimitive) return ((IFormattable)key).ToString(null, CultureInfo.InvariantCulture);
                if (key is DateTime) return DateTimeS.From((DateTime)key, null).Iso8601("T", "auto");
                throw Interop.Error(13, "TypeError: keys must be str, int, float, bool or None, not " + Interop.TypeLabel(key));
            }

            // Como o Python: 1.0, 2.5, 1e+20, NaN, Infinity
            static string Float(double d)
            {
                if (double.IsNaN(d)) return "NaN";
                if (double.IsPositiveInfinity(d)) return "Infinity";
                if (double.IsNegativeInfinity(d)) return "-Infinity";
                string text = d.ToString("R", CultureInfo.InvariantCulture);
                if (text.IndexOfAny(new[] { '.', 'E', 'e' }) < 0) text += ".0";
                return text.Replace("E+", "e+").Replace("E-", "e-");
            }

            void Quote(string s)
            {
                Output.Append('"');
                foreach (char c in s)
                {
                    switch (c)
                    {
                        case '"': Output.Append("\\\""); break;
                        case '\\': Output.Append("\\\\"); break;
                        case '\n': Output.Append("\\n"); break;
                        case '\r': Output.Append("\\r"); break;
                        case '\t': Output.Append("\\t"); break;
                        case '\b': Output.Append("\\b"); break;
                        case '\f': Output.Append("\\f"); break;
                        default:
                            if (c < ' ' || (EnsureAscii && c > '~')) Output.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                            else Output.Append(c);
                            break;
                    }
                }
                Output.Append('"');
            }
        }

        // Referência (não igualdade): duas listas iguais não são um ciclo
        sealed class ReferenceComparer : IEqualityComparer<object>
        {
            public static readonly ReferenceComparer Instance = new ReferenceComparer();
            public new bool Equals(object a, object b) { return ReferenceEquals(a, b); }
            public int GetHashCode(object o) { return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(o); }
        }
    }
}
