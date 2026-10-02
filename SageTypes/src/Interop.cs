using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace SageTypes
{
    // Conversões entre os valores que chegam do VBA e os tipos do Sage
    static class Interop
    {
        // Argumento opcional omitido no VBA
        public static bool IsMissing(object value)
        {
            return value is System.Reflection.Missing;
        }

        // Erro com o número do VBA (Err.Number) e a mensagem (Err.Description)
        public static COMException Error(int vbaNumber, string message)
        {
            return new COMException(message, unchecked((int)0x800A0000) | vbaNumber);
        }

        // StringS vira o texto dele (para comparar e guardar como chave)
        public static object Unwrap(object value)
        {
            StringS s = value as StringS;
            return s != null ? s.Value : value;
        }

        // ------------------------------------------------------------------
        // Valores dentro de ListS e DictionaryS
        // ------------------------------------------------------------------

        // Ao guardar: array vira ListS (2D: lista de linhas), com os de dentro também,
        // para "lista(2).Append x" alterar a lista guardada, como no Python; StringS vira
        // texto (é imutável, como str). O resto (números, datas, objetos) fica como está.
        public static object Store(object value)
        {
            StringS s = value as StringS;
            if (s != null) return s.Value;
            Array array = value as Array;
            if (array != null) return ListS.From(Items(array));
            return value;
        }

        // Ao ler: texto volta como StringS, para "lista(0).Upper" funcionar; o VBA usa o
        // Value dele onde espera texto (Debug.Print, x = lista(0), comparações).
        public static object Wrap(object value)
        {
            string s = value as string;
            if (s == null) return value;
            StringS result = new StringS();
            result.SetRaw(s);
            return result;
        }

        // Para devolver ao VBA "puro" (Value sem índice, ToArray): ListS vira array do
        // VBA, inclusive os de dentro; texto fica String
        public static object Plain(object value)
        {
            ListS list = value as ListS;
            if (list == null) return value;
            object[] result = new object[list.Items.Count];
            for (int i = 0; i < result.Length; i++) result[i] = Plain(list.Items[i]);
            return result;
        }

        // CStr do VBA; Empty, Null e Missing viram ""
        public static string Text(object value)
        {
            if (value == null || value is DBNull || IsMissing(value)) return "";
            string s = value as string;
            if (s != null) return s;
            StringS str = value as StringS;
            if (str != null) return str.Value;
            if (value is bool)
            {
                // CStr(True) do VBA segue o idioma do Office; aqui, o do Windows
                bool pt = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "pt";
                return (bool)value ? (pt ? "Verdadeiro" : "True") : (pt ? "Falso" : "False");
            }
            return Convert.ToString(value, CultureInfo.CurrentCulture);
        }

        // Índice ou quantidade: número inteiro (Integer, Long, Double sem fração...)
        public static int Integer(object value)
        {
            object v = Unwrap(value);
            if (v is string)
            {
                double d;
                if (double.TryParse((string)v, NumberStyles.Any, CultureInfo.CurrentCulture, out d)) v = d;
            }
            if (!PyOrder.IsNumber(v)) throw Error(13, "TypeError: indices must be integers, not " + TypeLabel(v));
            return Convert.ToInt32(v, CultureInfo.InvariantCulture);
        }

        public static string TypeLabel(object value)
        {
            if (value == null) return "Empty";
            if (Marshal.IsComObject(value)) return Microsoft.VisualBasic.Information.TypeName(value);
            return value.GetType().Name;
        }

        // Elementos de algo "iterável", como no Python: array (2D: as linhas, como arrays),
        // ListS, DictionaryS (as chaves), texto (os caracteres) ou objeto COM com For Each
        // (Collection, Range...)
        public static IEnumerable<object> Items(object source)
        {
            if (source == null || IsMissing(source)) throw Error(13, "TypeError: 'Empty' object is not iterable");
            ListS list = source as ListS;
            if (list != null) return new List<object>(list.Items);
            DictionaryS dict = source as DictionaryS;
            if (dict != null) return new List<object>((object[])dict.KeysArray());
            object text = Unwrap(source);
            if (text is string) return ((string)text).Select(c => (object)c.ToString()).ToList();
            Array array = source as Array;
            if (array != null && array.Rank == 2) return Rows(array);
            System.Collections.IEnumerable enumerable = source as System.Collections.IEnumerable;
            if (enumerable != null) return enumerable.Cast<object>().ToList();
            throw Error(13, "TypeError: '" + TypeLabel(source) + "' object is not iterable");
        }

        // ------------------------------------------------------------------
        // repr() do Python: {'a': 1, 'b': [1, 2.5], 'c': None}
        // ------------------------------------------------------------------

        public static string Repr(object value)
        {
            StringBuilder sb = new StringBuilder();
            Repr(value, sb, new HashSet<object>());
            return sb.ToString();
        }

        static void Repr(object value, StringBuilder sb, HashSet<object> open)
        {
            if (value == null || value == KeyComparer.Empty) { sb.Append("None"); return; }
            if (value is DBNull) { sb.Append("Null"); return; }
            string s = value as string;
            if (s != null) { Quote(s, sb); return; }
            StringS str = value as StringS;
            if (str != null) { Quote(str.Value, sb); return; }
            if (value is bool) { sb.Append((bool)value ? "True" : "False"); return; }
            if (value is DateTime)
            {
                DateTime d = (DateTime)value;
                sb.Append('#').Append(d.ToString(d.TimeOfDay == TimeSpan.Zero ? "yyyy-MM-dd" : "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)).Append('#');
                return;
            }
            if (value is double) { sb.Append(((double)value).ToString("R", CultureInfo.InvariantCulture)); return; }
            if (value is float) { sb.Append(((float)value).ToString("R", CultureInfo.InvariantCulture)); return; }
            if (value.GetType().IsPrimitive || value is decimal) { sb.Append(((IFormattable)value).ToString(null, CultureInfo.InvariantCulture)); return; }

            DictionaryS dict = value as DictionaryS;
            Array array = value as Array;
            ListS list = value as ListS;
            if (dict != null || array != null || list != null)
            {
                // Estrutura que contém a si mesma, como o Python: {...} / [...]
                if (!open.Add(value)) { sb.Append(dict != null ? "{...}" : "[...]"); return; }
                bool first = true;
                if (list != null)
                {
                    // Tupla como no Python: (1, 2) e (1,)
                    sb.Append(list.IsTuple ? '(' : '[');
                    foreach (object item in list.Items)
                    {
                        if (!first) sb.Append(", ");
                        first = false;
                        Repr(item, sb, open);
                    }
                    if (list.IsTuple && list.Items.Count == 1) sb.Append(',');
                    sb.Append(list.IsTuple ? ')' : ']');
                }
                else if (dict != null)
                {
                    sb.Append('{');
                    foreach (KeyValuePair<object, object> pair in dict.Pairs())
                    {
                        if (!first) sb.Append(", ");
                        first = false;
                        Repr(pair.Key, sb, open);
                        sb.Append(": ");
                        Repr(pair.Value, sb, open);
                    }
                    sb.Append('}');
                }
                else
                {
                    sb.Append('[');
                    foreach (object item in array)
                    {
                        if (!first) sb.Append(", ");
                        first = false;
                        Repr(item, sb, open);
                    }
                    sb.Append(']');
                }
                open.Remove(value);
                return;
            }

            // Outros objetos (Range, Collection...): <TypeName>
            sb.Append('<').Append(Marshal.IsComObject(value) ? Microsoft.VisualBasic.Information.TypeName(value) : value.GetType().Name).Append('>');
        }

        static void Quote(string s, StringBuilder sb)
        {
            sb.Append('\'');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '\'': sb.Append("\\'"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default: sb.Append(c); break;
                }
            }
            sb.Append('\'');
        }

        // ------------------------------------------------------------------
        // Pares chave/valor
        // ------------------------------------------------------------------

        // DictionaryS, Scripting.Dictionary, Array/ListS (k1, v1, k2, v2...) ou de pares
        // (Array(k1, v1), Array(k2, v2)), como dict() do Python
        public static IEnumerable<KeyValuePair<object, object>> PairsOf(object source)
        {
            DictionaryS dict = source as DictionaryS;
            if (dict != null) return new List<KeyValuePair<object, object>>(dict.Pairs());

            List<KeyValuePair<object, object>> result = new List<KeyValuePair<object, object>>();
            ListS sourceList = source as ListS;
            Array array = source as Array;
            if (array != null || sourceList != null)
            {
                List<object> items = sourceList != null ? new List<object>(sourceList.Items) : array.Cast<object>().ToList();
                bool pairs = items.Count > 0;
                foreach (object item in items)
                {
                    if (PairCount(item) != 2) { pairs = false; break; }
                }
                if (pairs)
                {
                    foreach (object item in items)
                    {
                        List<object> p = Items(item).ToList();
                        result.Add(new KeyValuePair<object, object>(p[0], p[1]));
                    }
                    return result;
                }
                if (items.Count % 2 != 0) throw Error(5, "Array com número ímpar de elementos: use Array(chave1, valor1, chave2, valor2...)");
                for (int i = 0; i < items.Count; i += 2)
                    result.Add(new KeyValuePair<object, object>(items[i], items[i + 1]));
                return result;
            }

            // Scripting.Dictionary ou outro objeto COM com Keys() e Item(chave)
            if (source != null && Marshal.IsComObject(source))
            {
                dynamic com = source;
                Array keys = null;
                try { keys = com.Keys() as Array; }
                catch (Exception) { }
                if (keys != null)
                {
                    foreach (object key in keys)
                        result.Add(new KeyValuePair<object, object>(key, com.Item(key)));
                    return result;
                }
            }

            throw Error(13, "TypeError: esperado DictionaryS, Scripting.Dictionary ou Array(chave1, valor1...)");
        }

        // Linhas de um array 2D (como o Range.Value: base 1), cada uma um array
        static List<object> Rows(Array array)
        {
            List<object> rows = new List<object>();
            int c0 = array.GetLowerBound(1), c1 = array.GetUpperBound(1);
            for (int r = array.GetLowerBound(0); r <= array.GetUpperBound(0); r++)
            {
                object[] row = new object[c1 - c0 + 1];
                for (int c = c0; c <= c1; c++) row[c - c0] = array.GetValue(r, c);
                rows.Add(row);
            }
            return rows;
        }

        // Tamanho de um par candidato (Array ou ListS); -1 se não for um
        static int PairCount(object item)
        {
            Array a = item as Array;
            if (a != null) return a.Length;
            ListS l = item as ListS;
            if (l != null) return l.Items.Count;
            return -1;
        }
    }
}
