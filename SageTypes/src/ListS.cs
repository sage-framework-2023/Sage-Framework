using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace SageTypes
{
    // Tipo de ListS: lista (pode mudar) ou tupla (não pode), como no Sage.xlam
    [ComVisible(true), Guid("8D3F5A72-1E94-4C6B-A0D8-2F7B9E4C1A53")]
    public enum sgArrayTypes
    {
        sgList = 0,
        sgTuple = 1,
    }

    // ListS: lista como a do Python, no lugar do ListS do Sage.xlam (mesma API, mais os
    // métodos do Python).
    //
    //   Dim l As New Sage.ListS
    //   l = Array(3, 1, 2)
    //   l.Append 4                         ' altera a própria lista (e a devolve)
    //   Debug.Print l(-1), l.Slice(1, 3).ToString, l.Sort.ToString   ' 4  [1, 2]  [1, 2, 3, 4]
    //
    // Índices começam em 0; negativos contam do fim; fora do intervalo: IndexError
    // (erro 9). Os métodos que alteram a lista também a devolvem, então tanto
    // "l.Append x" quanto o encadeamento do Sage.xlam (".Remove(-1).Join(...)") funcionam.
    //
    // Interface dual: acrescente membros só no fim, com o próximo DispId (ver StringS).
    [ComVisible(true), Guid("5A1C8E3D-7B26-4F90-9D4E-B3C6A8F2D715"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _ListS
    {
        // l(i) lê; Set l(i) = objeto grava; l(i) = valor é o LetValue abaixo.
        // Sem índice: lê um array do VBA; l = Array(...) / outraLista substitui tudo.
        [DispId(0), IndexerName("Value")] object this[[Optional] object Index] { get; set; }
        [DispId(1)] sgArrayTypes ArrayType { get; set; }
        [DispId(2)] StringS Join([Optional] object Delimiter);
        [DispId(3)] int Length();
        [DispId(4)] ListS Remove(object Index);
        [DispId(5)] ListS Append(object Element);
        [DispId(6)] int Count { get; }
        [DispId(7)] ListS Extend(object Iterable);
        [DispId(8)] ListS Insert(object Index, object Element);
        [DispId(9)] object Pop([Optional] object Index);
        [DispId(10)] ListS RemoveValue(object Value);
        [DispId(11)] ListS Clear();
        [DispId(12)] int Index(object Value, [Optional] object Start, [Optional] object Stop);
        [DispId(13)] int CountOf(object Value);
        [DispId(14)] ListS Sort([Optional] object Reverse);
        [DispId(15)] ListS Reverse();
        [DispId(16)] ListS Copy();
        [DispId(17)] bool Contains(object Value);
        [DispId(18)] ListS Slice([Optional] object Start, [Optional] object Stop, [Optional] object Step);
        [DispId(19)] ListS Concat(object Other);
        [DispId(20)] ListS Repeat(int Times);
        [DispId(21)] object Sum();
        [DispId(22)] object Min();
        [DispId(23)] object Max();
        [DispId(24)] string ToString();
        [DispId(25)] object ToArray();
        [DispId(-4)] IEnumerator GetEnumerator();

        // Vira o Property Let de Value (DispId 0) na geração do .tlb (install.ps1).
        [DispId(1000), PropertyLet("Value")] void LetValue([Optional] object Index, object Value);

        // sort(key=lambda r: r[chave]) para listas de linhas (listas ou dicionários); altera a própria lista
        [DispId(26)] ListS SortBy(object Key, [Optional] object Reverse);
        // list(dict.fromkeys(l)): sem repetidos, na ordem em que aparecem
        [DispId(27)] ListS Unique();
        // sorted(l) e list(reversed(l)): listas novas; a original fica igual
        [DispId(28)] ListS Sorted([Optional] object Reverse);
        [DispId(29)] ListS Reversed();
    }

    [ComVisible(true), Guid("E7A42C19-6D3B-4E85-8F1A-0C9B5D2E7A64"), ProgId("Sage.ListS")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_ListS))]
    public sealed class ListS : _ListS, IEnumerable
    {
        readonly List<object> items = new List<object>();
        sgArrayTypes type = sgArrayTypes.sgList;

        public ListS() { }

        internal static ListS From(IEnumerable<object> source)
        {
            ListS list = new ListS();
            list.items.AddRange(source.Select(Interop.Store));
            return list;
        }

        internal IList<object> Items { get { return items; } }
        internal bool IsTuple { get { return type == sgArrayTypes.sgTuple; } }

        // ------------------------------------------------------------------
        // Valor (membro padrão)
        // ------------------------------------------------------------------

        // Texto volta como StringS (l(0).Upper); sem índice, array do VBA "puro"
        [IndexerName("Value")]
        public object this[object Index]
        {
            get { return Interop.IsMissing(Index) ? ToArray() : Interop.Wrap(items[Position(Index)]); }
            set { Store(Index, value); }
        }

        public void LetValue(object Index, object Value) { Store(Index, Value); }

        void Store(object index, object value)
        {
            Mutable();
            if (Interop.IsMissing(index))
            {
                // antes de limpar: pode ser a própria lista
                List<object> source = Interop.Items(value).Select(Interop.Store).ToList();
                items.Clear();
                items.AddRange(source);
            }
            else items[Position(index)] = Interop.Store(value);
        }

        public sgArrayTypes ArrayType
        {
            get { return type; }
            set { type = value; }
        }

        // ------------------------------------------------------------------
        // API do Sage.xlam
        // ------------------------------------------------------------------

        // Separador padrão " ", como no Sage.xlam
        public StringS Join(object Delimiter)
        {
            string delimiter = Interop.IsMissing(Delimiter) ? " " : Interop.Text(Delimiter);
            StringS result = new StringS();
            result.SetRaw(string.Join(delimiter, items.Select(Interop.Text).ToArray()));
            return result;
        }

        public int Length() { return items.Count; }

        // Por índice (negativo conta do fim), como no Sage.xlam; RemoveValue remove por
        // valor, como remove() do Python
        public ListS Remove(object Index)
        {
            Mutable();
            items.RemoveAt(Position(Index));
            return this;
        }

        public ListS Append(object Element)
        {
            Mutable();
            items.Add(Interop.Store(Element));
            return this;
        }

        // ------------------------------------------------------------------
        // Python
        // ------------------------------------------------------------------

        public int Count { get { return items.Count; } }

        // Acrescenta cada elemento de um array, ListS, DictionaryS (chaves), Collection...
        public ListS Extend(object Iterable)
        {
            Mutable();
            items.AddRange(Interop.Items(Iterable).Select(Interop.Store).ToList());
            return this;
        }

        // Como insert() do Python: índice além do fim acrescenta no fim
        public ListS Insert(object Index, object Element)
        {
            Mutable();
            int i = Interop.Integer(Index);
            if (i < 0) i = Math.Max(0, items.Count + i);
            items.Insert(Math.Min(i, items.Count), Interop.Store(Element));
            return this;
        }

        // Tira e devolve o elemento (o último, sem índice)
        public object Pop(object Index)
        {
            Mutable();
            if (items.Count == 0) throw Interop.Error(9, "IndexError: pop from empty list");
            int i = Interop.IsMissing(Index) ? items.Count - 1 : Position(Index);
            object value = items[i];
            items.RemoveAt(i);
            return Interop.Wrap(value);
        }

        public ListS RemoveValue(object Value)
        {
            Mutable();
            int i = Find(Value, 0, items.Count);
            if (i < 0) throw Interop.Error(5, "ValueError: list.remove(x): x not in list");
            items.RemoveAt(i);
            return this;
        }

        public ListS Clear()
        {
            Mutable();
            items.Clear();
            return this;
        }

        public int Index(object Value, object Start, object Stop)
        {
            int start = Bound(Start, 0), stop = Bound(Stop, items.Count);
            int i = Find(Value, start, stop);
            if (i < 0) throw Interop.Error(5, "ValueError: " + Interop.Repr(Value) + " is not in list");
            return i;
        }

        public int CountOf(object Value)
        {
            int count = 0;
            foreach (object item in items) if (KeyComparer.Instance.Equals(item, Value)) count++;
            return count;
        }

        // Ordenação estável, como sort() do Python: números pelo valor, texto pela ordem
        // dos caracteres (maiúsculas antes); misturar texto e número é TypeError
        public ListS Sort(object Reverse)
        {
            Mutable();
            bool descending = !Interop.IsMissing(Reverse) && Convert.ToBoolean(Reverse, CultureInfo.InvariantCulture);
            List<object> sorted = descending
                ? items.OrderByDescending(x => x, PyOrder.Instance).ToList()
                : items.OrderBy(x => x, PyOrder.Instance).ToList();
            items.Clear();
            items.AddRange(sorted);
            return this;
        }

        public ListS Reverse()
        {
            Mutable();
            items.Reverse();
            return this;
        }

        // Linhas que são listas (Key = número da coluna, negativo conta do fim) ou dicionários
        // (Key = chave); Key também pode ser Array(...) com várias, em ordem de prioridade.
        // Estável, como o sort do Python.
        public ListS SortBy(object Key, object Reverse)
        {
            Mutable();
            object raw = Interop.Unwrap(Key);
            List<object> keys = (raw is Array || Key is ListS) ? Interop.Items(Key).ToList() : new List<object> { Key };
            if (keys.Count == 0) throw Interop.Error(5, "ValueError: informe a coluna (ou a chave) para ordenar.");
            bool descending = !Interop.IsMissing(Reverse) && Convert.ToBoolean(Reverse, CultureInfo.InvariantCulture);
            List<KeyValuePair<object[], object>> rows = items.Select(item => new KeyValuePair<object[], object>(keys.Select(k => Field(item, k)).ToArray(), item)).ToList();
            IComparer<object[]> order = new RowOrder();
            List<object> sorted = (descending ? rows.OrderByDescending(r => r.Key, order) : rows.OrderBy(r => r.Key, order)).Select(r => r.Value).ToList();
            items.Clear();
            items.AddRange(sorted);
            return this;
        }

        static object Field(object row, object key)
        {
            ListS list = row as ListS;
            if (list != null) return Interop.Unwrap(list[key]);
            DictionaryS dict = row as DictionaryS;
            if (dict != null) return Interop.Unwrap(dict[key]);
            throw Interop.Error(13, "TypeError: SortBy ordena listas de linhas (listas ou dicionários), não " + Interop.TypeLabel(row) + ".");
        }

        sealed class RowOrder : IComparer<object[]>
        {
            public int Compare(object[] a, object[] b)
            {
                for (int i = 0; i < a.Length; i++)
                {
                    int c = PyOrder.Instance.Compare(a[i], b[i]);
                    if (c != 0) return c;
                }
                return 0;
            }
        }

        // Números iguais pelo valor (1 = 1.0), texto diferenciando maiúsculas, como no Python;
        // listas e dicionários iguais pelo conteúdo
        public ListS Unique()
        {
            HashSet<object> seen = new HashSet<object>(new UniqueKey());
            List<object> result = new List<object>();
            foreach (object item in items)
            {
                object key = item is ListS || item is DictionaryS ? (object)("\u0001" + Interop.Repr(item)) : (item ?? KeyComparer.Empty);
                if (seen.Add(key)) result.Add(item);
            }
            return From(result);
        }

        sealed class UniqueKey : IEqualityComparer<object>
        {
            public new bool Equals(object a, object b) { return KeyComparer.Instance.Equals(a, b); }
            public int GetHashCode(object o) { return KeyComparer.Instance.GetHashCode(o); }
        }

        public ListS Sorted(object Reverse) { return Copy().AsList().Sort(Reverse); }

        public ListS Reversed() { return Copy().AsList().Reverse(); }

        // Cópia que pode ser alterada (a de uma tupla também é tupla)
        ListS AsList()
        {
            type = sgArrayTypes.sgList;
            return this;
        }

        public ListS Copy()
        {
            ListS copy = From(items);
            copy.type = type;
            return copy;
        }

        public bool Contains(object Value) { return Find(Value, 0, items.Count) >= 0; }

        // Como l[start:stop:step] do Python; os omitidos seguem as mesmas regras
        public ListS Slice(object Start, object Stop, object Step)
        {
            int step = Interop.IsMissing(Step) ? 1 : Interop.Integer(Step);
            if (step == 0) throw Interop.Error(5, "ValueError: slice step cannot be zero");
            int n = items.Count;
            int start, stop;
            if (step > 0)
            {
                start = Interop.IsMissing(Start) ? 0 : Clamp(Interop.Integer(Start), n, 0, n);
                stop = Interop.IsMissing(Stop) ? n : Clamp(Interop.Integer(Stop), n, 0, n);
            }
            else
            {
                start = Interop.IsMissing(Start) ? n - 1 : Clamp(Interop.Integer(Start), n, -1, n - 1);
                stop = Interop.IsMissing(Stop) ? -1 : Clamp(Interop.Integer(Stop), n, -1, n - 1);
            }
            ListS result = new ListS();
            for (int i = start; step > 0 ? i < stop : i > stop; i += step) result.items.Add(items[i]);
            result.type = type;
            return result;
        }

        // Índice do Python numa fatia: negativo conta do fim, depois limita ao intervalo
        static int Clamp(int i, int n, int low, int high)
        {
            if (i < 0) i += n;
            return Math.Max(low, Math.Min(high, i));
        }

        // Como lista1 + lista2: nova lista
        public ListS Concat(object Other)
        {
            ListS result = Copy();
            result.items.AddRange(Interop.Items(Other).Select(Interop.Store).ToList());
            return result;
        }

        // Como lista * n: nova lista
        public ListS Repeat(int Times)
        {
            ListS result = new ListS();
            for (int i = 0; i < Times; i++) result.items.AddRange(items);
            result.type = type;
            return result;
        }

        // Soma: inteiros continuam inteiros (Long); com algum decimal, Double
        public object Sum()
        {
            long whole = 0;
            double real = 0;
            bool isReal = false;
            foreach (object item in items)
            {
                if (!PyOrder.IsNumber(item))
                    throw Interop.Error(13, "TypeError: unsupported operand type(s) for +: " + Interop.Repr(item));
                if (item is double || item is float || item is decimal) isReal = true;
                real += Convert.ToDouble(item, CultureInfo.InvariantCulture);
                if (!isReal) whole += Convert.ToInt64(item, CultureInfo.InvariantCulture);
            }
            if (isReal) return real;
            if (whole >= int.MinValue && whole <= int.MaxValue) return (int)whole;
            return whole;
        }

        public object Min()
        {
            if (items.Count == 0) throw Interop.Error(5, "ValueError: min() arg is an empty sequence");
            return Interop.Wrap(items.Aggregate((a, b) => PyOrder.Instance.Compare(b, a) < 0 ? b : a));
        }

        public object Max()
        {
            if (items.Count == 0) throw Interop.Error(5, "ValueError: max() arg is an empty sequence");
            return Interop.Wrap(items.Aggregate((a, b) => PyOrder.Instance.Compare(b, a) > 0 ? b : a));
        }

        // [1, 'a', [2, 3]]; tupla: (1, 'a')
        public override string ToString() { return Interop.Repr(this); }

        // Array do VBA (base 0), com as listas de dentro também como arrays e texto como String
        public object ToArray() { return Interop.Plain(this); }

        // For Each percorre uma cópia (a lista pode mudar no laço); texto vem como StringS
        public IEnumerator GetEnumerator() { return items.Select(Interop.Wrap).ToArray().GetEnumerator(); }

        // ------------------------------------------------------------------

        void Mutable()
        {
            if (type == sgArrayTypes.sgTuple) throw Interop.Error(13, "TypeError: 'tuple' object does not support item assignment");
        }

        // Índice válido (negativo conta do fim) ou IndexError
        int Position(object index)
        {
            int i = Interop.Integer(index);
            if (i < 0) i += items.Count;
            if (i < 0 || i >= items.Count) throw Interop.Error(9, "IndexError: list index out of range");
            return i;
        }

        int Bound(object value, int fallback)
        {
            if (Interop.IsMissing(value)) return fallback;
            return Clamp(Interop.Integer(value), items.Count, 0, items.Count);
        }

        int Find(object value, int start, int stop)
        {
            for (int i = start; i < stop; i++)
                if (KeyComparer.Instance.Equals(items[i], value)) return i;
            return -1;
        }
    }

    // Ordem do Python: números entre si, texto entre si (ordinal), datas entre si
    sealed class PyOrder : IComparer<object>
    {
        public static readonly PyOrder Instance = new PyOrder();

        public static bool IsNumber(object o)
        {
            return o is byte || o is short || o is int || o is long || o is float || o is double || o is decimal ||
                o is sbyte || o is ushort || o is uint || o is ulong;
        }

        public int Compare(object a, object b)
        {
            a = Interop.Unwrap(a);
            b = Interop.Unwrap(b);
            if (IsNumber(a) && IsNumber(b))
                return Convert.ToDouble(a, CultureInfo.InvariantCulture).CompareTo(Convert.ToDouble(b, CultureInfo.InvariantCulture));
            if (a is string && b is string) return string.CompareOrdinal((string)a, (string)b);
            if (a is DateTime && b is DateTime) return ((DateTime)a).CompareTo((DateTime)b);
            if (a is bool && b is bool) return ((bool)a).CompareTo((bool)b);
            throw Interop.Error(13, "TypeError: '<' not supported between " + Interop.TypeLabel(a) + " and " + Interop.TypeLabel(b));
        }
    }
}
