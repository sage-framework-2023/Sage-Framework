using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace SageTypes
{
    // DictionaryS: dicionário como o do Python, no lugar do DictionaryS do Sage.xlam
    // (mesma API, mais os métodos do Python).
    //
    //   Dim d As New Sage.DictionaryS
    //   d("nome") = "Ana": d(1) = 10: Set d("lista") = outroDicionario
    //   For Each k In d: Debug.Print k, d(k): Next   ' chaves, na ordem de inserção
    //   Debug.Print d.ToString                       ' {'nome': 'Ana', 1: 10, ...}
    //
    // Como no Python: chave inexistente gera erro (KeyError, erro 9); chaves de qualquer
    // tipo imutável, com 1, 1.0 e CLng(1) iguais e "1" diferente; arrays não são chave.
    //
    // Interface dual: acrescente membros só no fim, com o próximo DispId (ver StringS).
    [ComVisible(true), Guid("C2D7A9E4-5B13-4F68-8E0A-7D4B1C9F3E62"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _DictionaryS
    {
        // d(chave) lê; Set d(chave) = objeto grava. d(chave) = valor (Property Let) é o
        // LetValue abaixo. Sem chave: lê um Scripting.Dictionary com tudo (como no
        // Sage.xlam); d = Array(k1, v1, k2, v2...) ou d = outroDicionario substitui tudo.
        [DispId(0), IndexerName("Value")] object this[[Optional] object Key] { get; set; }
        [DispId(1)] object Item(object Key);
        [DispId(2)] object Key(object Key);
        [DispId(3)] object GetItem(object Key, [Optional] object DefaultValue);
        [DispId(4)] bool Exists(object Key);
        [DispId(5)] ListS Keys();
        [DispId(6)] ListS Itens();
        [DispId(7)] int Length();
        [DispId(8)] DictionaryS Copy();
        [DispId(9)] void Clear();
        [DispId(10)] void Remove(object Key);
        [DispId(11)] void Append([Optional] object Value);
        [DispId(12)] int Count { get; }
        [DispId(13)] bool Contains(object Key);
        [DispId(14)] ListS Values();
        [DispId(15)] ListS Items();
        [DispId(16)] object Pop(object Key, [Optional] object DefaultValue);
        [DispId(17)] object PopItem();
        [DispId(18)] object SetDefault(object Key, [Optional] object DefaultValue);
        [DispId(19)] void Update(object Other);
        [DispId(20)] DictionaryS FromKeys(object Keys, [Optional] object Value);
        [DispId(21)] string ToString();
        [DispId(-4)] IEnumerator GetEnumerator();

        // Vira o Property Let de Value (DispId 0) na geração do .tlb (install.ps1): o
        // .NET só exporta Property Set para valores Variant. Fica sempre por último.
        [DispId(1000), PropertyLet("Value")] void LetValue([Optional] object Key, object Value);

        // Do Scripting.Dictionary, para o código VBA existente funcionar igual
        [DispId(22)] void Add(object Key, object Item);
        [DispId(23)] void RemoveAll();
        // d | outro do Python 3.9: dicionário novo com os dois (o de outro vence nas chaves repetidas)
        [DispId(24)] DictionaryS Union(object Other);
    }

    [ComVisible(true), Guid("4B8E6D21-9F3C-4A57-B1D0-E5C27A8F6B39"), ProgId("Sage.DictionaryS")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_DictionaryS))]
    public sealed class DictionaryS : _DictionaryS, IEnumerable
    {
        readonly OrderedDictionary entries = new OrderedDictionary(KeyComparer.Instance);

        public DictionaryS() { }

        // ------------------------------------------------------------------
        // Valor (membro padrão)
        // ------------------------------------------------------------------

        [IndexerName("Value")]
        public object this[object Key]
        {
            get
            {
                if (Interop.IsMissing(Key)) return ToScriptingDictionary();
                return Get(Key);
            }
            set { Store(Key, value); }
        }

        public void LetValue(object Key, object Value) { Store(Key, Value); }

        void Store(object key, object value)
        {
            if (Interop.IsMissing(key))
            {
                entries.Clear();
                Load(value, true);
                return;
            }
            entries[KeyOf(key)] = Interop.Store(value);
        }

        // Texto volta como StringS (d("nome").Upper), como na ListS
        object Get(object key)
        {
            object k = KeyOf(key);
            if (!entries.Contains(k)) throw Interop.Error(9, "KeyError: " + Interop.Repr(k));
            return Interop.Wrap(entries[k]);
        }

        // ------------------------------------------------------------------
        // API do Sage.xlam
        // ------------------------------------------------------------------

        public object Item(object Key) { return Get(Key); }

        // A chave como foi guardada (1 para Key(1.0), se a primeira foi 1)
        public object Key(object Key)
        {
            object k = KeyOf(Key);
            foreach (object stored in entries.Keys)
                if (KeyComparer.Instance.Equals(stored, k)) return Interop.Wrap(PlainKey(stored));
            throw Interop.Error(9, "KeyError: " + Interop.Repr(k));
        }

        // Como get() do Python: sem a chave, devolve o padrão (Empty se omitido)
        public object GetItem(object Key, object DefaultValue)
        {
            object k = KeyOf(Key);
            if (entries.Contains(k)) return Interop.Wrap(entries[k]);
            return Interop.IsMissing(DefaultValue) ? null : Interop.Wrap(DefaultValue);
        }

        public bool Exists(object Key) { return Contains(Key); }

        public ListS Keys() { return ListS.From((object[])KeysArray()); }

        // Chaves como o VBA as vê (a chave Empty volta como Empty)
        internal object KeysArray()
        {
            object[] result = new object[entries.Count];
            entries.Keys.CopyTo(result, 0);
            for (int i = 0; i < result.Length; i++) result[i] = PlainKey(result[i]);
            return result;
        }

        static object PlainKey(object key) { return key == KeyComparer.Empty ? null : key; }

        public ListS Itens() { return Values(); }

        public int Length() { return entries.Count; }

        // Cópia rasa (os objetos guardados são os mesmos), como copy() do Python
        public DictionaryS Copy()
        {
            DictionaryS copy = new DictionaryS();
            foreach (DictionaryEntry e in entries) copy.entries.Add(e.Key, e.Value);
            return copy;
        }

        public void Clear() { entries.Clear(); }

        // Como o Scripting.Dictionary: chave repetida é erro (457); para substituir, d(chave) = valor
        public void Add(object Key, object Item)
        {
            object k = KeyOf(Key);
            if (entries.Contains(k))
                throw Interop.Error(457, "KeyError: a chave " + Interop.Repr(k) + " já existe (para substituir o valor, use d(chave) = valor).");
            entries[k] = Interop.Store(Item);
        }

        public void RemoveAll() { Clear(); }

        public DictionaryS Union(object Other)
        {
            DictionaryS result = Copy();
            result.Update(Other);
            return result;
        }

        // Sem erro se a chave não existe (como no Sage.xlam); Pop gera KeyError
        public void Remove(object Key) { entries.Remove(KeyOf(Key)); }

        // Acrescenta pares (array k1, v1, k2, v2... ou outro dicionário) sem trocar os
        // valores das chaves que já existem
        public void Append(object Value)
        {
            if (!Interop.IsMissing(Value)) Load(Value, false);
        }

        // ------------------------------------------------------------------
        // Python
        // ------------------------------------------------------------------

        public int Count { get { return entries.Count; } }

        public bool Contains(object Key) { return entries.Contains(KeyOf(Key)); }

        public ListS Values()
        {
            object[] result = new object[entries.Count];
            entries.Values.CopyTo(result, 0);
            return ListS.From(result);
        }

        // Pares (chave, valor), cada um uma ListS, como items() do Python
        public ListS Items()
        {
            List<object> result = new List<object>(entries.Count);
            foreach (DictionaryEntry e in entries) result.Add(ListS.From(new object[] { PlainKey(e.Key), e.Value }));
            return ListS.From(result);
        }

        public object Pop(object Key, object DefaultValue)
        {
            object k = KeyOf(Key);
            if (!entries.Contains(k))
            {
                if (Interop.IsMissing(DefaultValue)) throw Interop.Error(9, "KeyError: " + Interop.Repr(k));
                return Interop.Wrap(DefaultValue);
            }
            object value = entries[k];
            entries.Remove(k);
            return Interop.Wrap(value);
        }

        // Tira e devolve o último par inserido, ListS (chave, valor)
        public object PopItem()
        {
            if (entries.Count == 0) throw Interop.Error(9, "KeyError: 'popitem(): dictionary is empty'");
            int last = entries.Count - 1;
            object[] keys = new object[entries.Count];
            entries.Keys.CopyTo(keys, 0);
            ListS pair = ListS.From(new object[] { PlainKey(keys[last]), entries[last] });
            entries.RemoveAt(last);
            return pair;
        }

        public object SetDefault(object Key, object DefaultValue)
        {
            object k = KeyOf(Key);
            if (!entries.Contains(k)) entries[k] = Interop.IsMissing(DefaultValue) ? null : Interop.Store(DefaultValue);
            return Interop.Wrap(entries[k]);
        }

        // Como update() do Python: acrescenta e substitui
        public void Update(object Other) { Load(Other, true); }

        public DictionaryS FromKeys(object Keys, object Value)
        {
            DictionaryS result = new DictionaryS();
            object value = Interop.IsMissing(Value) ? null : Interop.Store(Value);
            foreach (object k in Interop.Items(Keys)) result.entries[result.KeyOf(k)] = value;
            return result;
        }

        // {'a': 1, 'b': [1, 2]}, como repr() do Python
        public override string ToString()
        {
            return Interop.Repr(this);
        }

        internal IEnumerable<KeyValuePair<object, object>> Pairs()
        {
            foreach (DictionaryEntry e in entries) yield return new KeyValuePair<object, object>(e.Key, e.Value);
        }

        // For Each percorre uma cópia das chaves (o dicionário pode mudar no laço); texto vem como StringS
        public IEnumerator GetEnumerator()
        {
            object[] keys = (object[])KeysArray();
            for (int i = 0; i < keys.Length; i++) keys[i] = Interop.Wrap(keys[i]);
            return keys.GetEnumerator();
        }

        // ------------------------------------------------------------------

        // Chave guardada: StringS vira o texto dele (chaves são imutáveis, como no Python)
        object KeyOf(object key)
        {
            if (key is Array) throw Interop.Error(13, "TypeError: unhashable type: 'Array'");
            StringS s = key as StringS;
            if (s != null) return s.Value;
            return key ?? KeyComparer.Empty;
        }

        // Array (k1, v1, k2, v2...), DictionaryS ou Scripting.Dictionary
        void Load(object source, bool replace)
        {
            foreach (KeyValuePair<object, object> pair in Interop.PairsOf(source))
            {
                object k = KeyOf(pair.Key);
                if (replace || !entries.Contains(k)) entries[k] = Interop.Store(pair.Value);
            }
        }

        object ToScriptingDictionary()
        {
            dynamic dict = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary"));
            foreach (DictionaryEntry e in entries)
                dict.Add(PlainKey(e.Key), Interop.Plain(e.Value)); // VBA "puro": listas como arrays
            return dict;
        }
    }

    // Igualdade de chaves como no Python: números pelo valor (1 = 1.0 = CLng(1)),
    // texto diferenciando maiúsculas, datas pelo instante, objetos pela identidade.
    sealed class KeyComparer : IEqualityComparer
    {
        public static readonly KeyComparer Instance = new KeyComparer();
        public static readonly object Empty = new EmptyKey(); // chave Empty/Nothing

        sealed class EmptyKey { public override string ToString() { return "None"; } }

        static bool IsNumber(object o) { return PyOrder.IsNumber(o); }

        // Também compara elementos de ListS (Contains, Index, CountOf): StringS vale o texto
        public new bool Equals(object a, object b)
        {
            a = Interop.Unwrap(a);
            b = Interop.Unwrap(b);
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null) return false;
            if (IsNumber(a) && IsNumber(b)) return Convert.ToDouble(a, CultureInfo.InvariantCulture) == Convert.ToDouble(b, CultureInfo.InvariantCulture);
            string sa = a as string, sb = b as string;
            if (sa != null || sb != null) return sa != null && sb != null && string.Equals(sa, sb, StringComparison.Ordinal);
            return a.Equals(b);
        }

        public int GetHashCode(object o)
        {
            o = Interop.Unwrap(o);
            if (o == null) return 0;
            if (IsNumber(o)) return Convert.ToDouble(o, CultureInfo.InvariantCulture).GetHashCode();
            string s = o as string;
            if (s != null) return StringComparer.Ordinal.GetHashCode(s);
            return o.GetHashCode();
        }
    }
}
