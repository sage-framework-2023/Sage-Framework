namespace SageToolbox
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;
    using IStream = System.Runtime.InteropServices.ComTypes.IStream;
    using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

    // A Caixa de ferramentas dos UserForms (Microsoft\Forms\EXCEL.box, no perfil do usuário): os
    // controles do Sage ficam numa página "Sage" própria. Usado pelo SageShortcuts (ToolboxSetup) e
    // pelo desinstalador (SageSetup), por isso sem dependências além do .NET.
    //
    // O arquivo é um form do MSForms gravado em OLE (IStorage). A Caixa é o MultiPage i01\i01:
    //   f    o form do MultiPage: um "site" para as abas (TabStrip) e um para cada página
    //        (nome "PageN", ID da página, BitFlags com 0x2 = a página à vista)
    //   o    as abas (TabStrip): legenda, dica, nome interno, tag e atalho de cada aba, e o TabData
    //   x    as páginas: um registro vazio por página (mais um), a quantidade e os IDs
    //   iNN  uma página (NN = ID): o seletor de controles, com a lista no stream pr
    // A lista de controles (pr): cabeçalho de 24 bytes ([8] próximo índice, [16] quantidade), um
    // registro de 32 bytes por controle (índice, 8 bytes zerados, tamanho do nome, CLSID) e os nomes
    // em UTF-16. Os índices ficam em sequência (1..N, o próximo em N + 1), senão o diálogo Controles
    // adicionais deixa de abrir.
    //
    // O formato foi conferido com uma página criada pelo próprio VBE (Nova página). Qualquer coisa
    // fora dele: nada é gravado e os controles vão para a primeira página, como antes.
    // O VBE mantém o arquivo aberto e o regrava ao fechar: só com o Excel fechado.
    static class ToolboxFile
    {
        public const string PageCaption = "Sage";

        // Põe os controles (CLSID, nome) na página "Sage", criada se faltar. Os controles do Sage que
        // estiverem em outras páginas (todos: as versões até a 1.2 os punham na página Controles)
        // passam para ela. false: arquivo fora do formato esperado (nada foi gravado)
        public static bool Install(string box, IList<KeyValuePair<Guid, string>> add, IList<KeyValuePair<Guid, string>> all)
        {
            return Open(box, delegate(Toolbox t)
            {
                Page sage = t.Find(PageCaption);
                List<KeyValuePair<Guid, string>> moved = new List<KeyValuePair<Guid, string>>();
                foreach (Page p in t.Pages)
                    if (p != sage)
                        foreach (KeyValuePair<Guid, string> c in all)
                            if (p.Palette.Contains(c.Key)) moved.Add(c);
                if (sage == null && add.Count == 0 && moved.Count == 0) return;
                if (sage == null) sage = t.AddPage(PageCaption);
                Page target = sage ?? t.Pages[0]; // sem como criar a página: a primeira, como antes
                List<Guid> ids = new List<Guid>();
                foreach (KeyValuePair<Guid, string> c in moved) ids.Add(c.Key);
                foreach (Page p in t.Pages)
                    if (p != target) p.Palette.Remove(ids);
                // Na ordem da lista do Sage
                foreach (KeyValuePair<Guid, string> c in all)
                    if (moved.Contains(c) || add.Contains(c)) target.Palette.Add(c.Key, c.Value);
            });
        }

        // Tira os controles de todas as páginas e apaga a página "Sage" se ela ficar vazia
        public static bool Remove(string box, IList<Guid> ids)
        {
            return Open(box, delegate(Toolbox t)
            {
                foreach (Page p in t.Pages) p.Palette.Remove(ids);
                Page sage = t.Find(PageCaption);
                if (sage != null && sage.Palette.Count == 0) t.RemovePage(sage);
            });
        }

        static bool Open(string box, Action<Toolbox> change)
        {
            IStorage root;
            if (StgOpenStorage(box, IntPtr.Zero, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out root) != 0) return false;
            Toolbox t = null;
            try
            {
                t = Toolbox.Read(root);
                if (t == null) return false;
                change(t);
                t.Write();
                root.Commit(0);
                return true;
            }
            finally
            {
                if (t != null) t.Release();
                Marshal.ReleaseComObject(root);
            }
        }

        // ------------------------------------------------------------------
        // O MultiPage (i01\i01) e as páginas
        // ------------------------------------------------------------------

        sealed class Page
        {
            public int Id;
            public IStorage Storage;
            public Palette Palette;
            public byte[] Site;        // o site da página no f (null: só sem MultiPage reconhecido)
            public Strings Tab;        // legenda, dica, nome, tag e atalho da aba
            public byte[] TabData;
            public bool Created;

            public void Release()
            {
                if (Storage != null) Marshal.ReleaseComObject(Storage);
                Storage = null;
            }
        }

        sealed class Strings { public byte[] Caption, Tip, Name, Tag, Accelerator; }

        sealed class Toolbox
        {
            public readonly List<Page> Pages = new List<Page>();
            IStorage outer, multi;
            bool layout;               // f, o e x reconhecidos: dá para criar e apagar páginas
            bool layoutChanged;
            readonly List<Page> removed = new List<Page>();

            // f
            byte[] formHead;           // cabeçalho e bloco de dados do form (até os sites)
            byte[] tabSite;            // o site das abas
            byte[] formTail;
            // o
            int listIndex, tabsAllocated;
            byte[] tabSize, textProps;
            // x
            List<byte[]> pageRecords;
            byte[] multiHead;          // máscara, quantidade e ID do MultiPage

            public static Toolbox Read(IStorage root)
            {
                Toolbox t = new Toolbox();
                t.outer = OpenStorage(root, "i01");
                t.multi = t.outer == null ? null : OpenStorage(t.outer, "i01");
                if (t.multi == null) { t.Release(); return null; }
                try { t.layout = t.ReadLayout(); }
                catch (Exception) { t.layout = false; }
                if (!t.layout)
                {
                    // Só a primeira página, como antes
                    foreach (Page p0 in t.Pages) p0.Release();
                    t.Pages.Clear();
                    IStorage s = OpenStorage(t.multi, "i04");
                    Palette p = s == null ? null : Palette.Read(s);
                    if (p == null) { if (s != null) Marshal.ReleaseComObject(s); t.Release(); return null; }
                    t.Pages.Add(new Page { Id = 4, Storage = s, Palette = p });
                }
                return t;
            }

            bool ReadLayout()
            {
                byte[] f = ReadStream(multi, "f"), o = ReadStream(multi, "o"), x = ReadStream(multi, "x");
                if (f == null || o == null || x == null) return false;

                // x: registros de página (00 02 04 00 + 4 bytes), depois o MultiPage (00 02 0C 00 + 12 bytes) e os IDs
                pageRecords = new List<byte[]>();
                int pos = 0;
                while (pos + 8 <= x.Length && U16(x, pos + 2) == 4) { pageRecords.Add(Slice(x, pos, 8)); pos += 8; }
                if (pos + 16 > x.Length || U16(x, pos) != 0x200 || U16(x, pos + 2) != 0x0C || I32(x, pos + 4) != 6) return false;
                multiHead = Slice(x, pos, 16);
                int count = I32(x, pos + 8);
                pos += 16;
                if (count < 1 || count > 64 || pos + count * 4 != x.Length || pageRecords.Count != count + 1) return false;
                List<int> ids = new List<int>();
                for (int i = 0; i < count; i++) ids.Add(I32(x, pos + i * 4));

                // o: as abas (TabStrip), com a máscara que o VBE grava
                if (U16(o, 0) != 0x200 || I32(o, 4) != 0x00FA8031) return false;
                int cb = U16(o, 2);
                listIndex = I32(o, 8);
                int itemsSize = I32(o, 12), tipsSize = I32(o, 16), namesSize = I32(o, 20);
                tabsAllocated = I32(o, 24);
                int tagsSize = I32(o, 28), tabDataCount = I32(o, 32), accelSize = I32(o, 36);
                if (tabDataCount != count) return false;
                tabSize = Slice(o, 40, 8);
                pos = 48;
                List<byte[]> captions = ReadStrings(o, ref pos, count, itemsSize), tips = ReadStrings(o, ref pos, count, tipsSize),
                    names = ReadStrings(o, ref pos, count, namesSize), tags = ReadStrings(o, ref pos, count, tagsSize),
                    accels = ReadStrings(o, ref pos, count, accelSize);
                if (captions == null || tips == null || names == null || tags == null || accels == null || pos != 4 + cb) return false;
                if (o.Length < pos + count * 4) return false;
                textProps = Slice(o, pos, o.Length - pos - count * 4);
                int dataAt = o.Length - count * 4;

                // f: o form do MultiPage; os sites depois do bloco de dados
                if (U16(f, 0) != 0x400) return false;
                int head = 4 + U16(f, 2);
                if (head + 8 > f.Length) return false;
                formHead = Slice(f, 0, head);
                int sites = I32(f, head), bytes = I32(f, head + 4);
                pos = head + 8;
                int end = pos + bytes;
                if (end > f.Length) return false;
                // Um tipo só de site (OleSiteConcrete), em uma entrada de profundidade/tipo
                if (f[pos] != 0 || (f[pos + 1] & 0x80) == 0 || (f[pos + 1] & 0x7F) != sites || f[pos + 2] != 1) return false;
                pos += 4;
                List<byte[]> siteList = new List<byte[]>();
                for (int i = 0; i < sites; i++)
                {
                    if (pos + 4 > end || U16(f, pos) != 0) return false;
                    int size = 4 + U16(f, pos + 2);
                    siteList.Add(Slice(f, pos, size));
                    pos += size;
                }
                if (pos != end) return false;
                formTail = Slice(f, end, f.Length - end);
                if (siteList.Count != count + 1) return false;
                tabSite = siteList[0];
                // O site das abas guarda o tamanho do stream o (ObjectStreamSize)
                if (tabSite.Length < 16 || (I32(tabSite, 4) & 0x20) == 0 || I32(tabSite, 12) != o.Length) return false;

                for (int i = 0; i < count; i++)
                {
                    byte[] site = siteList[i + 1];
                    if (site.Length < 24 || I32(site, 4) != 0x1D5 || I32(site, 12) != ids[i]) return false;
                    IStorage s = OpenStorage(multi, StorageName(ids[i]));
                    Palette p = s == null ? null : Palette.Read(s);
                    if (p == null) { if (s != null) Marshal.ReleaseComObject(s); return false; }
                    Pages.Add(new Page
                    {
                        Id = ids[i], Storage = s, Palette = p, Site = site,
                        Tab = new Strings { Caption = captions[i], Tip = tips[i], Name = names[i], Tag = tags[i], Accelerator = accels[i] },
                        TabData = Slice(o, dataAt + i * 4, 4),
                    });
                }
                return true;
            }

            public Page Find(string caption)
            {
                if (!layout) return null;
                foreach (Page p in Pages)
                    if (string.Equals(Decode(p.Tab.Caption), caption, StringComparison.OrdinalIgnoreCase)) return p;
                return null;
            }

            // Página nova no fim, como a "Nova página" do VBE (null: sem como criar)
            public Page AddPage(string caption)
            {
                if (!layout || Pages.Count == 0) return null;
                int id = 0;
                foreach (Page p in Pages) id = Math.Max(id, p.Id);
                id = Math.Max(id, I32(formHead, 8)) + 1;
                if (id > 9) return null; // o nome da pasta com dois dígitos: até aqui conferido
                Page first = Pages[0];

                IStorage s;
                if (multi.CreateStorage(StorageName(id), STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, out s) != 0) return null;
                STATSTG stat;
                first.Storage.Stat(out stat, 1);
                s.SetClass(ref stat.clsid);
                byte[] compObj = ReadStream(first.Storage, "\u0001CompObj");
                if (compObj != null) WriteStream(s, "\u0001CompObj", compObj, true);

                string baseName = Decode(first.Tab.Name).TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');
                tabsAllocated++;
                Page page = new Page
                {
                    Id = id, Storage = s, Palette = Palette.Empty(s), Created = true,
                    Site = PageSite(first.Site, id, Pages.Count + 1, false),
                    Tab = new Strings { Caption = Encode(caption), Tip = Encode(""), Name = Encode(baseName + tabsAllocated), Tag = Encode(""), Accelerator = Encode("") },
                    TabData = (byte[])first.TabData.Clone(),
                };
                Pages.Add(page);
                pageRecords.Add(new byte[] { 0x00, 0x02, 0x04, 0x00, 0, 0, 0, 0 });
                Array.Copy(BitConverter.GetBytes(id), 0, formHead, 8, 4);
                Array.Copy(BitConverter.GetBytes(I32(formHead, 16) + 1), 0, formHead, 16, 4);
                layoutChanged = true;
                return page;
            }

            public void RemovePage(Page page)
            {
                if (!layout || Pages.Count < 2) return;
                Pages.Remove(page);
                removed.Add(page);
                pageRecords.RemoveAt(pageRecords.Count - 1);
                if (listIndex >= Pages.Count) listIndex = 0;
                layoutChanged = true;
            }

            public void Write()
            {
                foreach (Page p in Pages) p.Palette.Write();
                if (!layoutChanged) return;

                // A página à vista: ListIndex nas abas e o bit 0x2 no site
                for (int i = 0; i < Pages.Count; i++)
                    Pages[i].Site = PageSite(Pages[i].Site, Pages[i].Id, i + 1, i == listIndex);

                // o
                MemoryStream strings = new MemoryStream();
                int[] sizes = new int[5];
                for (int k = 0; k < 5; k++)
                    foreach (Page p in Pages)
                    {
                        byte[] b = k == 0 ? p.Tab.Caption : k == 1 ? p.Tab.Tip : k == 2 ? p.Tab.Name : k == 3 ? p.Tab.Tag : p.Tab.Accelerator;
                        strings.Write(b, 0, b.Length);
                        sizes[k] += b.Length;
                    }
                MemoryStream o = new MemoryStream();
                int cb = 4 + 8 * 4 + tabSize.Length + (int)strings.Length;
                o.Write(new byte[] { 0x00, 0x02 }, 0, 2); o.Write(BitConverter.GetBytes((ushort)cb), 0, 2);
                Put(o, 0x00FA8031); Put(o, listIndex); Put(o, sizes[0]); Put(o, sizes[1]); Put(o, sizes[2]);
                Put(o, tabsAllocated); Put(o, sizes[3]); Put(o, Pages.Count); Put(o, sizes[4]);
                o.Write(tabSize, 0, tabSize.Length);
                strings.WriteTo(o);
                o.Write(textProps, 0, textProps.Length);
                foreach (Page p in Pages) o.Write(p.TabData, 0, 4);
                Array.Copy(BitConverter.GetBytes((int)o.Length), 0, tabSite, 12, 4);

                // f
                MemoryStream f = new MemoryStream();
                f.Write(formHead, 0, formHead.Length);
                int sites = Pages.Count + 1, bytes = 4 + tabSite.Length;
                foreach (Page p in Pages) bytes += p.Site.Length;
                Put(f, sites); Put(f, bytes);
                f.Write(new byte[] { 0, (byte)(0x80 | sites), 1, 0 }, 0, 4);
                f.Write(tabSite, 0, tabSite.Length);
                foreach (Page p in Pages) f.Write(p.Site, 0, p.Site.Length);
                f.Write(formTail, 0, formTail.Length);

                // x
                MemoryStream x = new MemoryStream();
                foreach (byte[] r in pageRecords) x.Write(r, 0, r.Length);
                byte[] mh = (byte[])multiHead.Clone();
                Array.Copy(BitConverter.GetBytes(Pages.Count), 0, mh, 8, 4);
                x.Write(mh, 0, mh.Length);
                foreach (Page p in Pages) Put(x, p.Id);

                WriteStream(multi, "f", f.ToArray(), false);
                WriteStream(multi, "o", o.ToArray(), false);
                WriteStream(multi, "x", x.ToArray(), false);
                foreach (Page p in removed)
                {
                    if (p.Storage != null) { Marshal.ReleaseComObject(p.Storage); p.Storage = null; }
                    multi.DestroyElement(StorageName(p.Id));
                }
                foreach (Page p in Pages) if (p.Created) p.Storage.Commit(0);
                multi.Commit(0);
                outer.Commit(0);
            }

            public void Release()
            {
                foreach (Page p in Pages) p.Release();
                foreach (Page p in removed) p.Release();
                if (multi != null) Marshal.ReleaseComObject(multi);
                if (outer != null) Marshal.ReleaseComObject(outer);
                multi = outer = null;
            }
        }

        // O site de uma página no f: o modelo de outra página com o ID, a ordem (TabIndex e o nome
        // "PageN") e o bit de página à vista trocados
        static byte[] PageSite(byte[] model, int id, int order, bool visible)
        {
            int nameLen = I32(model, 8) & 0x7FFFFFFF;
            int namePadded = (nameLen + 3) & ~3;
            byte[] position = Slice(model, 24 + namePadded, model.Length - 24 - namePadded);
            byte[] name = Encoding.ASCII.GetBytes("Page" + order);
            int padded = (name.Length + 3) & ~3;
            byte[] site = new byte[24 + padded + position.Length];
            Array.Copy(model, 0, site, 0, 24);
            Array.Copy(BitConverter.GetBytes((ushort)(site.Length - 4)), 0, site, 2, 2);
            Array.Copy(BitConverter.GetBytes(name.Length | unchecked((int)0x80000000)), 0, site, 8, 4);
            Array.Copy(BitConverter.GetBytes(id), 0, site, 12, 4);
            int flags = I32(model, 16);
            flags = visible ? flags | 2 : flags & ~2;
            Array.Copy(BitConverter.GetBytes(flags), 0, site, 16, 4);
            Array.Copy(BitConverter.GetBytes((ushort)order), 0, site, 20, 2);
            Array.Copy(name, 0, site, 24, name.Length);
            Array.Copy(position, 0, site, 24 + padded, position.Length);
            return site;
        }

        // ------------------------------------------------------------------
        // A lista de controles de uma página (pr)
        // ------------------------------------------------------------------

        sealed class Palette
        {
            IStorage storage;
            byte[] header;
            readonly List<byte[]> records = new List<byte[]>();
            readonly List<string> names = new List<string>();
            bool changed;

            public int Count { get { return records.Count; } }

            public static Palette Read(IStorage storage)
            {
                byte[] data = ReadStream(storage, "pr");
                if (data == null || data.Length < 24 || BitConverter.ToInt32(data, 0) != 0x00080200) return null;
                int count = BitConverter.ToInt32(data, 16);
                if (count < 0 || count > 200 || 24 + count * 32 > data.Length) return null;
                Palette p = new Palette { storage = storage, header = Slice(data, 0, 24) };
                int pos = 24 + count * 32;
                for (int i = 0; i < count; i++)
                {
                    byte[] r = Slice(data, 24 + i * 32, 32);
                    int chars = BitConverter.ToInt32(r, 12);
                    if (chars < 0 || pos + chars * 2 > data.Length) return null;
                    p.records.Add(r);
                    p.names.Add(Encoding.Unicode.GetString(data, pos, chars * 2));
                    pos += chars * 2;
                }
                return pos == data.Length ? p : null;
            }

            // Página nova: o cabeçalho que o VBE grava numa página vazia
            public static Palette Empty(IStorage storage)
            {
                return new Palette
                {
                    storage = storage, changed = true,
                    header = new byte[] { 0x00, 0x02, 0x08, 0x00, 0x08, 0, 0, 0, 0x01, 0, 0, 0, 0x02, 0, 0, 0, 0, 0, 0, 0, 0x20, 0, 0, 0 },
                };
            }

            public bool Contains(Guid id)
            {
                byte[] bytes = id.ToByteArray();
                foreach (byte[] r in records) if (Same(r, 16, bytes)) return true;
                return false;
            }

            public void Remove(IList<Guid> ids)
            {
                for (int i = records.Count - 1; i >= 0; i--)
                    foreach (Guid id in ids)
                        if (Same(records[i], 16, id.ToByteArray())) { records.RemoveAt(i); names.RemoveAt(i); changed = true; break; }
            }

            public void Add(Guid id, string name)
            {
                byte[] bytes = id.ToByteArray();
                foreach (byte[] r in records) if (Same(r, 16, bytes)) return;
                byte[] rec = new byte[32];
                Array.Copy(BitConverter.GetBytes(name.Length), 0, rec, 12, 4);
                Array.Copy(bytes, 0, rec, 16, 16);
                records.Add(rec);
                names.Add(name);
                changed = true;
            }

            public void Write()
            {
                if (!changed) return;
                for (int i = 0; i < records.Count; i++) Array.Copy(BitConverter.GetBytes(i + 1), 0, records[i], 0, 4);
                MemoryStream m = new MemoryStream();
                m.Write(header, 0, 24);
                foreach (byte[] r in records) m.Write(r, 0, 32);
                foreach (string n in names) { byte[] b = Encoding.Unicode.GetBytes(n); m.Write(b, 0, b.Length); }
                byte[] data = m.ToArray();
                Array.Copy(BitConverter.GetBytes(records.Count + 1), 0, data, 8, 4);
                Array.Copy(BitConverter.GetBytes(records.Count), 0, data, 16, 4);
                WriteStream(storage, "pr", data, true);
                storage.Commit(0);
            }
        }

        // ------------------------------------------------------------------
        // Textos do MSForms: tamanho (com 0x80000000 = 1 byte por caractere) e o texto, em múltiplos de 4
        // ------------------------------------------------------------------

        static List<byte[]> ReadStrings(byte[] data, ref int pos, int count, int size)
        {
            List<byte[]> list = new List<byte[]>();
            int start = pos;
            for (int i = 0; i < count; i++)
            {
                if (pos + 4 > data.Length) return null;
                int len = I32(data, pos) & 0x7FFFFFFF;
                int total = 4 + ((len + 3) & ~3);
                if (pos + total > data.Length) return null;
                list.Add(Slice(data, pos, total));
                pos += total;
            }
            return pos - start == size ? list : null;
        }

        static string Decode(byte[] s)
        {
            int raw = I32(s, 0), len = raw & 0x7FFFFFFF;
            return (raw & unchecked((int)0x80000000)) != 0 ? Encoding.GetEncoding(1252).GetString(s, 4, len) : Encoding.Unicode.GetString(s, 4, len);
        }

        static byte[] Encode(string text)
        {
            byte[] chars = Encoding.GetEncoding(1252).GetBytes(text);
            byte[] s = new byte[4 + ((chars.Length + 3) & ~3)];
            Array.Copy(BitConverter.GetBytes(chars.Length == 0 ? 0 : chars.Length | unchecked((int)0x80000000)), 0, s, 0, 4);
            Array.Copy(chars, 0, s, 4, chars.Length);
            return s;
        }

        // ------------------------------------------------------------------
        // Bytes e IStorage
        // ------------------------------------------------------------------

        static string StorageName(int id) { return "i" + id.ToString("00"); }
        static int I32(byte[] b, int at) { return BitConverter.ToInt32(b, at); }
        static int U16(byte[] b, int at) { return BitConverter.ToUInt16(b, at); }
        static void Put(MemoryStream m, int v) { m.Write(BitConverter.GetBytes(v), 0, 4); }

        static byte[] Slice(byte[] b, int at, int len)
        {
            byte[] r = new byte[len];
            Array.Copy(b, at, r, 0, len);
            return r;
        }

        static bool Same(byte[] a, int offset, byte[] b)
        {
            for (int i = 0; i < b.Length; i++) if (a[offset + i] != b[i]) return false;
            return true;
        }

        static IStorage OpenStorage(IStorage parent, string name)
        {
            IStorage child;
            return parent.OpenStorage(name, null, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out child) == 0 ? child : null;
        }

        static byte[] ReadStream(IStorage storage, string name)
        {
            IStream stream;
            if (storage.OpenStream(name, IntPtr.Zero, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, out stream) != 0) return null;
            try
            {
                STATSTG stat;
                stream.Stat(out stat, 1);
                byte[] data = new byte[stat.cbSize];
                IntPtr read = Marshal.AllocCoTaskMem(4);
                try { if (data.Length > 0) stream.Read(data, data.Length, read); }
                finally { Marshal.FreeCoTaskMem(read); }
                return data;
            }
            finally { Marshal.ReleaseComObject(stream); }
        }

        static void WriteStream(IStorage storage, string name, byte[] data, bool create)
        {
            IStream stream;
            int hr = create
                ? storage.CreateStream(name, STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, out stream)
                : storage.OpenStream(name, IntPtr.Zero, STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, out stream);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            try
            {
                stream.Seek(0, 0, IntPtr.Zero);
                stream.SetSize(data.Length);
                if (data.Length > 0) stream.Write(data, data.Length, IntPtr.Zero);
                stream.Commit(0);
            }
            finally { Marshal.ReleaseComObject(stream); }
        }

        const int STGM_READ = 0x0, STGM_READWRITE = 0x2, STGM_SHARE_EXCLUSIVE = 0x10, STGM_CREATE = 0x1000;

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        static extern int StgOpenStorage(string name, IntPtr priority, int mode, IntPtr exclude, int reserved, out IStorage storage);

        [ComImport, Guid("0000000B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IStorage
        {
            [PreserveSig] int CreateStream([MarshalAs(UnmanagedType.LPWStr)] string name, int mode, int r1, int r2, out IStream stream);
            [PreserveSig] int OpenStream([MarshalAs(UnmanagedType.LPWStr)] string name, IntPtr r1, int mode, int r2, out IStream stream);
            [PreserveSig] int CreateStorage([MarshalAs(UnmanagedType.LPWStr)] string name, int mode, int r1, int r2, out IStorage storage);
            [PreserveSig] int OpenStorage([MarshalAs(UnmanagedType.LPWStr)] string name, IStorage priority, int mode, IntPtr exclude, int reserved, out IStorage storage);
            void CopyTo(); void MoveElementTo();
            [PreserveSig] int Commit(int flags);
            void Revert(); void EnumElements();
            [PreserveSig] int DestroyElement([MarshalAs(UnmanagedType.LPWStr)] string name);
            void RenameElement(); void SetElementTimes();
            [PreserveSig] int SetClass(ref Guid clsid);
            void SetStateBits();
            void Stat(out STATSTG stat, int flags);
        }
    }
}
