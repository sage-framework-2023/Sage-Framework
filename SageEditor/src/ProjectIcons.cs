using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace SageEditor
{
    // Ícones do vscode-icons (MIT, github.com/vscode-icons/vscode-icons) na janela Projeto e nas
    // abas do código. Os PNGs (16 a 32 px) ficam embutidos no DLL (res\icons).
    //
    // A árvore da janela Projeto é um SysTreeView32 com uma lista de imagens; o VBE escolhe o
    // ícone de cada item pela posição na lista (3 projeto, 4 formulário, 7 módulo, 8 classe,
    // 12/13 pasta aberta/fechada, 20 pasta de trabalho, 21 planilha, 22 gráfico). O Sage põe uma
    // lista própria, de 32 bits, com as mesmas posições (os ícones trocados e os originais nas
    // demais) e as pastas coloridas no fim: cada pasta recebe a sua conforme o que tem dentro,
    // e volta para ela quando o VBE a troca ao expandir ou recolher.
    static class ProjectIcons
    {
        static readonly Dictionary<int, string> replaced = new Dictionary<int, string>
        {
            { 3, "file_type_vbproj" }, { 4, "file_type_xaml" }, { 7, "file_type_vba" }, { 8, "vba_class" },
            { 20, "file_type_excel" }, { 21, "file_type_excel2" }, { 22, "file_type_excel2" },
        };
        const int FolderOpen = 12, FolderClosed = 13;
        static readonly string[] folderKinds = { "default_folder", "folder_type_src", "folder_type_view", "folder_type_module", "folder_type_component", "folder_type_library" };

        static IntPtr vbeWindow, tree, original, ours;
        static int folderBase = -1, tick;
        static readonly Native.SubclassProc parentProc = ParentProc;
        static readonly UIntPtr SubclassId = (UIntPtr)0x5A73;
        static IntPtr parent;
        // Nome do componente -> ícone, lido da árvore (para as abas)
        static readonly Dictionary<string, string> byName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public static void Start(IntPtr vbe) { vbeWindow = vbe; }

        public static void Poll()
        {
            if (vbeWindow == IntPtr.Zero || ++tick % 4 != 0) return;
            if (!Settings.Icons) { Restore(); return; }
            if (tree == IntPtr.Zero || !Native.IsWindow(tree)) Attach();
            if (tree == IntPtr.Zero) return;
            IntPtr current = SendMessage(tree, TVM_GETIMAGELIST, (IntPtr)TVSIL_NORMAL, IntPtr.Zero);
            if (current != ours)
            {
                if (current == IntPtr.Zero) return;
                original = current;
                Build();
                if (ours == IntPtr.Zero) return;
                SendMessage(tree, TVM_SETIMAGELIST, (IntPtr)TVSIL_NORMAL, ours);
            }
            FixItems();
        }

        static void Attach()
        {
            tree = IntPtr.Zero;
            Native.EnumChildWindows(vbeWindow, delegate(IntPtr h, IntPtr l)
            {
                if (Native.ClassName(h) == "SysTreeView32" && Native.ClassName(Native.GetParent(h)) == "PROJECT") { tree = h; return false; }
                return true;
            }, IntPtr.Zero);
            // Janela Projeto flutuante: fora da janela principal
            if (tree == IntPtr.Zero)
                Native.EnumThreadWindows(Native.GetCurrentThreadId(), delegate(IntPtr top, IntPtr l)
                {
                    if (top == vbeWindow) return true;
                    Native.EnumChildWindows(top, delegate(IntPtr h, IntPtr l2)
                    {
                        if (Native.ClassName(h) == "SysTreeView32" && Native.ClassName(Native.GetParent(h)) == "PROJECT") { tree = h; return false; }
                        return true;
                    }, IntPtr.Zero);
                    return tree == IntPtr.Zero;
                }, IntPtr.Zero);
            if (tree == IntPtr.Zero) return;
            IntPtr p = Native.GetParent(tree);
            if (p != parent)
            {
                if (parent != IntPtr.Zero && Native.IsWindow(parent)) Native.RemoveWindowSubclass(parent, parentProc, SubclassId);
                parent = p;
                Native.SetWindowSubclass(parent, parentProc, SubclassId, UIntPtr.Zero);
            }
            ours = IntPtr.Zero; // árvore nova: monta a lista de novo
        }

        // Lista própria: as posições do VBE (trocadas ou copiadas) e as pastas coloridas no fim
        static void Build()
        {
            int cx, cy;
            if (!ImageList_GetIconSize(original, out cx, out cy) || cx <= 0) return;
            int count = ImageList_GetImageCount(original);
            IntPtr list = ImageList_Create(cx, cy, ILC_COLOR32, count + folderKinds.Length * 2, 4);
            for (int i = 0; i < count; i++)
            {
                string name;
                IntPtr icon = replaced.TryGetValue(i, out name) ? IconHandle(name, cx) : ImageList_GetIcon(original, i, ILD_TRANSPARENT);
                ImageList_ReplaceIcon(list, -1, icon);
                DestroyIcon(icon);
            }
            folderBase = count;
            foreach (string kind in folderKinds)
            {
                foreach (string name in new[] { kind, kind + "_opened" })
                {
                    IntPtr icon = IconHandle(name, cx);
                    ImageList_ReplaceIcon(list, -1, icon);
                    DestroyIcon(icon);
                }
            }
            if (ours != IntPtr.Zero && ours != list) ImageList_Destroy(ours);
            ours = list;
        }

        // Cada pasta com o ícone do que tem dentro, aberto ou fechado; e o mapa nome -> ícone
        static void FixItems()
        {
            string before = Signature();
            byName.Clear();
            Walk(SendMessage(tree, TVM_GETNEXTITEM, (IntPtr)TVGN_ROOT, IntPtr.Zero));
            // As abas leem o ícone daqui: redesenha quando os componentes mudam
            if (Signature() != before) EditorTabs.Refresh();
        }

        static string Signature()
        {
            StringBuilder sb = new StringBuilder();
            foreach (KeyValuePair<string, string> p in byName) sb.Append(p.Key).Append('=').Append(p.Value).Append(';');
            return sb.ToString();
        }

        static void Walk(IntPtr item)
        {
            while (item != IntPtr.Zero)
            {
                string text;
                TVITEM it = Get(item, true, out text);
                IntPtr child = SendMessage(tree, TVM_GETNEXTITEM, (IntPtr)TVGN_CHILD, item);
                bool folder = it.iImage == FolderOpen || it.iImage == FolderClosed ||
                    (folderBase >= 0 && it.iImage >= folderBase && it.iImage < folderBase + folderKinds.Length * 2);
                if (folder)
                {
                    int kind = child == IntPtr.Zero ? 0 : KindOf(Get(child).iImage);
                    bool expanded = (it.state & TVIS_EXPANDED) != 0;
                    int want = folderBase + kind * 2 + (expanded ? 1 : 0);
                    if (it.iImage != want || it.iSelectedImage != want)
                    {
                        TVITEM set = new TVITEM { mask = TVIF_HANDLE | TVIF_IMAGE | TVIF_SELECTEDIMAGE, hItem = item, iImage = want, iSelectedImage = want };
                        SendItem(tree, TVM_SETITEMW, IntPtr.Zero, ref set);
                    }
                }
                else
                {
                    string name;
                    if (replaced.TryGetValue(it.iImage, out name) && text != null)
                    {
                        // "Planilha1 (Planilha1)", "VBAProject (Pasta1.xlsm)": o nome vem antes do parêntese
                        string key = text;
                        int p = key.IndexOf(" (", StringComparison.Ordinal);
                        if (p > 0) key = key.Substring(0, p);
                        byName[key] = name;
                    }
                }
                if (child != IntPtr.Zero) Walk(child);
                item = SendMessage(tree, TVM_GETNEXTITEM, (IntPtr)TVGN_NEXT, item);
            }
        }

        // Pelo ícone do primeiro item dentro da pasta
        static int KindOf(int image)
        {
            switch (image)
            {
                case 20: case 21: case 22: return 1; // Microsoft Excel Objetos
                case 4: return 2;                     // Formulários
                case 7: return 3;                     // Módulos
                case 8: return 4;                     // Módulos de classe
            }
            return image >= 0 && image < 3 ? 5 : 0;  // Referências (e o resto: pasta comum)
        }

        static TVITEM Get(IntPtr item)
        {
            string ignored;
            return Get(item, false, out ignored);
        }

        static TVITEM Get(IntPtr item, bool withText, out string text)
        {
            text = null;
            IntPtr buffer = withText ? Marshal.AllocHGlobal(520) : IntPtr.Zero;
            try
            {
                TVITEM it = new TVITEM { mask = TVIF_HANDLE | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_STATE | (withText ? TVIF_TEXT : 0), hItem = item, stateMask = TVIS_EXPANDED, pszText = buffer, cchTextMax = 259 };
                SendItem(tree, TVM_GETITEMW, IntPtr.Zero, ref it);
                if (withText) text = Marshal.PtrToStringUni(buffer);
                return it;
            }
            finally { if (buffer != IntPtr.Zero) Marshal.FreeHGlobal(buffer); }
        }

        // Ao expandir ou recolher, o VBE volta a pasta para o ícone dele: o Sage a corrige na hora
        static IntPtr ParentProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, UIntPtr id, UIntPtr refData)
        {
            IntPtr result = Native.DefSubclassProc(hwnd, msg, wParam, lParam);
            try
            {
                if (msg == WM_NOTIFY && ours != IntPtr.Zero && Settings.Icons)
                {
                    NMHDR h = (NMHDR)Marshal.PtrToStructure(lParam, typeof(NMHDR));
                    if (h.hwndFrom == tree && (h.code == TVN_ITEMEXPANDEDW || h.code == TVN_ITEMEXPANDEDA)) FixItems();
                }
                else if (msg == Native.WM_NCDESTROY)
                {
                    Native.RemoveWindowSubclass(hwnd, parentProc, SubclassId);
                    if (hwnd == parent) parent = IntPtr.Zero;
                }
            }
            catch (Exception ex) { Log.Error(ex); }
            return result;
        }

        // Volta a lista e os ícones de pasta do VBE (configuração desligada ou add-in descarregado)
        static void Restore()
        {
            if (ours == IntPtr.Zero) return;
            if (tree != IntPtr.Zero && Native.IsWindow(tree) && original != IntPtr.Zero)
            {
                SendMessage(tree, TVM_SETIMAGELIST, (IntPtr)TVSIL_NORMAL, original);
                RestoreFolders(SendMessage(tree, TVM_GETNEXTITEM, (IntPtr)TVGN_ROOT, IntPtr.Zero));
            }
            ImageList_Destroy(ours);
            ours = IntPtr.Zero;
            folderBase = -1;
            byName.Clear();
        }

        static void RestoreFolders(IntPtr item)
        {
            while (item != IntPtr.Zero)
            {
                TVITEM it = Get(item);
                if (folderBase >= 0 && it.iImage >= folderBase)
                {
                    int want = (it.state & TVIS_EXPANDED) != 0 ? FolderOpen : FolderClosed;
                    TVITEM set = new TVITEM { mask = TVIF_HANDLE | TVIF_IMAGE | TVIF_SELECTEDIMAGE, hItem = item, iImage = want, iSelectedImage = want };
                    SendItem(tree, TVM_SETITEMW, IntPtr.Zero, ref set);
                }
                IntPtr child = SendMessage(tree, TVM_GETNEXTITEM, (IntPtr)TVGN_CHILD, item);
                if (child != IntPtr.Zero) RestoreFolders(child);
                item = SendMessage(tree, TVM_GETNEXTITEM, (IntPtr)TVGN_NEXT, item);
            }
        }

        public static void Shutdown()
        {
            Restore();
            if (parent != IntPtr.Zero && Native.IsWindow(parent)) Native.RemoveWindowSubclass(parent, parentProc, SubclassId);
            parent = IntPtr.Zero;
            tree = IntPtr.Zero;
            foreach (Bitmap b in cache.Values) b.Dispose();
            cache.Clear();
        }

        // ------------------------------------------------------------------
        // Imagens
        // ------------------------------------------------------------------

        static readonly int[] sizes = { 16, 20, 24, 28, 32 };
        static readonly Dictionary<string, Bitmap> cache = new Dictionary<string, Bitmap>();

        // O ícone no tamanho pedido: o PNG desse tamanho, ou o maior seguinte reduzido
        public static Bitmap Image(string name, int size)
        {
            string key = name + "_" + size;
            Bitmap b;
            if (cache.TryGetValue(key, out b)) return b;
            int source = sizes[sizes.Length - 1];
            foreach (int s in sizes) if (s >= size) { source = s; break; }
            using (System.IO.Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("icons." + name + "_" + source + ".png"))
            {
                if (stream == null) return null;
                using (Bitmap png = new Bitmap(stream))
                {
                    b = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    using (Graphics g = Graphics.FromImage(b))
                    {
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        g.DrawImage(png, 0, 0, size, size);
                    }
                }
            }
            cache[key] = b;
            return b;
        }

        static IntPtr IconHandle(string name, int size)
        {
            Bitmap b = Image(name, size);
            return b == null ? IntPtr.Zero : b.GetHicon();
        }

        // Ícone de uma aba do código: o do componente na árvore; senão, pelo tipo da janela
        public static string ForTab(string name, string kind)
        {
            string icon;
            if (byName.TryGetValue(name, out icon)) return icon;
            return string.Equals(kind, "UserForm", StringComparison.OrdinalIgnoreCase) ? "file_type_xaml" : "file_type_vba";
        }

        // ------------------------------------------------------------------

        const int TV_FIRST = 0x1100, TVM_GETIMAGELIST = TV_FIRST + 8, TVM_SETIMAGELIST = TV_FIRST + 9, TVM_GETNEXTITEM = TV_FIRST + 10,
            TVM_GETITEMW = TV_FIRST + 62, TVM_SETITEMW = TV_FIRST + 63;
        const int TVSIL_NORMAL = 0, TVGN_ROOT = 0, TVGN_NEXT = 1, TVGN_CHILD = 4;
        const uint TVIF_TEXT = 1, TVIF_IMAGE = 2, TVIF_STATE = 8, TVIF_HANDLE = 0x10, TVIF_SELECTEDIMAGE = 0x20, TVIS_EXPANDED = 0x20;
        const int WM_NOTIFY = 0x004E, TVN_ITEMEXPANDEDA = -415, TVN_ITEMEXPANDEDW = -455;
        const uint ILC_COLOR32 = 0x20, ILD_TRANSPARENT = 1;

        [StructLayout(LayoutKind.Sequential)]
        struct TVITEM
        {
            public uint mask; public IntPtr hItem; public uint state, stateMask; public IntPtr pszText; public int cchTextMax, iImage, iSelectedImage, cChildren; public IntPtr lParam;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct NMHDR { public IntPtr hwndFrom; public UIntPtr idFrom; public int code; }

        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int m, IntPtr w, IntPtr l);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] static extern IntPtr SendItem(IntPtr h, int m, IntPtr w, ref TVITEM l);
        [DllImport("comctl32.dll")] static extern IntPtr ImageList_Create(int cx, int cy, uint flags, int initial, int grow);
        [DllImport("comctl32.dll")] static extern bool ImageList_Destroy(IntPtr il);
        [DllImport("comctl32.dll")] static extern int ImageList_GetImageCount(IntPtr il);
        [DllImport("comctl32.dll")] static extern bool ImageList_GetIconSize(IntPtr il, out int cx, out int cy);
        [DllImport("comctl32.dll")] static extern IntPtr ImageList_GetIcon(IntPtr il, int i, uint flags);
        [DllImport("comctl32.dll")] static extern int ImageList_ReplaceIcon(IntPtr il, int i, IntPtr icon);
        [DllImport("user32.dll")] static extern bool DestroyIcon(IntPtr h);
    }
}
