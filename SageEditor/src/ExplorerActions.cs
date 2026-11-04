using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageEditor
{
    // Botões Novo Formulário, Novo Módulo e Nova Classe na linha do projeto ativo do explorador,
    // à direita, como os de Novo Arquivo do VS Code. A árvore é do VBE: os botões são desenhados
    // por cima dela depois de cada WM_PAINT, e os cliques neles não chegam à árvore
    // (UnifiedExplorer.TreeProc). O projeto ativo é o do item selecionado (o mesmo do
    // VBE.ActiveVBProject).
    static class ExplorerActions
    {
        const int vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2, vbext_ct_MSForm = 3;
        static readonly int[] kinds = { vbext_ct_MSForm, vbext_ct_StdModule, vbext_ct_ClassModule };
        static readonly string[] icons = { "file_type_xaml", "file_type_vba", "vba_class" };

        static int hot = -1, pressed = -1;
        static Native.RECT lastRow;
        static bool tracking;
        static ToolTip tip;
        static readonly Dictionary<string, Bitmap> cache = new Dictionary<string, Bitmap>();

        static string Text(int i) { return i == 0 ? Strings.NewForm : i == 1 ? Strings.NewModule : Strings.NewClass; }

        // ------------------------------------------------------------------
        // Posição: a linha do projeto do item selecionado, botões encostados à direita
        // ------------------------------------------------------------------

        static bool Row(IntPtr tree, out Native.RECT row)
        {
            row = new Native.RECT();
            IntPtr item = Native.SendMessage(tree, TVM_GETNEXTITEM, (IntPtr)TVGN_CARET, IntPtr.Zero);
            if (item == IntPtr.Zero) item = Native.SendMessage(tree, TVM_GETNEXTITEM, (IntPtr)TVGN_ROOT, IntPtr.Zero);
            if (item == IntPtr.Zero) return false;
            for (IntPtr up = item; up != IntPtr.Zero; up = Native.SendMessage(tree, TVM_GETNEXTITEM, (IntPtr)TVGN_PARENT, up)) item = up;
            // TVM_GETITEMRECT: o item vai no começo do RECT; wParam 0 = a linha inteira
            IntPtr buffer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Native.RECT)));
            try
            {
                Marshal.WriteIntPtr(buffer, item);
                if (Native.SendMessage(tree, TVM_GETITEMRECT, IntPtr.Zero, buffer) == IntPtr.Zero) return false; // fora da vista
                row = (Native.RECT)Marshal.PtrToStructure(buffer, typeof(Native.RECT));
            }
            finally { Marshal.FreeHGlobal(buffer); }
            Native.RECT client;
            Native.GetClientRect(tree, out client);
            row.Left = 0;
            row.Right = client.Right;
            return row.Bottom > row.Top;
        }

        static Rectangle Button(Native.RECT row, int i)
        {
            int size = row.Bottom - row.Top, gap = Math.Max(10, size / 2); // espaço entre os botões
            int x = row.Right - 4 - (kinds.Length - i) * (size + gap) + gap;
            return new Rectangle(x, row.Top, size, size);
        }

        static int HitTest(IntPtr tree, IntPtr lParam)
        {
            Native.RECT row;
            if (!Row(tree, out row)) return -1;
            Point p = new Point((short)((long)lParam & 0xFFFF), (short)(((long)lParam >> 16) & 0xFFFF));
            for (int i = 0; i < kinds.Length; i++) if (Button(row, i).Contains(p)) return i;
            return -1;
        }

        // ------------------------------------------------------------------
        // Desenho
        // ------------------------------------------------------------------

        // Depois do WM_PAINT da árvore
        public static void Paint(IntPtr tree)
        {
            Native.RECT row;
            bool visible = Row(tree, out row);
            // Outro projeto ativo (ou a linha rolou): a linha antiga é redesenhada sem os botões
            if (lastRow.Bottom > lastRow.Top && (!visible || row.Top != lastRow.Top || row.Right != lastRow.Right))
            {
                Native.RECT old = lastRow;
                lastRow = new Native.RECT();
                InvalidateRect(tree, ref old, true);
            }
            if (!visible) return;
            lastRow = row;

            Theme t = Theme.Find(Settings.ColorTheme);
            int bk = (int)Native.SendMessage(tree, TVM_GETBKCOLOR, IntPtr.Zero, IntPtr.Zero);
            Color back = bk == -1 ? SystemColors.Window : Color.FromArgb(bk & 0xFF, (bk >> 8) & 0xFF, (bk >> 16) & 0xFF);
            Color fore = t.IsDefault ? SystemColors.WindowText : t.Foreground;
            Color hover = t.IsDefault ? SystemColors.ControlLight : t.Hover;

            Rectangle first = Button(row, 0);
            int size = first.Height;
            Rectangle area = new Rectangle(first.Left - 4, row.Top, row.Right - first.Left + 4, size);
            if (area.Width <= 0 || size <= 0) return;
            IntPtr dc = Native.GetDC(tree);
            if (dc == IntPtr.Zero) return;
            try
            {
                using (Bitmap strip = new Bitmap(area.Width, area.Height))
                {
                    using (Graphics g = Graphics.FromImage(strip))
                    {
                        g.Clear(back); // o fim de um nome comprido fica por baixo dos botões
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        for (int i = 0; i < kinds.Length; i++)
                        {
                            Rectangle b = Button(row, i);
                            b.Offset(-area.Left, -area.Top);
                            if (i == hot)
                                using (SolidBrush br = new SolidBrush(hover)) g.FillRectangle(br, Rectangle.Inflate(b, -1, -1));
                            Bitmap image = Icon(icons[i], IconSize(tree, size), i == hot ? hover : back, fore);
                            if (image != null) g.DrawImage(image, b.Left + (b.Width - image.Width) / 2, b.Top + (b.Height - image.Height) / 2);
                        }
                    }
                    using (Graphics target = Graphics.FromHdc(dc)) target.DrawImageUnscaled(strip, area.Left, area.Top);
                }
            }
            finally { Native.ReleaseDC(tree, dc); }
        }

        // O tamanho dos ícones da própria árvore (a lista de imagens, que acompanha o DPI), sem
        // passar da altura da linha
        static int IconSize(IntPtr tree, int row)
        {
            int cx, cy;
            IntPtr list = Native.SendMessage(tree, TVM_GETIMAGELIST, (IntPtr)TVSIL_NORMAL, IntPtr.Zero);
            if (list == IntPtr.Zero || !ImageList_GetIconSize(list, out cx, out cy) || cx <= 0) cx = 16;
            return Math.Min(cx, row);
        }

        // O ícone do Sage com um "+" no canto de baixo à direita, num círculo da cor do fundo
        static Bitmap Icon(string name, int size, Color back, Color fore)
        {
            string key = name + "_" + size + "_" + back.ToArgb() + "_" + fore.ToArgb();
            Bitmap b;
            if (cache.TryGetValue(key, out b)) return b;
            Bitmap source = ProjectIcons.Image(name, size);
            if (source == null) return null;
            b = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(b))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.DrawImage(source, 0, 0, size, size);
                float r = size * 0.3f, cx = size - r, cy = size - r;
                using (SolidBrush br = new SolidBrush(back)) g.FillEllipse(br, cx - r, cy - r, 2 * r, 2 * r);
                float arm = r * 0.7f;
                using (Pen pen = new Pen(fore, Math.Max(1.5f, size / 10f)))
                {
                    g.DrawLine(pen, cx - arm, cy, cx + arm, cy);
                    g.DrawLine(pen, cx, cy - arm, cx, cy + arm);
                }
            }
            cache[key] = b;
            return b;
        }

        // ------------------------------------------------------------------
        // Mouse
        // ------------------------------------------------------------------

        public static void MouseMove(IntPtr tree, IntPtr lParam)
        {
            if (!tracking)
            {
                TRACKMOUSEEVENT tme = new TRACKMOUSEEVENT { cbSize = Marshal.SizeOf(typeof(TRACKMOUSEEVENT)), dwFlags = TME_LEAVE, hwndTrack = tree };
                tracking = TrackMouseEvent(ref tme);
            }
            SetHot(tree, HitTest(tree, lParam));
        }

        public static void MouseLeave(IntPtr tree)
        {
            tracking = false;
            pressed = -1;
            SetHot(tree, -1);
        }

        static void SetHot(IntPtr tree, int i)
        {
            if (i == hot) return;
            hot = i;
            Paint(tree);
            // O ToolTip do WinForms só aceita um Control: vai preso à faixa das Propriedades, que é
            // do Sage e fica na mesma janela, com a posição convertida da árvore para ela
            Control anchor = UnifiedExplorer.HeaderControl;
            if (anchor == null || !anchor.IsHandleCreated) return;
            if (tip == null) tip = new ToolTip { ShowAlways = true };
            Native.RECT row;
            if (i < 0 || !Row(tree, out row)) { tip.Hide(anchor); return; }
            Rectangle b = Button(row, i);
            POINT p = new POINT { X = b.Left, Y = b.Bottom + 2 };
            ClientToScreen(tree, ref p);
            Point local = anchor.PointToClient(new Point(p.X, p.Y));
            tip.Show(Text(i), anchor, local.X, local.Y);
        }

        static void HideTip()
        {
            Control anchor = UnifiedExplorer.HeaderControl;
            if (tip != null && anchor != null && anchor.IsHandleCreated) tip.Hide(anchor);
        }

        // true: o clique foi num botão e não vai para a árvore
        public static bool MouseDown(IntPtr tree, IntPtr lParam)
        {
            pressed = HitTest(tree, lParam);
            return pressed >= 0;
        }

        public static bool MouseUp(IntPtr tree, IntPtr lParam)
        {
            if (pressed < 0) return false;
            int i = pressed;
            pressed = -1;
            if (HitTest(tree, lParam) == i) Create(kinds[i]);
            return true;
        }

        // Novo componente no projeto ativo, aberto no código (ou no designer, o formulário)
        static void Create(int kind)
        {
            HideTip();
            try
            {
                dynamic project = UnifiedExplorer.Vbe.ActiveVBProject;
                if (project == null) return;
                dynamic component = project.VBComponents.Add(kind);
                component.Activate();
            }
            catch (Exception ex)
            {
                // Projeto protegido ou código em execução: a mensagem do VBE
                MessageBox.Show(ex.InnerException != null ? ex.InnerException.Message : ex.Message, "Sage",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ------------------------------------------------------------------

        const int TV_FIRST = 0x1100, TVM_GETITEMRECT = TV_FIRST + 4, TVM_GETIMAGELIST = TV_FIRST + 8, TVM_GETNEXTITEM = TV_FIRST + 10,
            TVM_GETBKCOLOR = TV_FIRST + 31;
        const int TVSIL_NORMAL = 0, TVGN_ROOT = 0, TVGN_PARENT = 3, TVGN_CARET = 9;
        const uint TME_LEAVE = 2;

        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr hwnd, ref POINT p);

        [StructLayout(LayoutKind.Sequential)]
        struct TRACKMOUSEEVENT { public int cbSize; public uint dwFlags; public IntPtr hwndTrack; public uint dwHoverTime; }

        [DllImport("user32.dll")] static extern bool TrackMouseEvent(ref TRACKMOUSEEVENT tme);
        [DllImport("comctl32.dll")] static extern bool ImageList_GetIconSize(IntPtr il, out int cx, out int cy);
        [DllImport("user32.dll")] static extern bool InvalidateRect(IntPtr hwnd, ref Native.RECT rect, bool erase);
    }
}
