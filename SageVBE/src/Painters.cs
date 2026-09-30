using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageVBE
{
    // Desenhos próprios para o que o user32 pinta com cores internas (sem passar
    // pelo GetSysColor desviado): bordas 3D, abas das Propriedades e os botões de
    // modo de exibição do editor.
    static class Painters
    {
        // --- DrawEdge / DrawFrameControl (bordas 3D e botões de legenda) ---

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        delegate bool EdgeProc(IntPtr hdc, IntPtr rect, uint a, uint b);

        static EdgeProc origEdge, origFrame;
        static readonly EdgeProc hookEdge = HookDrawEdge;
        static readonly EdgeProc hookFrame = HookDrawFrameControl;

        public static void HookEdges(IntPtr user32, System.Collections.Generic.Dictionary<string, IntPtr> hooks)
        {
            origEdge = (EdgeProc)Marshal.GetDelegateForFunctionPointer(Native.GetProcAddress(user32, "DrawEdge"), typeof(EdgeProc));
            origFrame = (EdgeProc)Marshal.GetDelegateForFunctionPointer(Native.GetProcAddress(user32, "DrawFrameControl"), typeof(EdgeProc));
            hooks["DrawEdge"] = Marshal.GetFunctionPointerForDelegate(hookEdge);
            hooks["DrawFrameControl"] = Marshal.GetFunctionPointerForDelegate(hookFrame);
        }

        const uint BF_LEFT = 0x1, BF_TOP = 0x2, BF_RIGHT = 0x4, BF_BOTTOM = 0x8, BF_DIAGONAL = 0x10,
            BF_MIDDLE = 0x800, BF_ADJUST = 0x2000;

        // Borda lisa de 1 px por nível (externo/interno), na cor de borda do tema
        static bool HookDrawEdge(IntPtr hdc, IntPtr rectPtr, uint edge, uint flags)
        {
            Theme t = ThemeEngine.DarkPainting;
            if (t == null || (flags & BF_DIAGONAL) != 0) return origEdge(hdc, rectPtr, edge, flags);

            Native.RECT r = (Native.RECT)Marshal.PtrToStructure(rectPtr, typeof(Native.RECT));
            int levels = ((edge & 0x3) != 0 ? 1 : 0) + ((edge & 0xC) != 0 ? 1 : 0);
            IntPtr line = t.Brush(t.FrameBorder);
            for (int i = 0; i < levels; i++)
            {
                if ((flags & BF_LEFT) != 0) { Fill(hdc, r.Left, r.Top, r.Left + 1, r.Bottom, line); r.Left++; }
                if ((flags & BF_TOP) != 0) { Fill(hdc, r.Left, r.Top, r.Right, r.Top + 1, line); r.Top++; }
                if ((flags & BF_RIGHT) != 0) { Fill(hdc, r.Right - 1, r.Top, r.Right, r.Bottom, line); r.Right--; }
                if ((flags & BF_BOTTOM) != 0) { Fill(hdc, r.Left, r.Bottom - 1, r.Right, r.Bottom, line); r.Bottom--; }
            }
            if ((flags & BF_MIDDLE) != 0) Fill(hdc, r.Left, r.Top, r.Right, r.Bottom, t.Brush(t.Face));
            if ((flags & BF_ADJUST) != 0) Marshal.StructureToPtr(r, rectPtr, false);
            return true;
        }

        const uint DFC_CAPTION = 1, DFCS_PUSHED = 0x200, DFCS_INACTIVE = 0x100;

        // Botões de legenda (fechar, minimizar, maximizar, restaurar) das janelas encaixadas
        static bool HookDrawFrameControl(IntPtr hdc, IntPtr rectPtr, uint type, uint state)
        {
            Theme t = ThemeEngine.DarkPainting;
            if (t == null || type != DFC_CAPTION) return origFrame(hdc, rectPtr, type, state);

            Native.RECT r = (Native.RECT)Marshal.PtrToStructure(rectPtr, typeof(Native.RECT));
            bool pushed = (state & DFCS_PUSHED) != 0;
            Fill(hdc, r.Left, r.Top, r.Right, r.Bottom, t.Brush(pushed ? t.FrameBorder : t.Face));

            using (Graphics g = Graphics.FromHdc(hdc))
            using (Pen pen = new Pen((state & DFCS_INACTIVE) != 0 ? t.Muted : t.Foreground, 1))
            {
                int w = r.Right - r.Left, h = r.Bottom - r.Top, s = Math.Max(4, Math.Min(w, h) / 2 - 1);
                int x = r.Left + (w - s) / 2, y = r.Top + (h - s) / 2;
                switch (state & 0xFF)
                {
                    case 0: // fechar
                        g.DrawLine(pen, x, y, x + s, y + s);
                        g.DrawLine(pen, x + s, y, x, y + s);
                        break;
                    case 1: // minimizar
                        g.DrawLine(pen, x, y + s, x + s, y + s);
                        break;
                    case 2: // maximizar
                        g.DrawRectangle(pen, x, y, s, s);
                        break;
                    case 3: // restaurar
                        g.DrawRectangle(pen, x, y + 2, s - 2, s - 2);
                        g.DrawLine(pen, x + 2, y, x + s, y);
                        g.DrawLine(pen, x + s, y, x + s, y + s - 2);
                        break;
                    default:
                        return origFrame(hdc, rectPtr, type, state);
                }
            }
            return true;
        }

        static void Fill(IntPtr hdc, int left, int top, int right, int bottom, IntPtr brush)
        {
            Native.RECT r = new Native.RECT { Left = left, Top = top, Right = right, Bottom = bottom };
            Native.FillRect(hdc, ref r, brush);
        }

        const int WS_BORDER = 0x00800000;
        const int WS_EX_CLIENTEDGE = 0x00000200, WS_EX_STATICEDGE = 0x00020000, WS_EX_WINDOWEDGE = 0x00000100;

        // Borda de janelas filhas (WS_BORDER / WS_EX_CLIENTEDGE...), repintada na cor do tema
        public static void PaintBorder(IntPtr hwnd, Theme t)
        {
            int style = Native.GetWindowLong(hwnd, Native.GWL_STYLE);
            if ((style & Native.WS_CHILD) == 0) return;
            int ex = Native.GetWindowLong(hwnd, Native.GWL_EXSTYLE);

            int width = 0;
            if ((ex & WS_EX_CLIENTEDGE) != 0) width += 2;
            if ((ex & WS_EX_STATICEDGE) != 0) width += 1;
            if ((ex & WS_EX_WINDOWEDGE) != 0) width += 2;
            if ((style & WS_BORDER) != 0) width += 1;
            if (width == 0) return;

            Native.RECT r;
            Native.GetWindowRect(hwnd, out r);
            Native.RECT box = new Native.RECT { Left = 0, Top = 0, Right = r.Right - r.Left, Bottom = r.Bottom - r.Top };
            IntPtr dc = Native.GetWindowDC(hwnd);
            try
            {
                IntPtr outer = t.Brush(t.FrameBorder), inner = t.Brush(t.Face);
                for (int i = 0; i < width; i++)
                {
                    Native.FrameRect(dc, ref box, i == 0 ? outer : inner);
                    box.Left++; box.Top++; box.Right--; box.Bottom--;
                }
            }
            finally { Native.ReleaseDC(hwnd, dc); }
        }

        // Abas (SysTabControl32) no estilo do VS Code: fundo liso, aba ativa com texto
        // claro e sublinhado na cor de destaque.
        public static void PaintTabs(IntPtr hwnd, Theme t)
        {
            Native.PAINTSTRUCT ps;
            IntPtr hdc = Native.BeginPaint(hwnd, out ps);
            try
            {
                Native.RECT client;
                Native.GetClientRect(hwnd, out client);
                using (Graphics g = Graphics.FromHdc(hdc))
                {
                    Color face = FromRef(t.Face), window = FromRef(t.Window);
                    Color text = t.Foreground, muted = t.Muted;
                    g.Clear(face);

                    IntPtr hfont = Native.SendMessage(hwnd, Native.WM_GETFONT, IntPtr.Zero, IntPtr.Zero);
                    Font font = hfont != IntPtr.Zero ? Font.FromHfont(hfont) : SystemFonts.MessageBoxFont;
                    try
                    {
                        int count = (int)Native.SendMessage(hwnd, Native.TCM_GETITEMCOUNT, IntPtr.Zero, IntPtr.Zero);
                        int selected = (int)Native.SendMessage(hwnd, Native.TCM_GETCURSEL, IntPtr.Zero, IntPtr.Zero);
                        int bottom = 0;
                        for (int i = 0; i < count; i++)
                        {
                            Native.RECT ir;
                            Native.SendMessage(hwnd, Native.TCM_GETITEMRECT, (IntPtr)i, out ir);
                            Rectangle rect = Rectangle.FromLTRB(ir.Left, ir.Top, ir.Right, ir.Bottom);
                            bottom = Math.Max(bottom, rect.Bottom);
                            bool on = i == selected;
                            if (on)
                            {
                                using (SolidBrush b = new SolidBrush(window)) g.FillRectangle(b, rect);
                                using (SolidBrush b = new SolidBrush(t.Accent)) g.FillRectangle(b, rect.Left, rect.Bottom - 2, rect.Width, 2);
                            }
                            TextRenderer.DrawText(g, ItemText(hwnd, i), font, rect, on ? text : muted,
                                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
                        }
                        // área abaixo das abas (onde fica a lista de propriedades)
                        using (SolidBrush b = new SolidBrush(window))
                            g.FillRectangle(b, 0, bottom, client.Right, client.Bottom - bottom);
                        using (Pen p = new Pen(FromRef(t.FrameBorder)))
                            g.DrawLine(p, 0, bottom, client.Right, bottom);
                    }
                    finally
                    {
                        if (hfont != IntPtr.Zero) font.Dispose();
                    }
                }
            }
            finally { Native.EndPaint(hwnd, ref ps); }
        }

        static string ItemText(IntPtr hwnd, int index)
        {
            IntPtr buffer = Marshal.AllocHGlobal(512);
            try
            {
                Native.TCITEM item = new Native.TCITEM();
                item.mask = Native.TCIF_TEXT;
                item.pszText = buffer;
                item.cchTextMax = 255;
                Native.SendMessage(hwnd, Native.TCM_GETITEMW, (IntPtr)index, ref item);
                return (Marshal.PtrToStringUni(buffer) ?? "").Replace("&", "");
            }
            finally { Marshal.FreeHGlobal(buffer); }
        }

        // Depois do desenho original, troca as cores clássicas (branco, cinzas, preto)
        // pelas do tema. Usado em controles pequenos com bitmaps (ObtbarWndClass).
        public static void Recolor(IntPtr hwnd, Theme t)
        {
            Native.RECT r;
            Native.GetClientRect(hwnd, out r);
            int w = r.Right, h = r.Bottom;
            if (w <= 0 || h <= 0 || w * h > 40000) return;

            IntPtr dc = Native.GetDC(hwnd);
            try
            {
                using (Bitmap bmp = new Bitmap(w, h, PixelFormat.Format32bppRgb))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        IntPtr mem = g.GetHdc();
                        Native.BitBlt(mem, 0, 0, w, h, dc, 0, 0, Native.SRCCOPY);
                        g.ReleaseHdc(mem);
                    }

                    int face = Argb(t.Face), fg = Argb(t.WindowText), border = Argb(t.FrameBorder);
                    int light = Argb(t.IsDark ? Theme.Ref("#3C3C3C") : Theme.Ref("#C8C8C8"));
                    BitmapData data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadWrite, PixelFormat.Format32bppRgb);
                    int[] px = new int[w * h];
                    Marshal.Copy(data.Scan0, px, 0, px.Length);
                    for (int i = 0; i < px.Length; i++)
                    {
                        int c = px[i] & 0xFFFFFF;
                        if (c == 0xFFFFFF) px[i] = light;
                        else if (c == 0xC0C0C0 || c == 0xD4D0C8 || c == 0xF0F0F0 || c == 0xE3E3E3) px[i] = face;
                        else if (c == 0x808080 || c == 0xA0A0A0 || c == 0x696969 || c == 0xA0A0A4) px[i] = border;
                        else if (c == 0x000000) px[i] = fg;
                    }
                    Marshal.Copy(px, 0, data.Scan0, px.Length);
                    bmp.UnlockBits(data);

                    using (Graphics g = Graphics.FromHdc(dc))
                        g.DrawImageUnscaled(bmp, 0, 0);
                }
            }
            finally { Native.ReleaseDC(hwnd, dc); }
        }

        // Troca pixels de cores exatas (0xRRGGBB) pela cor de fundo do editor
        public static void ReplaceColors(IntPtr hwnd, Theme t, params int[] rgb)
        {
            Native.RECT r;
            Native.GetClientRect(hwnd, out r);
            int w = r.Right, h = r.Bottom;
            if (w <= 0 || h <= 0 || w * h > 400000) return;

            IntPtr dc = Native.GetDC(hwnd);
            try
            {
                using (Bitmap bmp = new Bitmap(w, h, PixelFormat.Format32bppRgb))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        IntPtr mem = g.GetHdc();
                        Native.BitBlt(mem, 0, 0, w, h, dc, 0, 0, Native.SRCCOPY);
                        g.ReleaseHdc(mem);
                    }
                    BitmapData data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadWrite, PixelFormat.Format32bppRgb);
                    int[] px = new int[w * h];
                    Marshal.Copy(data.Scan0, px, 0, px.Length);
                    int target = Argb(t.Window);
                    bool changed = false;
                    for (int i = 0; i < px.Length; i++)
                    {
                        int c = px[i] & 0xFFFFFF;
                        for (int k = 0; k < rgb.Length; k++)
                            if (c == rgb[k]) { px[i] = target; changed = true; break; }
                    }
                    if (!changed) { bmp.UnlockBits(data); return; }
                    Marshal.Copy(px, 0, data.Scan0, px.Length);
                    bmp.UnlockBits(data);
                    using (Graphics g = Graphics.FromHdc(dc))
                        g.DrawImageUnscaled(bmp, 0, 0);
                }
            }
            finally { Native.ReleaseDC(hwnd, dc); }
        }

        // COLORREF (0x00BBGGRR) -> pixel 0xFFRRGGBB
        static int Argb(int colorRef)
        {
            return unchecked((int)0xFF000000) | ((colorRef & 0xFF) << 16) | (colorRef & 0xFF00) | ((colorRef >> 16) & 0xFF);
        }

        static Color FromRef(int colorRef)
        {
            return Color.FromArgb(colorRef & 0xFF, (colorRef >> 8) & 0xFF, (colorRef >> 16) & 0xFF);
        }
    }
}
