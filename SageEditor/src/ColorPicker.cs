using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageEditor
{
    // O seletor de cor dos decoradores (ColorDecorators): área de saturação e brilho, faixa de
    // matiz, a cor antiga e a nova, e os campos #RRGGBB e R, G, B. Cada mudança chega ao código
    // logo em seguida (Changed, com um intervalo curto para não reescrever a linha a cada pixel);
    // Enter ou clicar fora fecha e mantém, Esc fecha e volta a cor original (Cancel). Com as
    // cores do tema do editor. Roda na thread do Excel, como o resto do SageEditor.
    sealed class ColorPicker : Form
    {
        static ColorPicker open;

        public static void Open(Point screen, int colorRef, Action<int> changed, Action cancel, IntPtr owner)
        {
            CloseOpen();
            open = new ColorPicker(colorRef, changed, cancel);
            open.Place(screen);
            open.Show(new OwnerWindow(owner));
            open.Activate();
            open.hex.Focus();
            open.hex.SelectAll();
        }

        public static void CloseOpen()
        {
            if (open != null && !open.IsDisposed) open.Finish(false);
            open = null;
        }

        sealed class OwnerWindow : IWin32Window
        {
            readonly IntPtr handle;
            public OwnerWindow(IntPtr h) { handle = h; }
            public IntPtr Handle { get { return handle; } }
        }

        readonly Action<int> changed;
        readonly Action cancel;
        readonly Color original;
        double hue, sat, val;            // 0..360, 0..1, 0..1
        Color color;
        Bitmap square;
        double squareHue = -1;
        bool updating, finished, dirty;
        int dragging;                    // 0 nada, 1 área, 2 matiz
        readonly Timer apply = new Timer { Interval = 60 };
        readonly float scale;
        readonly Color back, fore, border, field;
        readonly System.Windows.Forms.TextBox hex, red, green, blue;

        ColorPicker(int colorRef, Action<int> changed, Action cancel)
        {
            this.changed = changed; this.cancel = cancel;
            original = color = Color.FromArgb(colorRef & 0xFF, (colorRef >> 8) & 0xFF, (colorRef >> 16) & 0xFF);
            ToHsv(color, out hue, out sat, out val);

            Theme t = ThemeEngine.Effective;
            back = t != null ? FromRef(t.Window) : SystemColors.Window;
            fore = t != null ? FromRef(t.WindowText) : SystemColors.WindowText;
            border = Mix(back, fore, 0.3);
            field = Mix(back, fore, 0.08);

            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) scale = g.DpiX / 96f;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            BackColor = back;
            KeyPreview = true;
            ClientSize = new Size(S(236), S(222));
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            Font = new Font("Segoe UI", 9f);

            hex = Field(S(52), S(166), S(76));
            red = Field(S(52), S(194), S(42));
            green = Field(S(118), S(194), S(42));
            blue = Field(S(184), S(194), S(42));
            hex.TextChanged += delegate { FromHex(); };
            red.TextChanged += delegate { FromRgb(); };
            green.TextChanged += delegate { FromRgb(); };
            blue.TextChanged += delegate { FromRgb(); };
            apply.Tick += delegate { apply.Stop(); Push(); };
            ShowFields();
        }

        System.Windows.Forms.TextBox Field(int x, int y, int width)
        {
            System.Windows.Forms.TextBox box = new System.Windows.Forms.TextBox
            {
                BorderStyle = BorderStyle.FixedSingle, BackColor = field, ForeColor = fore,
                Location = new Point(x, y), Width = width, Font = new Font("Consolas", 9.5f),
            };
            box.GotFocus += delegate { box.SelectAll(); };
            Controls.Add(box);
            return box;
        }

        int S(int px) { return (int)Math.Round(px * scale); }

        Rectangle Area { get { return new Rectangle(S(10), S(10), S(184), S(140)); } }
        Rectangle Hue { get { return new Rectangle(S(204), S(10), S(22), S(140)); } }
        Rectangle Before { get { return new Rectangle(S(10), S(166), S(18), S(22)); } }
        Rectangle After { get { return new Rectangle(S(28), S(166), S(18), S(22)); } }

        void Place(Point screen)
        {
            Rectangle work = Screen.FromPoint(screen).WorkingArea;
            int x = Math.Max(work.Left, Math.Min(screen.X, work.Right - Width));
            int y = screen.Y + Height > work.Bottom ? screen.Y - Height - S(24) : screen.Y;
            Location = new Point(x, Math.Max(work.Top, y));
        }

        // ------------------------------------------------------------------
        // Desenho
        // ------------------------------------------------------------------

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(back);
            using (Pen p = new Pen(border)) g.DrawRectangle(p, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);

            // Saturação (horizontal) e brilho (vertical) na matiz atual
            Rectangle a = Area;
            if (square == null || squareHue != hue || square.Width != a.Width) BuildSquare(a.Size);
            g.DrawImageUnscaled(square, a.Location);
            using (Pen p = new Pen(border)) g.DrawRectangle(p, a.X - 1, a.Y - 1, a.Width + 1, a.Height + 1);
            int cx = a.Left + (int)Math.Round(sat * (a.Width - 1)), cy = a.Top + (int)Math.Round((1 - val) * (a.Height - 1));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (Pen outer = new Pen(Color.Black, 1.5f * scale)) g.DrawEllipse(outer, cx - S(6), cy - S(6), S(12), S(12));
            using (Pen inner = new Pen(Color.White, 1.5f * scale)) g.DrawEllipse(inner, cx - S(5), cy - S(5), S(10), S(10));
            g.SmoothingMode = SmoothingMode.Default;

            // Faixa de matiz
            Rectangle h = Hue;
            using (LinearGradientBrush b = new LinearGradientBrush(h, Color.Red, Color.Red, LinearGradientMode.Vertical))
            {
                ColorBlend blend = new ColorBlend(7);
                blend.Colors = new[] { Color.Red, Color.Magenta, Color.Blue, Color.Cyan, Color.Lime, Color.Yellow, Color.Red };
                blend.Positions = new[] { 0f, 1f / 6, 2f / 6, 3f / 6, 4f / 6, 5f / 6, 1f };
                b.InterpolationColors = blend;
                g.FillRectangle(b, h);
            }
            using (Pen p = new Pen(border)) g.DrawRectangle(p, h.X - 1, h.Y - 1, h.Width + 1, h.Height + 1);
            int hy = h.Top + (int)Math.Round((1 - hue / 360) * (h.Height - 1));
            using (Pen p = new Pen(fore, 2 * scale)) g.DrawRectangle(p, h.Left - S(2), hy - S(2), h.Width + S(4), S(4));

            // Cor original e a nova, e os rótulos
            using (SolidBrush b = new SolidBrush(original)) g.FillRectangle(b, Before);
            using (SolidBrush b = new SolidBrush(color)) g.FillRectangle(b, After);
            using (Pen p = new Pen(border)) g.DrawRectangle(p, Before.X, Before.Y, Before.Width + After.Width - 1, Before.Height - 1);
            TextFormatFlags f = TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding;
            Color muted = Mix(back, fore, 0.6);
            TextRenderer.DrawText(g, "R", Font, new Rectangle(S(30), S(194), S(18), hex.Height), muted, f);
            TextRenderer.DrawText(g, "G", Font, new Rectangle(S(96), S(194), S(18), hex.Height), muted, f);
            TextRenderer.DrawText(g, "B", Font, new Rectangle(S(162), S(194), S(18), hex.Height), muted, f);
        }

        void BuildSquare(Size size)
        {
            if (square != null) square.Dispose();
            square = new Bitmap(size.Width, size.Height, PixelFormat.Format32bppRgb);
            BitmapData data = square.LockBits(new Rectangle(Point.Empty, size), ImageLockMode.WriteOnly, PixelFormat.Format32bppRgb);
            int[] row = new int[size.Width];
            for (int y = 0; y < size.Height; y++)
            {
                double v = 1 - y / (double)Math.Max(1, size.Height - 1);
                for (int x = 0; x < size.Width; x++)
                    row[x] = FromHsv(hue, x / (double)Math.Max(1, size.Width - 1), v).ToArgb();
                Marshal.Copy(row, 0, data.Scan0 + y * data.Stride, size.Width);
            }
            square.UnlockBits(data);
            squareHue = hue;
        }

        // ------------------------------------------------------------------
        // Mouse e teclado
        // ------------------------------------------------------------------

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left) return;
            if (Area.Contains(e.Location)) dragging = 1;
            else if (Hue.Contains(e.Location)) dragging = 2;
            else if (Before.Contains(e.Location)) { SetColor(original, true); return; }
            else return;
            Capture = true;
            Track(e.Location);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (dragging != 0) Track(e.Location);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            dragging = 0;
            Capture = false;
        }

        void Track(Point p)
        {
            if (dragging == 1)
            {
                Rectangle a = Area;
                sat = Clamp((p.X - a.Left) / (double)Math.Max(1, a.Width - 1));
                val = 1 - Clamp((p.Y - a.Top) / (double)Math.Max(1, a.Height - 1));
            }
            else
            {
                Rectangle h = Hue;
                hue = (1 - Clamp((p.Y - h.Top) / (double)Math.Max(1, h.Height - 1))) * 360;
            }
            color = FromHsv(hue, sat, val);
            ShowFields();
            Queue();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Escape) { e.Handled = e.SuppressKeyPress = true; Finish(true); }
            else if (e.KeyCode == Keys.Return) { e.Handled = e.SuppressKeyPress = true; Finish(false); }
        }

        // Clicar fora: fecha e mantém a cor
        protected override void OnDeactivate(EventArgs e)
        {
            base.OnDeactivate(e);
            if (!finished) BeginInvoke((MethodInvoker)delegate { Finish(false); });
        }

        void Finish(bool revert)
        {
            if (finished) return;
            finished = true;
            apply.Stop();
            try
            {
                if (revert) cancel();
                else if (dirty) Push();
            }
            catch (Exception ex) { Log.Error(ex); }
            if (open == this) open = null;
            Close();
        }

        // ------------------------------------------------------------------
        // Campos e cor
        // ------------------------------------------------------------------

        void SetColor(Color c, bool fields)
        {
            color = c;
            double h;
            ToHsv(c, out h, out sat, out val);
            if (sat > 0 && val > 0) hue = h; // cinza: a matiz fica onde estava
            if (fields) ShowFields();
            Invalidate();
            Queue();
        }

        void ShowFields()
        {
            updating = true;
            hex.Text = string.Format("#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B);
            red.Text = color.R.ToString(CultureInfo.InvariantCulture);
            green.Text = color.G.ToString(CultureInfo.InvariantCulture);
            blue.Text = color.B.ToString(CultureInfo.InvariantCulture);
            updating = false;
            Invalidate();
        }

        void FromHex()
        {
            if (updating) return;
            string s = hex.Text.Trim().TrimStart('#');
            int v;
            if (s.Length != 6 || !int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v)) return;
            Color c = Color.FromArgb((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF);
            updating = true;
            red.Text = c.R.ToString(CultureInfo.InvariantCulture);
            green.Text = c.G.ToString(CultureInfo.InvariantCulture);
            blue.Text = c.B.ToString(CultureInfo.InvariantCulture);
            updating = false;
            SetColor(c, false);
        }

        void FromRgb()
        {
            if (updating) return;
            int r, g, b;
            if (!Part(red.Text, out r) || !Part(green.Text, out g) || !Part(blue.Text, out b)) return;
            Color c = Color.FromArgb(r, g, b);
            updating = true;
            hex.Text = string.Format("#{0:X2}{1:X2}{2:X2}", r, g, b);
            updating = false;
            SetColor(c, false);
        }

        static bool Part(string s, out int v)
        {
            return int.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out v) && v >= 0 && v <= 255;
        }

        void Queue()
        {
            dirty = true;
            apply.Stop();
            apply.Start();
        }

        void Push()
        {
            dirty = false;
            changed(color.R | (color.G << 8) | (color.B << 16));
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                apply.Dispose();
                if (square != null) square.Dispose();
            }
            base.Dispose(disposing);
        }

        // ------------------------------------------------------------------

        static double Clamp(double v) { return Math.Max(0, Math.Min(1, v)); }

        static Color FromRef(int c) { return Color.FromArgb(c & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF); }

        static Color Mix(Color a, Color b, double k)
        {
            return Color.FromArgb((int)(a.R + (b.R - a.R) * k), (int)(a.G + (b.G - a.G) * k), (int)(a.B + (b.B - a.B) * k));
        }

        static void ToHsv(Color c, out double h, out double s, out double v)
        {
            double r = c.R / 255.0, g = c.G / 255.0, b = c.B / 255.0;
            double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b)), d = max - min;
            v = max;
            s = max == 0 ? 0 : d / max;
            if (d == 0) h = 0;
            else if (max == r) h = 60 * (((g - b) / d) % 6);
            else if (max == g) h = 60 * ((b - r) / d + 2);
            else h = 60 * ((r - g) / d + 4);
            if (h < 0) h += 360;
        }

        static Color FromHsv(double h, double s, double v)
        {
            h = ((h % 360) + 360) % 360;
            double c = v * s, x = c * (1 - Math.Abs((h / 60) % 2 - 1)), m = v - c;
            double r, g, b;
            if (h < 60) { r = c; g = x; b = 0; }
            else if (h < 120) { r = x; g = c; b = 0; }
            else if (h < 180) { r = 0; g = c; b = x; }
            else if (h < 240) { r = 0; g = x; b = c; }
            else if (h < 300) { r = x; g = 0; b = c; }
            else { r = c; g = 0; b = x; }
            return Color.FromArgb((int)Math.Round((r + m) * 255), (int)Math.Round((g + m) * 255), (int)Math.Round((b + m) * 255));
        }
    }
}
