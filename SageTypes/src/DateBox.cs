using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Media;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageTypes
{
    // DateBox: caixa de data para UserForms, com calendário. Fica na Caixa de ferramentas
    // (Sage DateBox) e é usada como uma caixa de texto que só aceita datas:
    //
    //   dtEntrega.Value = Date                   ' Date ou Empty (sem data)
    //   dtEntrega.Format = "dd/mmm/yyyy"         ' formato do Format do VBA (padrão: data abreviada do Windows)
    //   dtEntrega.MinDate = Date                 ' limites (Empty: sem limite)
    //
    //   Private Sub dtEntrega_Change()
    //       lblPrazo.Caption = dtEntrega.Value - Date & " dias"
    //   End Sub
    //
    // Digitação livre, conferida ao sair da caixa ou no Enter: "12/10/2026", "12102026", "121026",
    // "1210" (ano atual), "12" (mês atual). Data inválida ou fora dos limites: volta a anterior.
    // Enquanto a caixa tem o foco, a data aparece como na digitação (dd/mm/aaaa); fora dela, no
    // Format. O botão à direita, F4 ou Alt+Seta para baixo abrem o calendário; Seta para cima/baixo
    // mudam um dia. O calendário não tira o foco do form (como a lista de uma ComboBox): as
    // setas, Page Up/Down e Enter funcionam nele enquanto está aberto; o clique no título passa
    // para meses e anos.
    //
    // Interface dual: acrescente membros só no fim, com o próximo DispId (ver StringS). As
    // propriedades gravadas no arquivo usam os nomes de SaveProperties: não os mude.
    [ComVisible(true), Guid("6E1F0B7C-3A52-4C8D-9E46-2B7D5A1C8F93"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _DateBox
    {
        [DispId(0)] object Value { get; set; }
        [DispId(1000), PropertyLet("Value")] void LetValue(object Value);
        [DispId(1)] string Text { get; set; }
        [DispId(2)] string Format { get; set; }
        [DispId(3)] object MinDate { get; set; }
        [DispId(1003), PropertyLet("MinDate")] void LetMinDate(object Value);
        [DispId(4)] object MaxDate { get; set; }
        [DispId(1004), PropertyLet("MaxDate")] void LetMaxDate(object Value);
        [DispId(5)] string Placeholder { get; set; }
        [DispId(6)] bool Locked { get; set; }
        [DispId(-514)] bool Enabled { get; set; }
        [DispId(-501), ComAliasName("stdole.OLE_COLOR")] int BackColor { get; set; }
        [DispId(-513), ComAliasName("stdole.OLE_COLOR")] int ForeColor { get; set; }
        [DispId(7), ComAliasName("stdole.OLE_COLOR")] int BorderColor { get; set; }
        [DispId(8)] sgBorderStyle BorderStyle { get; set; }
        [DispId(9)] sgTextAlign TextAlign { get; set; }
        [DispId(10)] string FontName { get; set; }
        [DispId(11)] double FontSize { get; set; }
        [DispId(12)] bool FontBold { get; set; }
        [DispId(13)] bool FontItalic { get; set; }
        [DispId(14)] void DropDown();
    }

    // Eventos (Private Sub dtEntrega_Change()...)
    [ComVisible(true), Guid("A3C95D27-8E14-4B6F-B0D2-7F48E1C63A5B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface __DateBox
    {
        [DispId(1)] void Change();
        [DispId(2)] void DropButtonClick();
        [DispId(3)] void KeyDown(ref short KeyCode, short Shift);
        [DispId(4)] void KeyPress(ref short KeyAscii);
    }

    public delegate void DateBoxEventHandler();
    public delegate void DateBoxKeyHandler(ref short KeyCode, short Shift);
    public delegate void DateBoxKeyPressHandler(ref short KeyAscii);

    [ComVisible(true), Guid("D41B7E05-2C9A-4F63-8B1E-5A0C3D9F7E28"), ProgId("Sage.DateBox")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_DateBox)), ComSourceInterfaces(typeof(__DateBox))]
    public sealed class DateBox : OleControl, _DateBox
    {
        public event DateBoxEventHandler Change;
        public event DateBoxEventHandler DropButtonClick;
        public event DateBoxKeyHandler KeyDown;
        public event DateBoxKeyPressHandler KeyPress;

        const int SystemWindow = unchecked((int)0x80000005), SystemWindowText = unchecked((int)0x80000008);
        const int DefaultBorder = 0xB3ADAB;
        internal static readonly Color Accent = Color.FromArgb(0, 120, 212);

        DateTime? value, minDate, maxDate;
        string format, placeholder, fontName;
        bool locked, enabled, fontBold, fontItalic;
        sgBorderStyle borderStyle;
        sgTextAlign textAlign;
        int backColor, foreColor, borderColor;
        double fontSize;

        DateEdit edit;
        Font font;
        CalendarPopup popup;
        bool buttonHover;

        public DateBox() : base(90, 18)
        {
            InitProperties();
        }

        protected override string TypeName { get { return "Sage DateBox"; } }

        protected override void InitProperties()
        {
            value = minDate = maxDate = null;
            format = ""; placeholder = ""; fontName = "Segoe UI"; fontSize = 9;
            locked = false; enabled = true; fontBold = false; fontItalic = false;
            borderStyle = sgBorderStyle.sgBorderStyleSingle;
            textAlign = sgTextAlign.sgTextAlignLeft;
            backColor = SystemWindow; foreColor = SystemWindowText; borderColor = DefaultBorder;
            ResetFont();
            if (edit != null) Apply();
        }

        // Datas gravadas como texto ISO ("" = sem data)
        protected override void SaveProperties(IDictionary<string, object> p)
        {
            p["Value"] = Iso(value);
            p["Format"] = format;
            p["MinDate"] = Iso(minDate);
            p["MaxDate"] = Iso(maxDate);
            p["Placeholder"] = placeholder;
            p["Locked"] = locked;
            p["Enabled"] = enabled;
            p["BackColor"] = backColor;
            p["ForeColor"] = foreColor;
            p["BorderColor"] = borderColor;
            p["BorderStyle"] = (int)borderStyle;
            p["TextAlign"] = (int)textAlign;
            p["FontName"] = fontName;
            p["FontSize"] = fontSize;
            p["FontBold"] = fontBold;
            p["FontItalic"] = fontItalic;
        }

        protected override void LoadProperties(IDictionary<string, object> p)
        {
            value = FromIso(Get(p, "Value", Iso(value)));
            format = Get(p, "Format", format);
            minDate = FromIso(Get(p, "MinDate", Iso(minDate)));
            maxDate = FromIso(Get(p, "MaxDate", Iso(maxDate)));
            placeholder = Get(p, "Placeholder", placeholder);
            locked = Get(p, "Locked", locked);
            enabled = Get(p, "Enabled", enabled);
            backColor = Get(p, "BackColor", backColor);
            foreColor = Get(p, "ForeColor", foreColor);
            borderColor = Get(p, "BorderColor", borderColor);
            borderStyle = (sgBorderStyle)Get(p, "BorderStyle", (int)borderStyle);
            textAlign = (sgTextAlign)Get(p, "TextAlign", (int)textAlign);
            fontName = Get(p, "FontName", fontName);
            fontSize = Get(p, "FontSize", fontSize);
            fontBold = Get(p, "FontBold", fontBold);
            fontItalic = Get(p, "FontItalic", fontItalic);
            ResetFont();
            if (edit != null) Apply();
        }

        static string Iso(DateTime? d) { return d.HasValue ? d.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : ""; }

        static DateTime? FromIso(string s)
        {
            DateTime d;
            return DateTime.TryParseExact(s ?? "", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out d) ? d : (DateTime?)null;
        }

        // ------------------------------------------------------------------
        // Janela
        // ------------------------------------------------------------------

        protected override Control CreateContent()
        {
            edit = new DateEdit(this);
            edit.BorderStyle = System.Windows.Forms.BorderStyle.None;
            edit.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (!EventsEnabled || KeyDown == null) return;
                short code = (short)e.KeyCode;
                KeyDown(ref code, (short)((e.Shift ? 1 : 0) | (e.Control ? 2 : 0) | (e.Alt ? 4 : 0)));
                if (code == 0) { e.Handled = true; e.SuppressKeyPress = true; }
            };
            edit.KeyPress += delegate(object sender, KeyPressEventArgs e)
            {
                if (EventsEnabled && KeyPress != null)
                {
                    short ascii = (short)e.KeyChar;
                    KeyPress(ref ascii);
                    if (ascii == 0) { e.Handled = true; return; }
                    e.KeyChar = (char)ascii;
                }
                Typed(e);
            };
            edit.GotFocus += delegate
            {
                ShowText();
                edit.SelectAll();
                ContentFocus(true);
                if (Host != null) Host.Invalidate();
            };
            edit.LostFocus += delegate
            {
                if (!edit.Focused) { Commit(); ShowText(); }
                ContentFocus(false);
                if (Host != null) Host.Invalidate();
            };
            Host.Layout += delegate { LayoutEdit(); };
            Host.MouseDown += delegate(object sender, MouseEventArgs e)
            {
                if (!UserMode || e.Button != MouseButtons.Left || !enabled) return;
                edit.Focus();
                if (!ButtonRect(Host.ClientRectangle).Contains(e.Location) || locked) return;
                // O WinForms prende o mouse na caixa até o botão subir e, ao soltar, libera a captura,
                // que a essa altura é a do calendário: solta antes de abrir
                Host.Capture = false;
                Toggle();
            };
            Host.MouseMove += delegate(object sender, MouseEventArgs e)
            {
                bool over = UserMode && ButtonRect(Host.ClientRectangle).Contains(e.Location);
                if (over != buttonHover) { buttonHover = over; Host.Invalidate(); }
            };
            Host.MouseLeave += delegate { if (buttonHover) { buttonHover = false; Host.Invalidate(); } };
            Apply();
            return edit;
        }

        void ResetFont()
        {
            FontStyle style = (fontBold ? FontStyle.Bold : FontStyle.Regular) | (fontItalic ? FontStyle.Italic : FontStyle.Regular);
            Font old = font;
            font = new Font(fontName, (float)Math.Max(1, fontSize), style, GraphicsUnit.Point);
            if (edit != null) edit.Font = font;
            if (old != null) old.Dispose();
        }

        void Apply()
        {
            if (edit == null) return;
            edit.Font = font;
            edit.ReadOnly = locked;
            edit.Enabled = enabled;
            edit.TextAlign = textAlign == sgTextAlign.sgTextAlignCenter ? HorizontalAlignment.Center
                : textAlign == sgTextAlign.sgTextAlignRight ? HorizontalAlignment.Right : HorizontalAlignment.Left;
            edit.BackColor = Back;
            edit.ForeColor = ColorTranslator.FromOle(foreColor);
            ShowText();
            LayoutEdit();
            edit.Invalidate();
            if (Host != null) Host.Invalidate();
        }

        // O texto da caixa: para digitar (com o foco) ou no Format
        void ShowText()
        {
            if (edit == null) return;
            string t = edit.Focused && UserMode ? EditText(value) : Display(value);
            if (edit.Text == t) return;
            edit.Text = t;
        }

        Color Back { get { return enabled ? ColorTranslator.FromOle(backColor) : SystemColors.Control; } }

        static float ScaleOf(Graphics g) { return g.DpiX / 96f; }

        float Scale
        {
            get
            {
                if (Host == null || !Host.IsHandleCreated) return 1;
                using (Graphics g = Host.CreateGraphics()) return ScaleOf(g);
            }
        }

        int Border { get { return borderStyle == sgBorderStyle.sgBorderStyleSingle ? 1 : 0; } }

        Rectangle ButtonRect(Rectangle r) { return ButtonRect(r, Scale); }

        Rectangle ButtonRect(Rectangle r, float scale)
        {
            int w = Math.Min((int)Math.Round(22 * scale), Math.Max(0, r.Width - 2 * Border));
            return new Rectangle(r.Right - Border - w, r.Top + Border, w, Math.Max(0, r.Height - 2 * Border));
        }

        void LayoutEdit()
        {
            if (edit == null || Host == null) return;
            Size size = Host.ClientSize;
            float scale = Scale;
            int pad = (int)Math.Round(4 * scale) + Border;
            int button = ButtonRect(new Rectangle(Point.Empty, size), scale).Width;
            int h = Math.Min(edit.PreferredHeight, Math.Max(1, size.Height - 2 * Border));
            edit.SetBounds(pad, Math.Max(Border, (size.Height - h) / 2), Math.Max(1, size.Width - 2 * pad - button), h);
        }

        protected override void Paint(Graphics g, Rectangle r, bool windowed)
        {
            float scale = ScaleOf(g);
            using (SolidBrush back = new SolidBrush(Back)) g.FillRectangle(back, r);
            Rectangle button = ButtonRect(r, scale);
            if (!windowed)
            {
                int pad = (int)Math.Round(4 * scale) + Border;
                Rectangle area = new Rectangle(r.Left + pad, r.Top + Border, Math.Max(0, button.Left - r.Left - pad), r.Height - 2 * Border);
                string shown = value.HasValue ? Display(value) : placeholder;
                Color color = value.HasValue && enabled ? ColorTranslator.FromOle(foreColor) : SystemColors.GrayText;
                TextRenderer.DrawText(g, shown, font, area, color, TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis |
                    TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | AlignFlags());
            }
            bool open = popup != null && popup.Visible;
            if (windowed && enabled && !locked && (buttonHover || open))
                using (SolidBrush hover = new SolidBrush(open ? Color.FromArgb(204, 228, 247) : Color.FromArgb(229, 241, 251)))
                    g.FillRectangle(hover, button);
            DrawCalendarGlyph(g, button, enabled && !locked ? Color.FromArgb(96, 96, 96) : SystemColors.GrayText, scale);
            if (Border > 0)
            {
                bool focused = windowed && edit != null && (edit.Focused || open);
                Color color = !enabled ? SystemColors.ControlDark : focused ? Accent : ColorTranslator.FromOle(borderColor);
                using (Pen pen = new Pen(color)) g.DrawRectangle(pen, r.Left, r.Top, r.Width - 1, r.Height - 1);
            }
        }

        // Um calendário pequeno: moldura, faixa de cima e dois pontos de dias
        static void DrawCalendarGlyph(Graphics g, Rectangle button, Color color, float scale)
        {
            int w = (int)Math.Round(12 * scale), h = (int)Math.Round(11 * scale);
            if (button.Width < w + 2 || button.Height < h + 2) return;
            int x = button.Left + (button.Width - w) / 2, y = button.Top + (button.Height - h) / 2 + 1;
            using (Pen pen = new Pen(color))
            using (SolidBrush brush = new SolidBrush(color))
            {
                g.DrawRectangle(pen, x, y, w - 1, h - 1);
                g.FillRectangle(brush, x, y, w, Math.Max(2, (int)Math.Round(3 * scale)));
                int ring = Math.Max(1, (int)Math.Round(2 * scale));
                g.FillRectangle(brush, x + w / 4, y - ring, Math.Max(1, (int)scale), ring);
                g.FillRectangle(brush, x + 3 * w / 4 - 1, y - ring, Math.Max(1, (int)scale), ring);
                int dot = Math.Max(2, (int)Math.Round(2 * scale));
                for (int i = 0; i < 3; i++)
                    g.FillRectangle(brush, x + 2 + i * (w - 4 - dot) / 2, y + h - dot - (int)Math.Round(2 * scale), dot, dot);
            }
        }

        TextFormatFlags AlignFlags()
        {
            return textAlign == sgTextAlign.sgTextAlignCenter ? TextFormatFlags.HorizontalCenter
                : textAlign == sgTextAlign.sgTextAlignRight ? TextFormatFlags.Right : TextFormatFlags.Left;
        }

        internal void PaintPlaceholder(Graphics g)
        {
            if (placeholder.Length == 0 || edit.TextLength > 0) return;
            Rectangle area = edit.ClientRectangle;
            area.X += 1; area.Width -= 2;
            TextRenderer.DrawText(g, placeholder, edit.Font, area, SystemColors.GrayText,
                TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | AlignFlags());
        }

        // ------------------------------------------------------------------
        // Texto <-> data
        // ------------------------------------------------------------------

        static DateTimeFormatInfo Info { get { return CultureInfo.CurrentCulture.DateTimeFormat; } }

        // A ordem de dia, mês e ano na data abreviada do Windows ("dMy", "Mdy", "yMd")
        static string Order
        {
            get
            {
                string p = Info.ShortDatePattern;
                int d = p.IndexOf('d'), m = p.IndexOf('M'), y = p.IndexOf('y');
                if (d < 0 || m < 0 || y < 0) return "dMy";
                List<KeyValuePair<int, char>> parts = new List<KeyValuePair<int, char>>
                {
                    new KeyValuePair<int, char>(d, 'd'), new KeyValuePair<int, char>(m, 'M'), new KeyValuePair<int, char>(y, 'y'),
                };
                parts.Sort((a, b) => a.Key.CompareTo(b.Key));
                return new string(new[] { parts[0].Value, parts[1].Value, parts[2].Value });
            }
        }

        static string Separator { get { string s = Info.DateSeparator; return string.IsNullOrEmpty(s) ? "/" : s; } }

        // Para digitar: só números, dia e mês com 2 dígitos e ano com 4, na ordem do Windows
        static string EditText(DateTime? d)
        {
            if (!d.HasValue) return "";
            List<string> parts = new List<string>();
            foreach (char c in Order)
                parts.Add(c == 'd' ? d.Value.Day.ToString("00") : c == 'M' ? d.Value.Month.ToString("00") : d.Value.Year.ToString("0000"));
            return string.Join(Separator, parts.ToArray());
        }

        string Display(DateTime? d)
        {
            if (!d.HasValue) return "";
            if (format.Length > 0)
            {
                object input = d.Value;
                string text;
                try { if (VarFormat(ref input, format, 1, 1, 0, out text) == 0) return text; }
                catch (Exception) { }
            }
            return d.Value.ToString("d", CultureInfo.CurrentCulture);
        }

        // "12/10/2026", "12-10-26", "12102026", "121026", "1210" (ano atual), "12" (mês atual) ou
        // o que a data do Windows aceitar ("12 out 2026")
        internal static DateTime? Parse(string text, out bool ok)
        {
            ok = true;
            string s = (text ?? "").Trim();
            if (s.Length == 0) return null;
            DateTime today = DateTime.Today;
            List<string> parts = new List<string>();
            bool digitsOnly = true;
            string current = "";
            foreach (char c in s)
            {
                if (char.IsDigit(c)) current += c;
                else
                {
                    if (char.IsLetter(c)) digitsOnly = false;
                    if (current.Length > 0) { parts.Add(current); current = ""; }
                }
            }
            if (current.Length > 0) parts.Add(current);

            if (digitsOnly && parts.Count == 1)
            {
                string n = parts[0];
                if (n.Length <= 2) parts = new List<string> { n };
                else if (n.Length <= 4) parts = new List<string> { n.Substring(0, 2), n.Substring(2) };
                else if (n.Length == 6 || n.Length == 8)
                {
                    string order = Order;
                    int yearLen = n.Length - 4, at = 0;
                    parts = new List<string>();
                    foreach (char c in order) { int len = c == 'y' ? yearLen : 2; parts.Add(n.Substring(at, len)); at += len; }
                }
                else parts = null;
            }

            if (digitsOnly && parts != null && parts.Count >= 1 && parts.Count <= 3)
            {
                string order = Order;
                if (parts.Count < 3) order = order.Replace("y", "");
                int day = today.Day, month = today.Month, year = today.Year;
                if (parts.Count == 1) day = int.Parse(parts[0], CultureInfo.InvariantCulture);
                else
                    for (int i = 0; i < parts.Count; i++)
                    {
                        int v = int.Parse(parts[i], CultureInfo.InvariantCulture);
                        if (order[i] == 'd') day = v;
                        else if (order[i] == 'M') month = v;
                        else year = parts[i].Length <= 2 ? CultureInfo.CurrentCulture.Calendar.ToFourDigitYear(v) : v;
                    }
                if (year >= 100 && year <= 9999 && month >= 1 && month <= 12 && day >= 1 && day <= DateTime.DaysInMonth(year, month))
                    return new DateTime(year, month, day);
                ok = false;
                return null;
            }

            DateTime parsed;
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out parsed)) return parsed.Date;
            ok = false;
            return null;
        }

        bool InRange(DateTime d)
        {
            return (!minDate.HasValue || d >= minDate.Value) && (!maxDate.HasValue || d <= maxDate.Value);
        }

        // Confere o texto digitado: data válida vira o valor; inválida ou fora dos limites, volta a anterior
        internal void Commit()
        {
            if (edit == null || !UserMode || locked) return;
            bool ok;
            DateTime? d = Parse(edit.Text, out ok);
            if (!ok || (d.HasValue && !InRange(d.Value)))
            {
                SystemSounds.Beep.Play();
                ShowText();
                if (edit.Focused) edit.SelectAll();
                return;
            }
            SetValue(d, true);
            ShowText();
        }

        // Novo valor: grava, mostra e dispara o Change se mudou
        void SetValue(DateTime? d, bool fromUser)
        {
            if (d.HasValue) d = d.Value.Date;
            bool changed = d != value;
            value = d;
            ShowText();
            if (changed)
            {
                Changed();
                if (EventsEnabled && Change != null) Change();
            }
            else if (!fromUser) Changed();
        }

        // Digitação: só números e separadores; o separador entra sozinho depois do dia e do mês
        void Typed(KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if (c < 32 || locked) return;
            bool separator = c == '/' || c == '-' || c == '.' || c == ' ' || Separator.IndexOf(c) >= 0;
            if (!char.IsDigit(c) && !separator) { e.Handled = true; return; }
            if (edit.SelectionLength > 0 || edit.SelectionStart != edit.TextLength) return;
            string t = edit.Text;
            List<int> groups = Groups(t);
            if (separator)
            {
                // Sem separador repetido nem no começo
                if (t.Length == 0 || !char.IsDigit(t[t.Length - 1]) || groups.Count >= 3) e.Handled = true;
                else { e.Handled = true; edit.AppendText(Separator); }
                return;
            }
            groups = Groups(t + c);
            string order = Order;
            int index = groups.Count - 1;
            if (index < 2 && groups[index] == 2 && order[index] != 'y')
            {
                e.Handled = true;
                edit.AppendText(c + Separator);
            }
        }

        // Quantos dígitos em cada grupo ("12/1" -> 2, 1)
        static List<int> Groups(string t)
        {
            List<int> groups = new List<int> { 0 };
            foreach (char c in t)
            {
                if (char.IsDigit(c)) groups[groups.Count - 1]++;
                else if (groups[groups.Count - 1] > 0) groups.Add(0);
            }
            return groups;
        }

        // Seta para cima/baixo: um dia a mais ou a menos (sem data, começa em hoje)
        internal void Step(int days)
        {
            if (locked) return;
            Commit();
            DateTime d = value.HasValue ? value.Value.AddDays(days) : DateTime.Today;
            if (!InRange(d)) { SystemSounds.Beep.Play(); return; }
            SetValue(d, true);
            edit.SelectAll();
        }

        // ------------------------------------------------------------------
        // Calendário
        // ------------------------------------------------------------------

        internal bool PopupOpen { get { return popup != null && popup.Visible; } }

        internal void Toggle()
        {
            if (PopupOpen) ClosePopup();
            else OpenPopup();
        }

        void OpenPopup()
        {
            if (Host == null || !Host.IsHandleCreated || !UserMode || locked || !enabled) return;
            Commit();
            if (popup == null) popup = new CalendarPopup(this);
            popup.Prepare(value ?? (InRange(DateTime.Today) ? DateTime.Today : minDate ?? maxDate ?? DateTime.Today), value, minDate, maxDate, font);
            Point below = Host.PointToScreen(new Point(0, Host.Height));
            popup.ShowAt(below, Host.PointToScreen(Point.Empty).Y, GetAncestor(Host.Handle, GA_ROOT));
            Host.Invalidate();
            if (EventsEnabled && DropButtonClick != null) DropButtonClick();
        }

        internal void ClosePopup()
        {
            if (popup != null && popup.Visible) { popup.Capture = false; popup.Hide(); }
            if (Host != null) Host.Invalidate();
        }

        // Data escolhida no calendário (ou Limpar: null)
        internal void Picked(DateTime? d)
        {
            ClosePopup();
            if (edit != null && !edit.Focused) edit.Focus();
            SetValue(d, true);
            if (edit != null) edit.SelectAll();
        }

        internal bool PopupKey(Keys key) { return popup != null && popup.Key(key); }

        internal bool Forward(ref Message m) { return ForwardKey(ref m); }
        internal bool IsUserMode { get { return UserMode; } }
        internal bool IsLocked { get { return locked; } }

        // ------------------------------------------------------------------
        // _DateBox
        // ------------------------------------------------------------------

        // Date ou Empty
        public object Value
        {
            get
            {
                if (edit != null && edit.Focused && UserMode)
                {
                    // O que está digitado e ainda não foi conferido
                    bool ok;
                    DateTime? typed = Parse(edit.Text, out ok);
                    if (ok && (!typed.HasValue || InRange(typed.Value))) return typed.HasValue ? (object)typed.Value : null;
                }
                return value.HasValue ? (object)value.Value : null;
            }
            set { LetValue(value); }
        }

        public void LetValue(object Value)
        {
            DateTime? d = ToDate(Value, "Value");
            if (d.HasValue && !InRange(d.Value))
                throw Interop.Error(380, "ValueError: a data " + d.Value.ToString("d", CultureInfo.CurrentCulture) + " está fora de MinDate/MaxDate");
            SetValue(d, false);
        }

        public string Text
        {
            get { return edit != null ? edit.Text : Display(value); }
            set
            {
                bool ok;
                DateTime? d = Parse(value, out ok);
                if (!ok) throw Interop.Error(13, "TypeError: '" + value + "' não é uma data");
                LetValue(d.HasValue ? (object)d.Value : null);
            }
        }

        public string Format { get { return format; } set { format = value ?? ""; ShowText(); Changed(); } }

        public object MinDate
        {
            get { return minDate.HasValue ? (object)minDate.Value : null; }
            set { LetMinDate(value); }
        }

        public void LetMinDate(object Value) { minDate = ToDate(Value, "MinDate"); Changed(); }

        public object MaxDate
        {
            get { return maxDate.HasValue ? (object)maxDate.Value : null; }
            set { LetMaxDate(value); }
        }

        public void LetMaxDate(object Value) { maxDate = ToDate(Value, "MaxDate"); Changed(); }

        // Date, texto de data, número (data serial) ou Empty/Null/"" (sem data)
        static DateTime? ToDate(object v, string name)
        {
            v = Interop.Unwrap(v);
            if (v == null || v is DBNull || Interop.IsMissing(v)) return null;
            if (v is DateTime) return ((DateTime)v).Date;
            string s = v as string;
            if (s != null)
            {
                bool ok;
                DateTime? d = Parse(s, out ok);
                if (ok) return d;
                throw Interop.Error(13, "TypeError: " + name + " recebe uma data; '" + s + "' não é uma data");
            }
            if (v is double || v is float || v is int || v is short || v is long || v is decimal)
            {
                try { return DateTime.FromOADate(Convert.ToDouble(v, CultureInfo.InvariantCulture)).Date; }
                catch (ArgumentException) { }
            }
            throw Interop.Error(13, "TypeError: " + name + " recebe uma data (Date), não " + v.GetType().Name);
        }

        public string Placeholder { get { return placeholder; } set { placeholder = value ?? ""; Apply(); Changed(); } }
        public bool Locked { get { return locked; } set { locked = value; if (locked) ClosePopup(); Apply(); Changed(); } }
        public bool Enabled { get { return enabled; } set { enabled = value; if (!enabled) ClosePopup(); Apply(); Changed(); } }
        public int BackColor { get { return backColor; } set { backColor = value; Apply(); Changed(); } }
        public int ForeColor { get { return foreColor; } set { foreColor = value; Apply(); Changed(); } }
        public int BorderColor { get { return borderColor; } set { borderColor = value; Apply(); Changed(); } }
        public sgBorderStyle BorderStyle { get { return borderStyle; } set { borderStyle = value; Apply(); Changed(); } }
        public sgTextAlign TextAlign { get { return textAlign; } set { textAlign = value; Apply(); Changed(); } }

        public string FontName
        {
            get { return fontName; }
            set { if (!string.IsNullOrEmpty(value)) { fontName = value; ResetFont(); Apply(); Changed(); } }
        }

        public double FontSize
        {
            get { return fontSize; }
            set
            {
                if (value <= 0 || value > 400) throw Interop.Error(380, "ValueError: FontSize deve ficar entre 1 e 400");
                fontSize = value; ResetFont(); Apply(); Changed();
            }
        }

        public bool FontBold { get { return fontBold; } set { fontBold = value; ResetFont(); Apply(); Changed(); } }
        public bool FontItalic { get { return fontItalic; } set { fontItalic = value; ResetFont(); Apply(); Changed(); } }

        // Abre o calendário (como o DropDown da ComboBox)
        public void DropDown()
        {
            if (edit != null && !edit.Focused) edit.Focus();
            if (!PopupOpen) OpenPopup();
        }

        const int GA_ROOT = 2;
        [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr hwnd, int flags);
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
        static extern int VarFormat([MarshalAs(UnmanagedType.Struct)] ref object input, string format, int firstDay, int firstWeek, int flags,
            [MarshalAs(UnmanagedType.BStr)] out string output);
    }

    // A caixa de edição do DateBox. Como no TextBox, as teclas do form (Tab, Enter, Esc) são
    // passadas a ele; F4/Alt+Seta para baixo abrem o calendário e, com ele aberto, as teclas de
    // navegação vão para o calendário.
    sealed class DateEdit : System.Windows.Forms.TextBox
    {
        const int WM_PAINT = 0x0F, WM_KEYDOWN = 0x100, WM_SYSKEYDOWN = 0x104, WM_CHAR = 0x102, WM_NCHITTEST = 0x84, HTTRANSPARENT = -1;
        readonly DateBox owner;

        public DateEdit(DateBox owner) { this.owner = owner; }

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_NCHITTEST:
                    if (!owner.IsUserMode) { m.Result = (IntPtr)HTTRANSPARENT; return; }
                    break;
                case WM_SYSKEYDOWN:
                    {
                        Keys key = (Keys)(int)m.WParam;
                        if (key == Keys.Down || key == Keys.Up) { owner.Toggle(); return; }
                        break;
                    }
                case WM_KEYDOWN:
                    {
                        Keys key = (Keys)(int)m.WParam;
                        bool ctrl = (ModifierKeys & Keys.Control) != 0;
                        if (key == Keys.F4) { owner.Toggle(); return; }
                        if (owner.PopupOpen)
                        {
                            if (key == Keys.Escape) { owner.ClosePopup(); return; }
                            if (key == Keys.Tab) owner.ClosePopup();
                            else if (owner.PopupKey(key)) return;
                        }
                        if (key == Keys.Up && !ctrl) { owner.Step(1); return; }
                        if (key == Keys.Down && !ctrl) { owner.Step(-1); return; }
                        if (key == Keys.Return) owner.Commit();
                        if (key == Keys.Tab || key == Keys.Escape || key == Keys.Return)
                        {
                            owner.Forward(ref m);
                            return;
                        }
                        if (key == Keys.A && ctrl && (ModifierKeys & Keys.Alt) == 0) { SelectAll(); return; }
                        break;
                    }
                case WM_CHAR:
                    {
                        int c = (int)m.WParam;
                        if (c == '\t' || c == 27 || c == 1 || c == '\r') return;
                        break;
                    }
            }
            base.WndProc(ref m);
            if (m.Msg == WM_PAINT)
                using (Graphics g = CreateGraphics()) owner.PaintPlaceholder(g);
        }
    }

    // O calendário: janela solta, sem ativar (o foco continua na caixa), com a captura do mouse
    // enquanto aberta (clique fora fecha). Três vistas: dias, meses e anos.
    sealed class CalendarPopup : Form
    {
        enum View { Days, Months, Years }

        readonly DateBox owner;
        DateTime focus, shown;           // dia em destaque e o mês (ou ano) mostrado
        DateTime? selected, min, max;
        View view;
        Font font, bold;
        float scale = 1;
        Point mouse = new Point(-1, -1);
        bool releasing; // tratando um botão do mouse solto sobre o calendário

        public CalendarPopup(DateBox owner)
        {
            this.owner = owner;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            BackColor = SystemColors.Window;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000 | 0x80; // WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW
                cp.ClassStyle |= 0x20000;            // CS_DROPSHADOW
                return cp;
            }
        }

        public void Prepare(DateTime start, DateTime? selected, DateTime? min, DateTime? max, Font baseFont)
        {
            this.selected = selected; this.min = min; this.max = max;
            focus = start.Date;
            shown = new DateTime(focus.Year, focus.Month, 1);
            view = View.Days;
            if (font != null) { font.Dispose(); bold.Dispose(); }
            font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
            bold = new Font("Segoe UI", 9f, FontStyle.Bold, GraphicsUnit.Point);
        }

        // Abaixo da caixa; sem espaço embaixo, em cima
        public void ShowAt(Point below, int top, IntPtr owner)
        {
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) scale = g.DpiX / 96f;
            Size size = new Size(S(252), S(292));
            Rectangle screen = Screen.FromPoint(below).WorkingArea;
            int x = Math.Max(screen.Left, Math.Min(below.X, screen.Right - size.Width));
            int y = below.Y + S(2);
            if (y + size.Height > screen.Bottom) y = Math.Max(screen.Top, top - size.Height - S(2));
            Bounds = new Rectangle(x, y, size.Width, size.Height);
            if (!IsHandleCreated) CreateHandle();
            SetWindowLongPtr(Handle, GWLP_HWNDPARENT, owner); // fica na frente do form e some com ele
            Show();
            Capture = true;
            Invalidate();
        }

        int S(int px) { return (int)Math.Round(px * scale); }

        // Clique fora (ou outra janela pegou o mouse): fecha
        protected override void OnMouseCaptureChanged(EventArgs e)
        {
            base.OnMouseCaptureChanged(e);
            if (!Visible || Capture) return;
            // O WinForms libera a captura no fim de todo MouseUp (o do clique que abriu o calendário,
            // as setas, o título...): aí ela volta. Outra perda (outra janela pegou o mouse, Alt+Tab) fecha.
            bool ownRelease = releasing;
            BeginInvoke((Action)delegate
            {
                if (!Visible || Capture) return;
                if (ownRelease) Capture = true;
                else owner.ClosePopup();
            });
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x21) { m.Result = (IntPtr)3; return; } // WM_MOUSEACTIVATE: MA_NOACTIVATE
            // WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP, WM_XBUTTONUP
            bool up = m.Msg == 0x202 || m.Msg == 0x205 || m.Msg == 0x208 || m.Msg == 0x20C;
            if (up) releasing = true;
            try { base.WndProc(ref m); }
            finally { if (up) releasing = false; }
        }

        // ------------------------------------------------------------------
        // Geometria
        // ------------------------------------------------------------------

        Rectangle Header { get { return new Rectangle(S(8), S(8), ClientSize.Width - S(16), S(32)); } }
        Rectangle Prev { get { Rectangle h = Header; return new Rectangle(h.Right - S(64), h.Top, S(32), h.Height); } }
        Rectangle Next { get { Rectangle h = Header; return new Rectangle(h.Right - S(32), h.Top, S(32), h.Height); } }
        Rectangle Title { get { Rectangle h = Header; return new Rectangle(h.Left, h.Top, h.Width - S(68), h.Height); } }
        Rectangle Body { get { return new Rectangle(S(8), S(44), ClientSize.Width - S(16), S(208)); } }
        Rectangle TodayButton { get { return new Rectangle(S(8), ClientSize.Height - S(36), S(110), S(28)); } }
        Rectangle ClearButton { get { return new Rectangle(ClientSize.Width - S(8) - S(110), ClientSize.Height - S(36), S(110), S(28)); } }

        DayOfWeek FirstDay { get { return CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek; } }

        // O primeiro dia da grade (6 semanas a partir da semana do dia 1)
        DateTime GridStart
        {
            get
            {
                int back = ((int)shown.DayOfWeek - (int)FirstDay + 7) % 7;
                return shown.AddDays(-back);
            }
        }

        Rectangle DayCell(int i)
        {
            Rectangle b = Body;
            int w = b.Width / 7, h = (b.Height - S(28)) / 6;
            return new Rectangle(b.Left + (i % 7) * w, b.Top + S(28) + (i / 7) * h, w, h);
        }

        Rectangle BigCell(int i) // meses e anos: 4 colunas x 3 linhas
        {
            Rectangle b = Body;
            int w = b.Width / 4, h = b.Height / 3;
            return new Rectangle(b.Left + (i % 4) * w, b.Top + (i / 4) * h, w, h);
        }

        int FirstYear { get { return shown.Year - shown.Year % 12; } }

        bool Allowed(DateTime d) { return (!min.HasValue || d >= min.Value) && (!max.HasValue || d <= max.Value); }

        bool MonthAllowed(int year, int month)
        {
            DateTime first = new DateTime(year, month, 1), last = first.AddMonths(1).AddDays(-1);
            return (!min.HasValue || last >= min.Value) && (!max.HasValue || first <= max.Value);
        }

        bool YearAllowed(int year)
        {
            return (!min.HasValue || year >= min.Value.Year) && (!max.HasValue || year <= max.Value.Year);
        }

        // ------------------------------------------------------------------
        // Desenho
        // ------------------------------------------------------------------

        static string Capital(string s) { return s.Length == 0 ? s : char.ToUpper(s[0], CultureInfo.CurrentCulture) + s.Substring(1); }

        string TitleText
        {
            get
            {
                CultureInfo ci = CultureInfo.CurrentCulture;
                if (view == View.Days) return Capital(shown.ToString(ci.DateTimeFormat.YearMonthPattern, ci));
                if (view == View.Months) return shown.Year.ToString(ci);
                return FirstYear + " – " + (FirstYear + 11);
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(SystemColors.Window);
            using (Pen border = new Pen(Color.FromArgb(200, 200, 200))) g.DrawRectangle(border, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
            CultureInfo ci = CultureInfo.CurrentCulture;
            Color text = SystemColors.WindowText, gray = SystemColors.GrayText, faint = Color.FromArgb(190, 190, 190);

            // Cabeçalho
            Hover(g, Title);
            TextRenderer.DrawText(g, TitleText, bold, Title, text, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.NoPrefix);
            Hover(g, Prev); Hover(g, Next);
            Arrow(g, Prev, true, text); Arrow(g, Next, false, text);

            if (view == View.Days)
            {
                string[] names = ci.DateTimeFormat.ShortestDayNames;
                for (int c = 0; c < 7; c++)
                {
                    Rectangle cell = DayCell(c);
                    Rectangle head = new Rectangle(cell.Left, Body.Top, cell.Width, S(28));
                    string n = names[((int)FirstDay + c) % 7];
                    TextRenderer.DrawText(g, Capital(n), font, head, gray, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
                }
                DateTime d = GridStart;
                for (int i = 0; i < 42; i++, d = d.AddDays(1))
                {
                    Rectangle cell = DayCell(i);
                    bool allowed = Allowed(d), inMonth = d.Month == shown.Month;
                    Color color = !allowed ? faint : inMonth ? text : gray;
                    Cell(g, cell, d.Day.ToString(ci), color, allowed, selected.HasValue && d == selected.Value, d == DateTime.Today, d == focus);
                }
            }
            else if (view == View.Months)
            {
                string[] names = ci.DateTimeFormat.AbbreviatedMonthNames;
                for (int i = 0; i < 12; i++)
                {
                    bool allowed = MonthAllowed(shown.Year, i + 1);
                    bool isSel = selected.HasValue && selected.Value.Year == shown.Year && selected.Value.Month == i + 1;
                    bool isToday = DateTime.Today.Year == shown.Year && DateTime.Today.Month == i + 1;
                    Cell(g, BigCell(i), Capital(names[i].TrimEnd('.')), allowed ? text : faint, allowed, isSel, isToday, focus.Year == shown.Year && focus.Month == i + 1);
                }
            }
            else
            {
                for (int i = 0; i < 12; i++)
                {
                    int year = FirstYear + i;
                    bool allowed = YearAllowed(year);
                    Cell(g, BigCell(i), year.ToString(ci), allowed ? text : faint, allowed, selected.HasValue && selected.Value.Year == year,
                        DateTime.Today.Year == year, focus.Year == year);
                }
            }

            // Rodapé
            bool pt = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "pt";
            bool todayAllowed = Allowed(DateTime.Today);
            Hover(g, TodayButton, todayAllowed);
            TextRenderer.DrawText(g, pt ? "Hoje" : "Today", font, TodayButton, todayAllowed ? Color.FromArgb(0, 95, 184) : faint,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            if (!owner.IsLocked)
            {
                Hover(g, ClearButton);
                TextRenderer.DrawText(g, pt ? "Limpar" : "Clear", font, ClearButton, Color.FromArgb(0, 95, 184),
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            }
        }

        void Hover(Graphics g, Rectangle r, bool allowed = true)
        {
            if (!allowed || !r.Contains(mouse)) return;
            using (SolidBrush b = new SolidBrush(Color.FromArgb(229, 241, 251))) g.FillRectangle(b, r);
        }

        void Cell(Graphics g, Rectangle cell, string label, Color color, bool allowed, bool isSelected, bool isToday, bool isFocus)
        {
            Rectangle r = Rectangle.Inflate(cell, -S(2), -S(2));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            if (isSelected)
                using (SolidBrush b = new SolidBrush(DateBox.Accent)) Round(g, b, null, r);
            else if (allowed && r.Contains(mouse))
                using (SolidBrush b = new SolidBrush(Color.FromArgb(229, 241, 251))) Round(g, b, null, r);
            if (isToday && !isSelected)
                using (Pen p = new Pen(DateBox.Accent, Math.Max(1, scale))) Round(g, null, p, r);
            if (isFocus && !isSelected && !isToday)
                using (Pen p = new Pen(Color.FromArgb(150, 150, 150)) { DashStyle = DashStyle.Dot }) Round(g, null, p, r);
            g.SmoothingMode = SmoothingMode.Default;
            TextRenderer.DrawText(g, label, isToday || isSelected ? bold : font, cell, isSelected ? Color.White : color,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
        }

        void Round(Graphics g, Brush fill, Pen pen, Rectangle r)
        {
            int radius = S(4);
            using (GraphicsPath path = new GraphicsPath())
            {
                path.AddArc(r.Left, r.Top, radius * 2, radius * 2, 180, 90);
                path.AddArc(r.Right - radius * 2, r.Top, radius * 2, radius * 2, 270, 90);
                path.AddArc(r.Right - radius * 2, r.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
                path.AddArc(r.Left, r.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
                path.CloseFigure();
                if (fill != null) g.FillPath(fill, path);
                if (pen != null) g.DrawPath(pen, path);
            }
        }

        void Arrow(Graphics g, Rectangle r, bool left, Color color)
        {
            int cx = r.Left + r.Width / 2, cy = r.Top + r.Height / 2, s = S(4);
            Point[] pts = left
                ? new[] { new Point(cx + s / 2, cy - s), new Point(cx - s / 2, cy), new Point(cx + s / 2, cy + s) }
                : new[] { new Point(cx - s / 2, cy - s), new Point(cx + s / 2, cy), new Point(cx - s / 2, cy + s) };
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (Pen p = new Pen(color, Math.Max(1.2f, 1.4f * scale))) g.DrawLines(p, pts);
            g.SmoothingMode = SmoothingMode.Default;
        }

        // ------------------------------------------------------------------
        // Mouse e teclado
        // ------------------------------------------------------------------

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (mouse != e.Location) { mouse = e.Location; Invalidate(); }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (!ClientRectangle.Contains(e.Location)) { owner.ClosePopup(); return; }
            if (e.Button != MouseButtons.Left) return;
            Point p = e.Location;
            if (Prev.Contains(p)) Page(-1);
            else if (Next.Contains(p)) Page(1);
            else if (Title.Contains(p)) { if (view == View.Days) view = View.Months; else if (view == View.Months) view = View.Years; }
            else if (TodayButton.Contains(p)) { if (Allowed(DateTime.Today)) owner.Picked(DateTime.Today); return; }
            else if (ClearButton.Contains(p) && !owner.IsLocked) { owner.Picked(null); return; }
            else if (view == View.Days)
            {
                for (int i = 0; i < 42; i++)
                    if (DayCell(i).Contains(p))
                    {
                        DateTime d = GridStart.AddDays(i);
                        if (Allowed(d)) { owner.Picked(d); return; }
                    }
            }
            else
            {
                for (int i = 0; i < 12; i++)
                    if (BigCell(i).Contains(p))
                    {
                        if (view == View.Months)
                        {
                            if (!MonthAllowed(shown.Year, i + 1)) return;
                            shown = new DateTime(shown.Year, i + 1, 1);
                            focus = Clamp(new DateTime(shown.Year, shown.Month, Math.Min(focus.Day, DateTime.DaysInMonth(shown.Year, shown.Month))));
                            view = View.Days;
                        }
                        else
                        {
                            int year = FirstYear + i;
                            if (!YearAllowed(year)) return;
                            shown = new DateTime(year, shown.Month, 1);
                            view = View.Months;
                        }
                        break;
                    }
            }
            Invalidate();
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            Page(e.Delta > 0 ? -1 : 1);
            Invalidate();
        }

        // Página anterior/seguinte: um mês, um ano ou 12 anos, conforme a vista
        void Page(int direction)
        {
            try
            {
                if (view == View.Days) shown = shown.AddMonths(direction);
                else if (view == View.Months) shown = shown.AddYears(direction);
                else shown = shown.AddYears(12 * direction);
            }
            catch (ArgumentOutOfRangeException) { }
        }

        DateTime Clamp(DateTime d)
        {
            if (min.HasValue && d < min.Value) return min.Value;
            if (max.HasValue && d > max.Value) return max.Value;
            return d;
        }

        // Teclas com o calendário aberto (vêm da caixa, que mantém o foco)
        public bool Key(Keys key)
        {
            int days = 0, months = 0;
            bool ctrl = (ModifierKeys & Keys.Control) != 0;
            switch (key)
            {
                case Keys.Left: days = -1; break;
                case Keys.Right: days = 1; break;
                case Keys.Up: days = -7; break;
                case Keys.Down: days = 7; break;
                case Keys.PageUp: months = ctrl ? -12 : -1; break;
                case Keys.PageDown: months = ctrl ? 12 : 1; break;
                case Keys.Home: days = 1 - focus.Day; break;
                case Keys.End: days = DateTime.DaysInMonth(focus.Year, focus.Month) - focus.Day; break;
                case Keys.Return:
                    if (Allowed(focus)) owner.Picked(focus);
                    else SystemSounds.Beep.Play();
                    return true;
                default: return false;
            }
            try { focus = Clamp(months != 0 ? focus.AddMonths(months) : focus.AddDays(days)); }
            catch (ArgumentOutOfRangeException) { }
            view = View.Days;
            shown = new DateTime(focus.Year, focus.Month, 1);
            Invalidate();
            return true;
        }

        const int GWLP_HWNDPARENT = -8;
        [DllImport("user32.dll")] static extern IntPtr SetWindowLongPtr(IntPtr hwnd, int index, IntPtr value);
    }
}
