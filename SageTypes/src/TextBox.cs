using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageTypes
{
    // TextBox: caixa de texto para UserForms, com correção ortográfica (o corretor do Windows).
    // Fica na Caixa de ferramentas (clique direito > Controles adicionais > Sage TextBox) e é
    // usada como a do MSForms:
    //
    //   Private Sub txtNome_Change()
    //       lblContador.Caption = Len(txtNome.Value)
    //   End Sub
    //
    //   txtObs.Language = "en-US"                       ' idioma da correção (padrão: o do Windows)
    //   Debug.Print txtObs.Misspellings.Join(", ")      ' palavras erradas
    //
    // As palavras erradas ficam sublinhadas em vermelho; o clique direito mostra as sugestões,
    // Ignorar e Adicionar ao dicionário (o do Windows). Left, Top, Width, Height, Visible,
    // TabIndex, Tag, ControlTipText, SetFocus e os eventos Enter, Exit, BeforeUpdate e
    // AfterUpdate vêm do próprio UserForm, como em qualquer controle.
    //
    // Interface dual: acrescente membros só no fim, com o próximo DispId (ver StringS). As
    // propriedades gravadas no arquivo usam os nomes de SaveProperties: não os mude.
    [ComVisible(true), Guid("62D4E4F3-9AED-44EC-B465-9DFD5733D62E"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _TextBox
    {
        [DispId(0)] string Value { get; set; }
        [DispId(1)] string Text { get; set; }
        [DispId(2)] bool SpellCheck { get; set; }
        [DispId(3)] string Language { get; set; }
        [DispId(4)] bool MultiLine { get; set; }
        [DispId(5)] bool WordWrap { get; set; }
        [DispId(6)] sgScrollBars ScrollBars { get; set; }
        [DispId(7)] bool Locked { get; set; }
        [DispId(-514)] bool Enabled { get; set; }
        [DispId(8)] int MaxLength { get; set; }
        [DispId(9)] string PasswordChar { get; set; }
        [DispId(-501), ComAliasName("stdole.OLE_COLOR")] int BackColor { get; set; }
        [DispId(-513), ComAliasName("stdole.OLE_COLOR")] int ForeColor { get; set; }
        [DispId(10), ComAliasName("stdole.OLE_COLOR")] int BorderColor { get; set; }
        [DispId(11)] sgBorderStyle BorderStyle { get; set; }
        [DispId(12)] string Placeholder { get; set; }
        [DispId(13)] sgTextAlign TextAlign { get; set; }
        [DispId(14)] string FontName { get; set; }
        [DispId(15)] double FontSize { get; set; }
        [DispId(16)] bool FontBold { get; set; }
        [DispId(17)] bool FontItalic { get; set; }
        [DispId(18)] int SelStart { get; set; }
        [DispId(19)] int SelLength { get; set; }
        [DispId(20)] string SelText { get; set; }
        [DispId(21)] ListS Misspellings();
        [DispId(22)] void Copy();
        [DispId(23)] void Cut();
        [DispId(24)] void Paste();
        [DispId(25)] void SelectAll();
        [DispId(26)] int TextLength { get; }
    }

    // Eventos (Private Sub txtNome_Change()...)
    [ComVisible(true), Guid("01287B64-91C2-4452-950E-EFCCEF98CA23"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface __TextBox
    {
        [DispId(1)] void Change();
        [DispId(2)] void KeyDown(ref short KeyCode, short Shift);
        [DispId(3)] void KeyUp(ref short KeyCode, short Shift);
        [DispId(4)] void KeyPress(ref short KeyAscii);
        [DispId(5)] void DblClick();
    }

    [ComVisible(true), Guid("0BAF7F9B-A8C7-4E29-8FFB-D4977E119810")]
    public enum sgScrollBars
    {
        sgScrollBarsNone = 0,
        sgScrollBarsHorizontal = 1,
        sgScrollBarsVertical = 2,
        sgScrollBarsBoth = 3,
    }

    [ComVisible(true), Guid("BCC36366-3718-491B-B194-A1E4AE38BCCF")]
    public enum sgBorderStyle
    {
        sgBorderStyleNone = 0,
        sgBorderStyleSingle = 1,
    }

    [ComVisible(true), Guid("36796066-81E1-471C-A123-5D3D007D4C2B")]
    public enum sgTextAlign
    {
        sgTextAlignLeft = 1,
        sgTextAlignCenter = 2,
        sgTextAlignRight = 3,
    }

    public delegate void TextBoxEventHandler();
    public delegate void TextBoxKeyHandler(ref short KeyCode, short Shift);
    public delegate void TextBoxKeyPressHandler(ref short KeyAscii);

    [ComVisible(true), Guid("0AD72DAF-5E3F-47CA-8C4C-3E5050C6DFC5"), ProgId("Sage.TextBox")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_TextBox)), ComSourceInterfaces(typeof(__TextBox))]
    public sealed class TextBox : OleControl, _TextBox
    {
        public event TextBoxEventHandler Change;
        public event TextBoxKeyHandler KeyDown;
        public event TextBoxKeyHandler KeyUp;
        public event TextBoxKeyPressHandler KeyPress;
        public event TextBoxEventHandler DblClick;

        // Cores OLE: as do sistema (&H80000005 = fundo de janela, &H80000008 = texto de janela)
        const int SystemWindow = unchecked((int)0x80000005), SystemWindowText = unchecked((int)0x80000008);
        const int DefaultBorder = 0xB3ADAB; // RGB(171, 173, 179), a borda das caixas de texto do Windows

        string text, language, passwordChar, placeholder, fontName;
        bool spellCheck, multiLine, wordWrap, locked, enabled, fontBold, fontItalic;
        sgScrollBars scrollBars;
        sgBorderStyle borderStyle;
        sgTextAlign textAlign;
        int maxLength, backColor, foreColor, borderColor;
        double fontSize;

        EditBox edit;
        Font font;
        List<Spelling.Error> errors = new List<Spelling.Error>();
        Timer spellTimer;
        bool applying;

        public TextBox() : base(108, 18)
        {
            InitProperties();
        }

        protected override string TypeName { get { return "Sage TextBox"; } }

        protected override void InitProperties()
        {
            text = ""; language = ""; passwordChar = ""; placeholder = "";
            fontName = "Segoe UI"; fontSize = 9;
            spellCheck = true; multiLine = false; wordWrap = true; locked = false; enabled = true; fontBold = false; fontItalic = false;
            scrollBars = sgScrollBars.sgScrollBarsNone;
            borderStyle = sgBorderStyle.sgBorderStyleSingle;
            textAlign = sgTextAlign.sgTextAlignLeft;
            maxLength = 0; backColor = SystemWindow; foreColor = SystemWindowText; borderColor = DefaultBorder;
            ResetFont();
            if (edit != null) Apply();
        }

        protected override void SaveProperties(IDictionary<string, object> p)
        {
            p["Value"] = Value;
            p["SpellCheck"] = spellCheck;
            p["Language"] = language;
            p["MultiLine"] = multiLine;
            p["WordWrap"] = wordWrap;
            p["ScrollBars"] = (int)scrollBars;
            p["Locked"] = locked;
            p["Enabled"] = enabled;
            p["MaxLength"] = maxLength;
            p["PasswordChar"] = passwordChar;
            p["BackColor"] = backColor;
            p["ForeColor"] = foreColor;
            p["BorderColor"] = borderColor;
            p["BorderStyle"] = (int)borderStyle;
            p["Placeholder"] = placeholder;
            p["TextAlign"] = (int)textAlign;
            p["FontName"] = fontName;
            p["FontSize"] = fontSize;
            p["FontBold"] = fontBold;
            p["FontItalic"] = fontItalic;
        }

        protected override void LoadProperties(IDictionary<string, object> p)
        {
            text = Get(p, "Value", text);
            spellCheck = Get(p, "SpellCheck", spellCheck);
            language = Get(p, "Language", language);
            multiLine = Get(p, "MultiLine", multiLine);
            wordWrap = Get(p, "WordWrap", wordWrap);
            scrollBars = (sgScrollBars)Get(p, "ScrollBars", (int)scrollBars);
            locked = Get(p, "Locked", locked);
            enabled = Get(p, "Enabled", enabled);
            maxLength = Get(p, "MaxLength", maxLength);
            passwordChar = Get(p, "PasswordChar", passwordChar);
            backColor = Get(p, "BackColor", backColor);
            foreColor = Get(p, "ForeColor", foreColor);
            borderColor = Get(p, "BorderColor", borderColor);
            borderStyle = (sgBorderStyle)Get(p, "BorderStyle", (int)borderStyle);
            placeholder = Get(p, "Placeholder", placeholder);
            textAlign = (sgTextAlign)Get(p, "TextAlign", (int)textAlign);
            fontName = Get(p, "FontName", fontName);
            fontSize = Get(p, "FontSize", fontSize);
            fontBold = Get(p, "FontBold", fontBold);
            fontItalic = Get(p, "FontItalic", fontItalic);
            ResetFont();
            if (edit != null) Apply();
        }

        // ------------------------------------------------------------------
        // Janela
        // ------------------------------------------------------------------

        protected override Control CreateContent()
        {
            edit = new EditBox(this);
            edit.BorderStyle = System.Windows.Forms.BorderStyle.None;
            edit.TextChanged += delegate
            {
                text = edit.Text;
                QueueSpelling();
                if (!UserMode) Changed();
                if (!applying && EventsEnabled && Change != null) Change();
            };
            edit.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (!EventsEnabled || KeyDown == null) return;
                short code = (short)e.KeyCode;
                KeyDown(ref code, Shift(e));
                if (code == 0) { e.Handled = true; e.SuppressKeyPress = true; }
            };
            edit.KeyUp += delegate(object sender, KeyEventArgs e)
            {
                if (!EventsEnabled || KeyUp == null) return;
                short code = (short)e.KeyCode;
                KeyUp(ref code, Shift(e));
                if (code == 0) e.Handled = true;
            };
            edit.KeyPress += delegate(object sender, KeyPressEventArgs e)
            {
                if (!EventsEnabled || KeyPress == null) return;
                short ascii = (short)e.KeyChar;
                KeyPress(ref ascii);
                if (ascii == 0) e.Handled = true;
                else e.KeyChar = (char)ascii;
            };
            edit.DoubleClick += delegate { if (EventsEnabled && DblClick != null) DblClick(); };
            edit.GotFocus += delegate { ContentFocus(true); if (Host != null) Host.Invalidate(); };
            edit.LostFocus += delegate { ContentFocus(false); if (Host != null) Host.Invalidate(); };
            Host.Layout += delegate { LayoutEdit(); };
            Apply();
            return edit;
        }

        static short Shift(KeyEventArgs e)
        {
            return (short)((e.Shift ? 1 : 0) | (e.Control ? 2 : 0) | (e.Alt ? 4 : 0));
        }

        void ResetFont()
        {
            FontStyle style = (fontBold ? FontStyle.Bold : FontStyle.Regular) | (fontItalic ? FontStyle.Italic : FontStyle.Regular);
            Font old = font;
            font = new Font(fontName, (float)Math.Max(1, fontSize), style, GraphicsUnit.Point);
            if (edit != null) edit.Font = font;
            if (old != null) old.Dispose();
        }

        // Propriedades na janela (quando ela existe)
        void Apply()
        {
            if (edit == null) return;
            edit.Font = font;
            edit.Multiline = multiLine;
            edit.AcceptsReturn = multiLine;
            edit.WordWrap = wordWrap;
            edit.ScrollBars = (ScrollBars)(int)scrollBars;
            edit.ReadOnly = locked;
            edit.Enabled = enabled;
            edit.MaxLength = maxLength > 0 ? maxLength : 0;
            edit.PasswordChar = passwordChar.Length > 0 ? passwordChar[0] : '\0';
            edit.TextAlign = textAlign == sgTextAlign.sgTextAlignCenter ? HorizontalAlignment.Center
                : textAlign == sgTextAlign.sgTextAlignRight ? HorizontalAlignment.Right : HorizontalAlignment.Left;
            edit.BackColor = Back;
            edit.ForeColor = ColorTranslator.FromOle(foreColor);
            if (edit.Text != text)
            {
                applying = true;
                try { edit.Text = text; }
                finally { applying = false; }
            }
            LayoutEdit();
            QueueSpelling();
            edit.Invalidate();
            if (Host != null) Host.Invalidate();
        }

        Color Back { get { return enabled ? ColorTranslator.FromOle(backColor) : SystemColors.Control; } }

        float Scale
        {
            get
            {
                if (Host == null || !Host.IsHandleCreated) return 1;
                using (Graphics g = Host.CreateGraphics()) return g.DpiX / 96f;
            }
        }

        int Border { get { return borderStyle == sgBorderStyle.sgBorderStyleSingle ? 1 : 0; } }

        void LayoutEdit()
        {
            if (edit == null || Host == null) return;
            Size size = Host.ClientSize;
            int pad = (int)Math.Round(4 * Scale) + Border;
            if (multiLine)
            {
                int top = (int)Math.Round(3 * Scale) + Border;
                edit.SetBounds(pad, top, Math.Max(1, size.Width - 2 * pad), Math.Max(1, size.Height - top - Border - 1));
            }
            else
            {
                // Uma linha: o texto fica centralizado na altura
                int h = Math.Min(edit.PreferredHeight, Math.Max(1, size.Height - 2 * Border));
                edit.SetBounds(pad, Math.Max(Border, (size.Height - h) / 2), Math.Max(1, size.Width - 2 * pad), h);
            }
        }

        // Fundo e borda (com a janela) ou tudo (sem ela: designer, impressão)
        protected override void Paint(Graphics g, Rectangle r, bool windowed)
        {
            using (SolidBrush back = new SolidBrush(Back)) g.FillRectangle(back, r);
            if (!windowed)
            {
                float scale = g.DpiX / 96f;
                int pad = (int)Math.Round(4 * scale) + Border;
                Rectangle area = new Rectangle(r.Left + pad, r.Top + Border, r.Width - 2 * pad, r.Height - 2 * Border);
                if (multiLine) { area.Y += (int)Math.Round(3 * scale); area.Height -= (int)Math.Round(3 * scale); }
                string shown = text.Length > 0 ? (passwordChar.Length > 0 ? new string(passwordChar[0], text.Length) : text) : placeholder;
                Color color = text.Length > 0 && enabled ? ColorTranslator.FromOle(foreColor) : SystemColors.GrayText;
                TextFormatFlags flags = TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis | AlignFlags() |
                    (multiLine ? (wordWrap ? TextFormatFlags.WordBreak : TextFormatFlags.Default) : TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
                TextRenderer.DrawText(g, shown, font, area, color, flags);
            }
            if (Border > 0)
            {
                bool focused = windowed && edit != null && edit.Focused;
                Color color = !enabled ? SystemColors.ControlDark : focused ? Color.FromArgb(0, 120, 212) : ColorTranslator.FromOle(borderColor);
                using (Pen pen = new Pen(color)) g.DrawRectangle(pen, r.Left, r.Top, r.Width - 1, r.Height - 1);
            }
        }

        TextFormatFlags AlignFlags()
        {
            return textAlign == sgTextAlign.sgTextAlignCenter ? TextFormatFlags.HorizontalCenter
                : textAlign == sgTextAlign.sgTextAlignRight ? TextFormatFlags.Right : TextFormatFlags.Left;
        }

        // ------------------------------------------------------------------
        // Correção ortográfica: o texto é conferido um instante depois da última tecla, e as
        // palavras erradas são sublinhadas por cima do desenho da caixa (EditBox.WndProc)
        // ------------------------------------------------------------------

        string Lang { get { return language.Length > 0 ? language : Spelling.DefaultLanguage(); } }

        bool Checking { get { return spellCheck && passwordChar.Length == 0 && UserMode; } }

        void QueueSpelling()
        {
            if (edit == null) return;
            if (spellTimer == null)
            {
                spellTimer = new Timer { Interval = 350 };
                spellTimer.Tick += delegate { spellTimer.Stop(); RunSpelling(); };
            }
            spellTimer.Stop();
            spellTimer.Start();
        }

        void RunSpelling()
        {
            errors = Checking ? Spelling.Check(Lang, text) : new List<Spelling.Error>();
            if (edit != null) edit.Invalidate();
        }

        internal void PaintSquiggles(Graphics g)
        {
            if (errors.Count == 0 || !Checking) return;
            string current = edit.Text;
            int lineHeight = edit.Font.Height;
            using (Pen pen = new Pen(Color.FromArgb(232, 17, 35)))
                foreach (Spelling.Error e in errors)
                {
                    if (e.Start + e.Length > current.Length) continue;
                    // Um trecho por linha (a palavra pode quebrar no fim da linha)
                    int segStart = e.Start;
                    Point first = edit.GetPositionFromCharIndex(segStart);
                    for (int i = e.Start + 1; i <= e.Start + e.Length; i++)
                    {
                        Point p = i < current.Length ? edit.GetPositionFromCharIndex(i) : Point.Empty;
                        bool end = i == e.Start + e.Length;
                        if (end || p.Y != first.Y)
                        {
                            int right;
                            if (!end || i >= current.Length || p.Y != first.Y)
                            {
                                Point last = edit.GetPositionFromCharIndex(i - 1);
                                right = last.X + TextRenderer.MeasureText(g, current[i - 1].ToString(), edit.Font, Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix).Width;
                            }
                            else right = p.X;
                            Wave(g, pen, first.X, right, first.Y + lineHeight - 2);
                            segStart = i;
                            first = p;
                        }
                    }
                }
        }

        static void Wave(Graphics g, Pen pen, int x1, int x2, int y)
        {
            if (x2 - x1 < 2) return;
            List<Point> points = new List<Point>();
            for (int x = x1; x <= x2; x += 2) points.Add(new Point(x, ((x - x1) / 2) % 2 == 0 ? y : y + 2));
            if (points.Count > 1) g.DrawLines(pen, points.ToArray());
        }

        internal void PaintPlaceholder(Graphics g)
        {
            if (placeholder.Length == 0 || edit.TextLength > 0) return;
            TextFormatFlags flags = TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | AlignFlags() |
                (multiLine ? TextFormatFlags.WordBreak : TextFormatFlags.SingleLine);
            Rectangle area = edit.ClientRectangle;
            area.X += 1; area.Width -= 2;
            TextRenderer.DrawText(g, placeholder, edit.Font, area, SystemColors.GrayText, flags);
        }

        // Clique direito: sugestões (se a palavra estiver errada), Ignorar, Adicionar ao
        // dicionário, e os comandos de edição
        internal void ShowMenu(Point screen)
        {
            bool pt = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "pt";
            Point client = edit.PointToClient(screen);
            int index = edit.GetCharIndexFromPosition(client);
            Spelling.Error? hit = null;
            if (Checking)
                foreach (Spelling.Error e in errors)
                    if (index >= e.Start && index <= e.Start + e.Length && e.Start + e.Length <= edit.TextLength) { hit = e; break; }

            IntPtr menu = CreatePopupMenu();
            List<string> suggestions = new List<string>();
            string word = null;
            try
            {
                if (hit.HasValue && !locked)
                {
                    word = edit.Text.Substring(hit.Value.Start, hit.Value.Length);
                    suggestions = Spelling.Suggest(Lang, word, 6);
                    for (int i = 0; i < suggestions.Count; i++) AppendMenu(menu, MF_STRING, (IntPtr)(1 + i), suggestions[i]);
                    if (suggestions.Count == 0) AppendMenu(menu, MF_STRING | MF_GRAYED, IntPtr.Zero, pt ? "(sem sugestões)" : "(no suggestions)");
                    AppendMenu(menu, MF_SEPARATOR, IntPtr.Zero, null);
                    AppendMenu(menu, MF_STRING, (IntPtr)100, pt ? "Ignorar" : "Ignore");
                    AppendMenu(menu, MF_STRING, (IntPtr)101, pt ? "Adicionar ao dicionário" : "Add to dictionary");
                    AppendMenu(menu, MF_SEPARATOR, IntPtr.Zero, null);
                }
                bool selection = edit.SelectionLength > 0, editable = !locked && enabled;
                bool password = passwordChar.Length > 0;
                AppendMenu(menu, MF_STRING | (editable && edit.CanUndo ? 0 : MF_GRAYED), (IntPtr)200, pt ? "Desfazer" : "Undo");
                AppendMenu(menu, MF_SEPARATOR, IntPtr.Zero, null);
                AppendMenu(menu, MF_STRING | (editable && selection && !password ? 0 : MF_GRAYED), (IntPtr)201, pt ? "Recortar" : "Cut");
                AppendMenu(menu, MF_STRING | (selection && !password ? 0 : MF_GRAYED), (IntPtr)202, pt ? "Copiar" : "Copy");
                AppendMenu(menu, MF_STRING | (editable && Clipboard.ContainsText() ? 0 : MF_GRAYED), (IntPtr)203, pt ? "Colar" : "Paste");
                AppendMenu(menu, MF_STRING | (editable && selection ? 0 : MF_GRAYED), (IntPtr)204, pt ? "Excluir" : "Delete");
                AppendMenu(menu, MF_SEPARATOR, IntPtr.Zero, null);
                AppendMenu(menu, MF_STRING | (edit.TextLength > 0 ? 0 : MF_GRAYED), (IntPtr)205, pt ? "Selecionar tudo" : "Select all");
                if (suggestions.Count > 0) SetMenuDefaultItem(menu, 1, 0);

                int cmd = TrackPopupMenuEx(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen.X, screen.Y, edit.Handle, IntPtr.Zero);
                if (cmd >= 1 && cmd <= suggestions.Count)
                {
                    edit.Select(hit.Value.Start, hit.Value.Length);
                    edit.SelectedText = suggestions[cmd - 1];
                }
                else if (cmd == 100) { Spelling.Ignore(Lang, word); RunSpelling(); }
                else if (cmd == 101) { Spelling.Add(Lang, word); RunSpelling(); }
                else if (cmd == 200) edit.Undo();
                else if (cmd == 201) edit.Cut();
                else if (cmd == 202) edit.Copy();
                else if (cmd == 203) edit.Paste();
                else if (cmd == 204) edit.SelectedText = "";
                else if (cmd == 205) edit.SelectAll();
            }
            finally { DestroyMenu(menu); }
        }

        internal bool Forward(ref Message m) { return ForwardKey(ref m); }
        internal bool IsUserMode { get { return UserMode; } }
        internal bool IsMultiLine { get { return multiLine; } }

        // ------------------------------------------------------------------
        // _TextBox
        // ------------------------------------------------------------------

        public string Value
        {
            get { return edit != null ? edit.Text : text; }
            set
            {
                text = value ?? "";
                if (edit != null) { if (edit.Text != text) edit.Text = text; }
                else if (EventsEnabled && Change != null) Change();
                Changed();
            }
        }

        public string Text { get { return Value; } set { Value = value; } }

        public bool SpellCheck { get { return spellCheck; } set { spellCheck = value; Changed(); RunSpelling(); } }

        public string Language
        {
            get { return language; }
            set
            {
                string v = (value ?? "").Trim();
                if (v.Length > 0 && !Spelling.IsSupported(v))
                    throw Interop.Error(5, "ValueError: o idioma '" + v + "' não está instalado no Windows (instalados: " +
                        string.Join(", ", Spelling.Languages().ToArray()) + ")");
                language = v;
                Changed();
                RunSpelling();
            }
        }

        public bool MultiLine { get { return multiLine; } set { multiLine = value; Apply(); Changed(); } }
        public bool WordWrap { get { return wordWrap; } set { wordWrap = value; Apply(); Changed(); } }
        public sgScrollBars ScrollBars { get { return scrollBars; } set { scrollBars = value; Apply(); Changed(); } }
        public bool Locked { get { return locked; } set { locked = value; Apply(); Changed(); } }
        public bool Enabled { get { return enabled; } set { enabled = value; Apply(); Changed(); } }
        public int MaxLength { get { return maxLength; } set { maxLength = Math.Max(0, value); Apply(); Changed(); } }

        public string PasswordChar
        {
            get { return passwordChar; }
            set { passwordChar = string.IsNullOrEmpty(value) ? "" : value.Substring(0, 1); Apply(); Changed(); RunSpelling(); }
        }

        public int BackColor { get { return backColor; } set { backColor = value; Apply(); Changed(); } }
        public int ForeColor { get { return foreColor; } set { foreColor = value; Apply(); Changed(); } }
        public int BorderColor { get { return borderColor; } set { borderColor = value; Apply(); Changed(); } }
        public sgBorderStyle BorderStyle { get { return borderStyle; } set { borderStyle = value; Apply(); Changed(); } }
        public string Placeholder { get { return placeholder; } set { placeholder = value ?? ""; Apply(); Changed(); } }
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

        public int SelStart
        {
            get { return edit != null ? edit.SelectionStart : 0; }
            set { if (edit != null) edit.SelectionStart = Math.Max(0, value); }
        }

        public int SelLength
        {
            get { return edit != null ? edit.SelectionLength : 0; }
            set { if (edit != null) edit.SelectionLength = Math.Max(0, value); }
        }

        public string SelText
        {
            get { return edit != null ? edit.SelectedText : ""; }
            set { if (edit != null) edit.SelectedText = value ?? ""; }
        }

        public int TextLength { get { return Value.Length; } }

        // As palavras erradas do texto, na ordem (sem repetir)
        public ListS Misspellings()
        {
            string current = Value;
            List<object> words = new List<object>();
            HashSet<string> seen = new HashSet<string>();
            if (passwordChar.Length == 0)
                foreach (Spelling.Error e in Spelling.Check(Lang, current))
                {
                    string w = current.Substring(e.Start, e.Length);
                    if (seen.Add(w)) words.Add(w);
                }
            return ListS.From(words);
        }

        public void Copy() { if (edit != null && passwordChar.Length == 0) edit.Copy(); }
        public void Cut() { if (edit != null && passwordChar.Length == 0 && !locked) edit.Cut(); }
        public void Paste() { if (edit != null && !locked) edit.Paste(); }
        public void SelectAll() { if (edit != null) edit.SelectAll(); }

        const uint MF_STRING = 0, MF_GRAYED = 1, MF_SEPARATOR = 0x800, TPM_RETURNCMD = 0x100, TPM_RIGHTBUTTON = 0x2;
        [DllImport("user32.dll")] static extern IntPtr CreatePopupMenu();
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern bool AppendMenu(IntPtr menu, uint flags, IntPtr id, string text);
        [DllImport("user32.dll")] static extern int TrackPopupMenuEx(IntPtr menu, uint flags, int x, int y, IntPtr hwnd, IntPtr tpm);
        [DllImport("user32.dll")] static extern bool DestroyMenu(IntPtr menu);
        [DllImport("user32.dll")] static extern bool SetMenuDefaultItem(IntPtr menu, uint item, uint byPosition);
    }

    // A caixa de edição (EDIT do Windows) dentro da janela do TextBox. Sem o laço de mensagens
    // do WinForms (quem lê as mensagens é o Excel), as teclas de navegação do form (Tab, Enter,
    // Esc) e o Ctrl+A são tratados aqui.
    sealed class EditBox : System.Windows.Forms.TextBox
    {
        const int WM_PAINT = 0x0F, WM_KEYDOWN = 0x100, WM_CHAR = 0x102, WM_CONTEXTMENU = 0x7B, WM_NCHITTEST = 0x84;
        const int WM_HSCROLL = 0x114, WM_VSCROLL = 0x115, WM_MOUSEWHEEL = 0x20A, HTTRANSPARENT = -1;
        readonly TextBox owner;

        public EditBox(TextBox owner) { this.owner = owner; }

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_NCHITTEST:
                    if (!owner.IsUserMode) { m.Result = (IntPtr)HTTRANSPARENT; return; }
                    break;
                case WM_KEYDOWN:
                    {
                        Keys key = (Keys)(int)m.WParam;
                        bool ctrl = (ModifierKeys & Keys.Control) != 0;
                        if (key == Keys.Tab || key == Keys.Escape || (key == Keys.Return && !owner.IsMultiLine && !ctrl))
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
                        // Sem o bipe do Windows para Tab, Esc, Ctrl+A e Enter numa linha só
                        if (c == '\t' || c == 27 || c == 1 || (c == '\r' && !owner.IsMultiLine)) return;
                        break;
                    }
                case WM_CONTEXTMENU:
                    if (owner.IsUserMode)
                    {
                        long lp = m.LParam.ToInt64();
                        Point screen = lp == -1 ? PointToScreen(GetPositionFromCharIndex(SelectionStart))
                            : new Point((short)(lp & 0xFFFF), (short)((lp >> 16) & 0xFFFF));
                        owner.ShowMenu(screen);
                        return;
                    }
                    break;
            }
            base.WndProc(ref m);
            if (m.Msg == WM_PAINT)
            {
                using (Graphics g = CreateGraphics())
                {
                    owner.PaintPlaceholder(g);
                    owner.PaintSquiggles(g);
                }
            }
            else if (m.Msg == WM_HSCROLL || m.Msg == WM_VSCROLL || m.Msg == WM_MOUSEWHEEL) Invalidate();
        }
    }
}
