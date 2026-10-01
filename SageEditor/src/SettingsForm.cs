using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace SageEditor
{
    // Tela de Configurações no estilo do VS Code. Roda numa thread própria (com seu
    // loop de mensagens), para não depender do loop do Excel e não travar o VBE.
    sealed class SettingsForm : Form
    {
        static readonly object sync = new object();
        static SettingsForm instance;

        public static void ShowWindow(Action<Theme> applyTheme)
        {
            lock (sync)
            {
                if (instance != null)
                {
                    SettingsForm open = instance;
                    open.BeginInvoke((MethodInvoker)delegate
                    {
                        if (open.WindowState == FormWindowState.Minimized) open.WindowState = FormWindowState.Normal;
                        open.Activate();
                    });
                    return;
                }
                Thread thread = new Thread(delegate()
                {
                    try
                    {
                        Log.Info("Configurações: criando");
                        SettingsForm form = new SettingsForm(applyTheme);
                        form.Shown += delegate { Log.Info("Configurações: exibida " + form.Handle + " '" + form.Text + "'"); };
                        lock (sync) instance = form;
                        Application.Run(form);
                    }
                    catch (Exception ex) { Log.Error(ex); }
                    finally { lock (sync) instance = null; }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Name = "Sage.Settings";
                thread.Start();
            }
        }

        public static void CloseWindow()
        {
            lock (sync)
            {
                if (instance != null)
                {
                    SettingsForm open = instance;
                    open.BeginInvoke((MethodInvoker)open.Close);
                }
            }
        }

        // ------------------------------------------------------------------

        readonly Action<Theme> applyTheme;
        readonly Font uiFont = new Font("Segoe UI", 9.75f);
        readonly Font titleFont = new Font("Segoe UI", 9.75f, FontStyle.Bold);
        readonly Font sectionFont = new Font("Segoe UI Semibold", 14f);

        readonly Panel searchBox = new Panel();
        readonly TextBox search = new TextBox();
        readonly Label userTab = new Label();
        readonly Panel userTabLine = new Panel();
        readonly Panel separator = new Panel();
        readonly Panel nav = new Panel();
        readonly List<Label> navItems = new List<Label>();
        readonly Panel content = new Panel();
        readonly Label sectionTitle = new Label();
        readonly SettingItem themeItem;
        readonly Label editorTitle = new Label();
        readonly SettingItem lineItem, tabsItem;
        // Resultado da pesquisa (Control.Visible é falso enquanto a janela não aparece)
        bool showAppearance = true, showLines = true, showTabs = true;
        readonly Label noResults = new Label();
        Theme theme;

        SettingsForm(Action<Theme> applyTheme)
        {
            this.applyTheme = applyTheme;
            theme = Theme.Find(Settings.ColorTheme);

            Text = Strings.SettingsTitle;
            Font = uiFont;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(960, 620);
            MinimumSize = new Size(640, 400);
            ShowIcon = false;
            KeyPreview = true;
            KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) Close(); };

            // Barra de pesquisa
            searchBox.Padding = new Padding(8, 5, 8, 4);
            search.BorderStyle = BorderStyle.None;
            search.Dock = DockStyle.Fill;
            search.TextChanged += delegate { Filter(); };
            searchBox.Controls.Add(search);
            Controls.Add(searchBox);

            // Aba "Usuário"
            userTab.Text = Strings.UserTab;
            userTab.AutoSize = true;
            Controls.Add(userTab);
            Controls.Add(userTabLine);
            Controls.Add(separator);

            // Índice à esquerda
            AddNavItem(Strings.NavCommon);
            AddNavItem(Strings.NavAppearance);
            AddNavItem(Strings.NavEditor);
            Controls.Add(nav);

            // Conteúdo
            sectionTitle.Text = Strings.NavAppearance;
            sectionTitle.Font = sectionFont;
            sectionTitle.AutoSize = true;
            content.Controls.Add(sectionTitle);

            List<string> names = new List<string>();
            foreach (Theme t in Theme.All) names.Add(t.DisplayName);
            themeItem = new SettingItem(Strings.AppearanceCategory, Strings.ColorTheme, Strings.ColorThemeDescription,
                names.ToArray(), theme.DisplayName, Strings.ColorThemeKeywords);
            themeItem.ValueChanged += OnThemeChanged;
            content.Controls.Add(themeItem);

            editorTitle.Text = Strings.NavEditor;
            editorTitle.Font = sectionFont;
            editorTitle.AutoSize = true;
            content.Controls.Add(editorTitle);

            string[] onOff = { Strings.On, Strings.Off };
            lineItem = new SettingItem(Strings.EditorCategory, Strings.LineNumbers, Strings.LineNumbersDescription,
                onOff, Settings.LineNumbers ? Strings.On : Strings.Off, Strings.LineNumbersKeywords);
            lineItem.ValueChanged += delegate(string value)
            {
                Settings.LineNumbers = value == Strings.On;
                applyTheme(theme); // redesenha o VBE
            };
            content.Controls.Add(lineItem);

            tabsItem = new SettingItem(Strings.EditorCategory, Strings.ShowTabs, Strings.ShowTabsDescription,
                onOff, Settings.EditorTabs ? Strings.On : Strings.Off, Strings.ShowTabsKeywords);
            tabsItem.ValueChanged += delegate(string value)
            {
                Settings.EditorTabs = value == Strings.On;
                applyTheme(theme); // redesenha o VBE
            };
            content.Controls.Add(tabsItem);

            noResults.Text = Strings.NoResults;
            noResults.AutoSize = true;
            noResults.Visible = false;
            content.Controls.Add(noResults);

            content.AutoScroll = true;
            Controls.Add(content);

            Resize += delegate { DoLayout(); };
            DoLayout();
            ApplyColors();
            SelectNav(navItems[1]);
        }

        void AddNavItem(string text)
        {
            Label item = new Label();
            item.Text = text;
            item.AutoSize = false;
            item.Height = 26;
            item.TextAlign = ContentAlignment.MiddleLeft;
            item.Padding = new Padding(16, 0, 0, 0);
            item.Cursor = Cursors.Hand;
            item.Click += delegate { SelectNav(item); };
            navItems.Add(item);
            nav.Controls.Add(item);
        }

        void SelectNav(Label selected)
        {
            foreach (Label l in navItems)
            {
                bool on = l == selected;
                l.Font = on ? titleFont : uiFont;
                l.ForeColor = on ? theme.Foreground : theme.Muted;
            }
            if (selected == navItems[2]) { content.ScrollControlIntoView(lineItem); lineItem.Focus(); }
            else { content.ScrollControlIntoView(sectionTitle); themeItem.Focus(); }
        }

        void DoLayout()
        {
            int w = ClientSize.Width, h = ClientSize.Height;
            int margin = 24;
            searchBox.SetBounds(margin, 16, w - 2 * margin, 28);
            userTab.Location = new Point(margin, 56);
            userTabLine.SetBounds(margin, userTab.Bottom + 4, userTab.Width, 2);
            separator.SetBounds(margin, userTabLine.Bottom, w - 2 * margin, 1);

            int top = separator.Bottom + 12;
            nav.SetBounds(margin, top, 200, h - top);
            for (int i = 0; i < navItems.Count; i++)
                navItems[i].SetBounds(0, i * 26, nav.Width, 26);

            content.SetBounds(nav.Right + 16, top, w - nav.Right - 16 - margin, h - top);
            int itemWidth = Math.Min(content.ClientSize.Width, 820);
            int y = 0;
            if (showAppearance)
            {
                sectionTitle.Location = new Point(12, y);
                themeItem.SetBounds(0, sectionTitle.Bottom + 12, itemWidth, themeItem.PreferredHeight);
                y = themeItem.Bottom + 24;
            }
            if (showLines || showTabs)
            {
                editorTitle.Location = new Point(12, y);
                y = editorTitle.Bottom + 12;
                if (showLines)
                {
                    lineItem.SetBounds(0, y, itemWidth, lineItem.PreferredHeight);
                    y = lineItem.Bottom + 12;
                }
                if (showTabs)
                    tabsItem.SetBounds(0, y, itemWidth, tabsItem.PreferredHeight);
            }
            noResults.Location = new Point(12, 4);
        }

        void Filter()
        {
            string q = search.Text.Trim();
            showAppearance = themeItem.Matches(q);
            showLines = lineItem.Matches(q);
            showTabs = tabsItem.Matches(q);
            themeItem.Visible = sectionTitle.Visible = showAppearance;
            lineItem.Visible = showLines;
            tabsItem.Visible = showTabs;
            editorTitle.Visible = showLines || showTabs;
            noResults.Visible = !showAppearance && !showLines && !showTabs;
            DoLayout();
        }

        void OnThemeChanged(string name)
        {
            Theme selected = Theme.Find(name);
            Settings.ColorTheme = selected.Name;
            theme = selected;
            ApplyColors();
            applyTheme(selected);
        }

        void ApplyColors()
        {
            BackColor = theme.Background;
            ForeColor = theme.Foreground;
            searchBox.BackColor = theme.Input;
            search.BackColor = theme.Input;
            search.ForeColor = theme.Foreground;
            searchBox.Paint -= PaintSearchBorder;
            searchBox.Paint += PaintSearchBorder;
            userTab.ForeColor = theme.Foreground;
            userTabLine.BackColor = theme.Accent;
            separator.BackColor = theme.Border;
            nav.BackColor = theme.Background;
            content.BackColor = theme.Background;
            Native.SetWindowTheme(content.Handle, theme.IsDark ? "DarkMode_Explorer" : null, null); // barras de rolagem
            sectionTitle.ForeColor = theme.Foreground;
            editorTitle.ForeColor = theme.Foreground;
            lineItem.SetTheme(theme);
            tabsItem.SetTheme(theme);
            noResults.ForeColor = theme.Muted;
            foreach (Label l in navItems) l.BackColor = theme.Background;
            themeItem.SetTheme(theme);
            SelectNavColors();
            UseDarkTitleBar(theme.IsDark);
            searchBox.Invalidate();
        }

        void SelectNavColors()
        {
            foreach (Label l in navItems)
                l.ForeColor = l.Font.Bold ? theme.Foreground : theme.Muted;
        }

        void PaintSearchBorder(object sender, PaintEventArgs e)
        {
            using (Pen pen = new Pen(search.Focused ? theme.Accent : theme.Border))
                e.Graphics.DrawRectangle(pen, 0, 0, searchBox.Width - 1, searchBox.Height - 1);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            const int EM_SETCUEBANNER = 0x1501;
            Native.SendMessage(search.Handle, EM_SETCUEBANNER, (IntPtr)1,
                System.Runtime.InteropServices.Marshal.StringToHGlobalUni(Strings.SearchCue));
            search.GotFocus += delegate { searchBox.Invalidate(); };
            search.LostFocus += delegate { searchBox.Invalidate(); };
            UseDarkTitleBar(theme.IsDark);
        }

        void UseDarkTitleBar(bool dark)
        {
            if (!IsHandleCreated) return;
            int on = dark ? 1 : 0;
            Native.DwmSetWindowAttribute(Handle, Native.DWMWA_USE_IMMERSIVE_DARK_MODE, ref on, 4);
            int caption = theme.IsDefault ? Native.DWMWA_COLOR_DEFAULT : ColorRef(theme.Sidebar);
            int text = theme.IsDefault ? Native.DWMWA_COLOR_DEFAULT : ColorRef(theme.Foreground);
            Native.DwmSetWindowAttribute(Handle, Native.DWMWA_CAPTION_COLOR, ref caption, 4);
            Native.DwmSetWindowAttribute(Handle, Native.DWMWA_TEXT_COLOR, ref text, 4);
        }

        static int ColorRef(Color c)
        {
            return c.R | (c.G << 8) | (c.B << 16);
        }
    }

    // Um item de configuração: "Categoria: Nome", descrição e um seletor.
    sealed class SettingItem : Panel
    {
        readonly string category, name, description, keywords;
        readonly ComboBox combo = new ComboBox();
        readonly Font bold = new Font("Segoe UI", 9.75f, FontStyle.Bold);
        Theme theme;
        bool active;

        public event Action<string> ValueChanged;

        public SettingItem(string category, string name, string description, string[] options, string value, string keywords)
        {
            this.category = category;
            this.name = name;
            this.description = description;
            this.keywords = (category + " " + name + " " + description + " " + keywords).ToLowerInvariant();
            DoubleBuffered = true;

            combo.DropDownStyle = ComboBoxStyle.DropDownList;
            combo.FlatStyle = FlatStyle.Standard;
            combo.DrawMode = DrawMode.OwnerDrawFixed;
            combo.ItemHeight = 20;
            combo.Width = 320;
            combo.Items.AddRange(options);
            combo.SelectedItem = value;
            combo.DrawItem += DrawComboItem;
            combo.SelectedIndexChanged += delegate
            {
                if (ValueChanged != null) ValueChanged((string)combo.SelectedItem);
            };
            combo.GotFocus += delegate { SetActive(true); };
            combo.LostFocus += delegate { SetActive(false); };
            Controls.Add(combo);

            MouseEnter += delegate { Invalidate(); };
            MouseLeave += delegate { Invalidate(); };
            Resize += delegate { Relayout(); };
        }

        public int PreferredHeight { get { return DescriptionTop + DescriptionHeight + 8 + combo.Height + 16; } }
        int DescriptionTop { get { return 34; } }
        int DescriptionHeight
        {
            get
            {
                int w = Math.Max(100, Width - 40);
                return TextRenderer.MeasureText(description, Font, new Size(w, 0), TextFormatFlags.WordBreak).Height;
            }
        }

        public bool Matches(string query)
        {
            if (query.Length == 0) return true;
            foreach (string word in query.ToLowerInvariant().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                if (!keywords.Contains(word)) return false;
            return true;
        }

        public new void Focus() { combo.Focus(); }

        public void SetTheme(Theme t)
        {
            theme = t;
            BackColor = t.Background;
            combo.BackColor = t.Input;
            combo.ForeColor = t.Foreground;
            // seta e borda escuras do próprio Windows
            Native.SetWindowTheme(combo.Handle, t.IsDark ? "DarkMode_CFD" : null, null);
            Invalidate();
            combo.Invalidate();
        }

        void SetActive(bool value)
        {
            active = value;
            Invalidate();
        }

        void Relayout()
        {
            combo.Location = new Point(20, DescriptionTop + DescriptionHeight + 8);
            Height = PreferredHeight;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (theme == null) return;
            Graphics g = e.Graphics;
            bool hover = ClientRectangle.Contains(PointToClient(Cursor.Position));
            if (active || hover)
            {
                using (SolidBrush b = new SolidBrush(theme.Hover)) g.FillRectangle(b, ClientRectangle);
                if (active)
                    using (SolidBrush b = new SolidBrush(theme.Accent)) g.FillRectangle(b, 0, 0, 2, Height);
            }

            // "Categoria: " normal + "Nome" em negrito, como no VS Code
            TextFormatFlags flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine;
            Size catSize = TextRenderer.MeasureText(g, category, Font, Size.Empty, flags);
            TextRenderer.DrawText(g, category, Font, new Point(20, 12), theme.Foreground, flags);
            TextRenderer.DrawText(g, name, bold, new Point(20 + catSize.Width, 12), theme.Foreground, flags);

            Rectangle desc = new Rectangle(20, DescriptionTop, Width - 40, DescriptionHeight);
            TextRenderer.DrawText(g, description, Font, desc, theme.Muted, TextFormatFlags.WordBreak);
        }

        void DrawComboItem(object sender, DrawItemEventArgs e)
        {
            if (theme == null || e.Index < 0) return;
            bool selected = (e.State & DrawItemState.Selected) != 0 && (e.State & DrawItemState.ComboBoxEdit) == 0;
            Color bg = selected ? theme.Accent : theme.Input;
            Color fg = selected ? Color.White : theme.Foreground;
            using (SolidBrush b = new SolidBrush(bg)) e.Graphics.FillRectangle(b, e.Bounds);
            Rectangle r = new Rectangle(e.Bounds.X + 6, e.Bounds.Y, e.Bounds.Width - 6, e.Bounds.Height);
            TextRenderer.DrawText(e.Graphics, (string)combo.Items[e.Index], combo.Font, r, fg,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
        }
    }
}
