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
        readonly List<NavNode> navNodes = new List<NavNode>();
        NavNode selectedNode;
        readonly Panel content = new Panel();
        readonly Label sectionTitle = new Label();
        readonly SettingItem themeItem;
        readonly Label featuresTitle = new Label();
        // Os itens de cada seção (Editor > Aparência e Editor > Recursos), na ordem da tela
        readonly List<SettingItem> appearanceItems = new List<SettingItem>();
        readonly List<SettingItem> featureItems = new List<SettingItem>();
        List<SettingItem> group; // a seção que recebe os próximos AddOnOff/AddChoice
        // Resultado da pesquisa (Control.Visible é falso enquanto a janela não aparece)
        readonly HashSet<SettingItem> hidden = new HashSet<SettingItem>();
        readonly Label noResults = new Label();
        readonly SettingItem referenceItem;
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

            // Índice à esquerda, em árvore: Mais Usados; Editor > Aparência e Recursos
            NavNode common = AddNavNode(Strings.NavCommon, null, false);
            NavNode editor = AddNavNode(Strings.NavEditor, null, true);
            NavNode appearance = AddNavNode(Strings.NavAppearance, editor, false);
            NavNode features = AddNavNode(Strings.NavFeatures, editor, false);
            Controls.Add(nav);

            // Editor > Aparência: o que muda o visual do editor
            sectionTitle.Text = Strings.NavAppearance;
            sectionTitle.Font = sectionFont;
            sectionTitle.AutoSize = true;
            content.Controls.Add(sectionTitle);

            List<string> names = new List<string>();
            foreach (Theme t in Theme.All) names.Add(t.DisplayName);
            themeItem = new SettingItem(Strings.AppearanceCategory, Strings.ColorTheme, Strings.ColorThemeDescription,
                names.ToArray(), theme.DisplayName, Strings.ColorThemeKeywords);
            themeItem.ValueChanged += OnThemeChanged;
            appearanceItems.Add(themeItem);
            content.Controls.Add(themeItem);

            group = appearanceItems;
            AddOnOff(Strings.LineNumbers, Strings.LineNumbersDescription, Strings.LineNumbersKeywords, Settings.LineNumbers,
                delegate(bool on) { Settings.LineNumbers = on; applyTheme(theme); }); // redesenha o VBE
            AddOnOff(Strings.ShowTabs, Strings.ShowTabsDescription, Strings.ShowTabsKeywords, Settings.EditorTabs,
                delegate(bool on) { Settings.EditorTabs = on; applyTheme(theme); });
            AddOnOff(Strings.Icons, Strings.IconsDescription, Strings.IconsKeywords, Settings.Icons,
                delegate(bool on) { Settings.Icons = on; applyTheme(theme); }); // redesenha as abas
            AddOnOff(Strings.ColorDecorators, Strings.ColorDecoratorsDescription, Strings.ColorDecoratorsKeywords, Settings.ColorDecorators,
                delegate(bool on) { Settings.ColorDecorators = on; applyTheme(theme); }); // redesenha o código
            AddOnOff(Strings.Explorer, Strings.ExplorerDescription, Strings.ExplorerKeywords, Settings.UnifiedExplorer,
                delegate(bool on) { Settings.UnifiedExplorer = on; });

            // Editor > Recursos: o que o editor faz
            featuresTitle.Text = Strings.NavFeatures;
            featuresTitle.Font = sectionFont;
            featuresTitle.AutoSize = true;
            content.Controls.Add(featuresTitle);

            group = featureItems;
            AddOnOff(Strings.MultiCursor, Strings.MultiCursorDescription, Strings.MultiCursorKeywords, Settings.MultiCursor,
                delegate(bool on) { Settings.MultiCursor = on; });
            AddChoice(Strings.Errors, Strings.ErrorsDescription, Strings.ErrorsKeywords,
                new string[] { Strings.ErrorsStop, Strings.ErrorsDebug, Strings.Off },
                Settings.Errors == Settings.ErrorMode.Debug ? Strings.ErrorsDebug : Settings.Errors == Settings.ErrorMode.Off ? Strings.Off : Strings.ErrorsStop,
                delegate(string selected)
                {
                    Settings.Errors = selected == Strings.ErrorsDebug ? Settings.ErrorMode.Debug : selected == Strings.Off ? Settings.ErrorMode.Off : Settings.ErrorMode.Stop;
                });
            AddOnOff(Strings.AutoReference, Strings.AutoReferenceDescription, Strings.AutoReferenceKeywords, Settings.AutoReference,
                delegate(bool on) { Settings.AutoReference = on; });
            // A referência no projeto ativo: relida quando a janela volta ao foco (o projeto ativo muda)
            referenceItem = AddOnOff(Strings.SageReference, string.Format(Strings.SageReferenceDescription, Strings.SageReferenceNone),
                Strings.SageReferenceKeywords, false, OnReferenceChanged);
            RefreshReference();
            Activated += delegate { RefreshReference(); };
            AddOnOff(Strings.ToolboxControls, Strings.ToolboxControlsDescription, Strings.ToolboxControlsKeywords, ToolboxControls.On,
                delegate(bool on) { ToolboxControls.Set(on); });

            noResults.Text = Strings.NoResults;
            noResults.AutoSize = true;
            noResults.Visible = false;
            content.Controls.Add(noResults);

            content.AutoScroll = true;
            Controls.Add(content);

            common.Target = sectionTitle;
            editor.Target = sectionTitle;
            appearance.Target = sectionTitle;
            features.Target = featuresTitle;

            Resize += delegate { DoLayout(); };
            DoLayout();
            ApplyColors();
            SelectNav(appearance);
        }

        IEnumerable<SettingItem> AllItems()
        {
            foreach (SettingItem item in appearanceItems) yield return item;
            foreach (SettingItem item in featureItems) yield return item;
        }

        // Item "Editor: <nome>" com Ativado/Desativado
        SettingItem AddOnOff(string name, string description, string keywords, bool value, Action<bool> changed)
        {
            SettingItem item = new SettingItem(Strings.EditorCategory, name, description,
                new string[] { Strings.On, Strings.Off }, value ? Strings.On : Strings.Off, keywords);
            item.ValueChanged += delegate(string selected) { changed(selected == Strings.On); };
            group.Add(item);
            content.Controls.Add(item);
            return item;
        }

        // Item "Editor: <nome>" com uma lista de opções
        SettingItem AddChoice(string name, string description, string keywords, string[] options, string value, Action<string> changed)
        {
            SettingItem item = new SettingItem(Strings.EditorCategory, name, description, options, value, keywords);
            item.ValueChanged += delegate(string selected) { changed(selected); };
            group.Add(item);
            content.Controls.Add(item);
            return item;
        }

        // Estado da referência ao Sage no projeto ativo (lido na thread do Excel)
        void RefreshReference()
        {
            try
            {
                string project = null;
                SageReference.Status status = Connect.OnExcel(delegate(dynamic vbe)
                {
                    string p;
                    SageReference.Status st = SageReference.Get(vbe, out p);
                    project = p;
                    return st;
                });
                referenceItem.SetDescription(string.Format(Strings.SageReferenceDescription,
                    string.IsNullOrEmpty(project) ? Strings.SageReferenceNone : project));
                referenceItem.SetValue(status == SageReference.Status.Present ? Strings.On : Strings.Off);
                DoLayout();
            }
            catch (Exception ex) { Log.Error(ex); }
        }

        void OnReferenceChanged(bool on)
        {
            string error;
            try { error = Connect.OnExcel(delegate(dynamic vbe) { return SageReference.Set(vbe, on); }); }
            catch (Exception ex) { error = ex.Message; }
            if (error != null) MessageBox.Show(this, error, Strings.SettingsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            RefreshReference();
        }

        // Um item do índice: um grupo (Editor, que abre e fecha) ou uma seção (vai até o título dela)
        sealed class NavNode
        {
            public Label Label;
            public string Text;
            public NavNode Parent;
            public bool IsGroup, Expanded = true;
            public Control Target;
            public bool Shown { get { return Parent == null || Parent.Expanded; } }
        }

        NavNode AddNavNode(string text, NavNode parent, bool isGroup)
        {
            NavNode node = new NavNode { Text = text, Parent = parent, IsGroup = isGroup };
            Label label = new Label();
            label.AutoSize = false;
            label.Height = 26;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.Padding = new Padding(parent == null ? 4 : 28, 0, 0, 0); // os filhos recuados
            label.Cursor = Cursors.Hand;
            label.Click += delegate { SelectNav(node); };
            node.Label = label;
            navNodes.Add(node);
            nav.Controls.Add(label);
            NavText(node);
            return node;
        }

        // Os grupos com a seta (aberto ou fechado); os outros alinhados ao texto do grupo
        static void NavText(NavNode node)
        {
            node.Label.Text = (node.IsGroup ? (node.Expanded ? "\u25BE  " : "\u25B8  ") : node.Parent == null ? "     " : "") + node.Text;
        }

        void SelectNav(NavNode node)
        {
            if (node.IsGroup)
            {
                node.Expanded = !node.Expanded || selectedNode == null;
                NavText(node);
                DoLayout();
            }
            selectedNode = node;
            SelectNavColors();
            ScrollTo(node.Target);
        }

        // A seção no topo da área de conteúdo, com o foco no primeiro item dela
        void ScrollTo(Control title)
        {
            if (title == null || !title.Visible) return;
            content.AutoScrollPosition = new Point(0, title.Top - content.AutoScrollPosition.Y);
            List<SettingItem> items = title == featuresTitle ? featureItems : appearanceItems;
            foreach (SettingItem item in items) if (!hidden.Contains(item)) { item.Focus(); break; }
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
            int navY = 0;
            foreach (NavNode node in navNodes)
            {
                node.Label.Visible = node.Shown;
                if (!node.Shown) continue;
                node.Label.SetBounds(0, navY, nav.Width, 26);
                navY += 26;
            }

            content.SetBounds(nav.Right + 16, top, w - nav.Right - 16 - margin, h - top);
            int itemWidth = Math.Min(content.ClientSize.Width, 820);
            int y = content.AutoScrollPosition.Y; // as posições são da área visível, que pode estar rolada
            y = LayoutSection(sectionTitle, appearanceItems, y, itemWidth);
            y = LayoutSection(featuresTitle, featureItems, y, itemWidth);
            noResults.Location = new Point(12, 4);
        }

        // O título e os itens visíveis de uma seção; sem nenhum item (pesquisa), nem o título
        int LayoutSection(Label title, List<SettingItem> items, int y, int itemWidth)
        {
            bool any = items.Exists(i => !hidden.Contains(i));
            title.Visible = any;
            if (!any) return y;
            title.Location = new Point(12, y);
            y = title.Bottom + 12;
            foreach (SettingItem item in items)
            {
                if (hidden.Contains(item)) continue;
                item.SetBounds(0, y, itemWidth, item.PreferredHeight);
                y = item.Bottom + 12;
            }
            return y + 12;
        }

        void Filter()
        {
            string q = search.Text.Trim();
            hidden.Clear();
            int total = 0;
            foreach (SettingItem item in AllItems())
            {
                total++;
                bool match = item.Matches(q);
                item.Visible = match;
                if (!match) hidden.Add(item);
            }
            noResults.Visible = hidden.Count == total;
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
            featuresTitle.ForeColor = theme.Foreground;
            foreach (SettingItem item in AllItems()) item.SetTheme(theme);
            noResults.ForeColor = theme.Muted;
            foreach (NavNode node in navNodes) node.Label.BackColor = theme.Background;
            SelectNavColors();
            UseDarkTitleBar(theme.IsDark);
            searchBox.Invalidate();
        }

        void SelectNavColors()
        {
            foreach (NavNode node in navNodes)
            {
                bool on = node == selectedNode;
                node.Label.Font = on ? titleFont : uiFont;
                node.Label.ForeColor = on ? theme.Foreground : theme.Muted;
            }
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
        readonly string category, name, keywords;
        string description;
        bool quiet;
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
                if (!quiet && ValueChanged != null) ValueChanged((string)combo.SelectedItem);
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

        // Muda o valor mostrado sem disparar ValueChanged
        public void SetValue(string value)
        {
            if (Equals(combo.SelectedItem, value)) return;
            quiet = true;
            try { combo.SelectedItem = value; }
            finally { quiet = false; }
        }

        public void SetDescription(string text)
        {
            if (description == text) return;
            description = text;
            Relayout();
            Invalidate();
        }

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
