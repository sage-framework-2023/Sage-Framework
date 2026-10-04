// Sage Framework Installer, no estilo do Visual Studio Installer.
// Compilado em memória por SageInstaller.ps1 (C# 5). O Sage Framework é instalado
// inteiro: cada componente usa o install.ps1 da própria pasta, que continua sendo
// a fonte da verdade.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace SageInstaller
{
    public static class Program
    {
        public static void Run(string root)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new InstallerForm(root));
        }
    }

    // ------------------------------------------------------------------
    // Componentes do Sage Framework
    // ------------------------------------------------------------------
    abstract class Component
    {
        public string Name, Folder;
        public bool NeedsExcelClosed;
        public string Script { get { return Path.Combine(Folder, "install.ps1"); } }
        public abstract bool Installed { get; }
        public virtual bool UpdateAvailable { get { return false; } }
        public virtual DateTime InstalledAt { get { return DateTime.MinValue; } }
    }

    // Componente compilado pelo install.ps1 (src\*.cs) num DLL em %LOCALAPPDATA%\Sage
    abstract class CompiledComponent : Component
    {
        protected string Dll;

        protected CompiledComponent(string root, string folder, string dll)
        {
            Folder = Path.Combine(root, folder);
            Dll = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), dll);
            NeedsExcelClosed = true;
        }

        protected static bool KeyExists(string path)
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(path)) return k != null;
        }

        public override DateTime InstalledAt { get { return Installed ? File.GetLastWriteTime(Dll) : DateTime.MinValue; } }

        // Há código-fonte mais novo que o DLL instalado
        public override bool UpdateAvailable
        {
            get
            {
                if (!Installed) return false;
                DateTime installed = File.GetLastWriteTime(Dll);
                foreach (string f in Directory.GetFiles(Path.Combine(Folder, "src"), "*.cs"))
                    if (File.GetLastWriteTime(f) > installed) return true;
                return false;
            }
        }
    }

    sealed class SageEditor : CompiledComponent
    {
        public SageEditor(string root) : base(root, "SageEditor", @"Sage\Editor\SageEditor.dll")
        {
            Name = "Menu e temas do VBE";
        }

        public override bool Installed
        {
            get { return KeyExists(@"Software\Microsoft\VBA\VBE\6.0\Addins64\Sage.Editor") && File.Exists(Dll); }
        }
    }

    // Tipos para o VBA (StringS...): biblioteca "Sage" em Ferramentas > Referências
    sealed class SageTypes : CompiledComponent
    {
        public SageTypes(string root) : base(root, "SageTypes", @"Sage\Types\SageTypes.dll")
        {
            Name = "Tipos para o VBA";
        }

        public override bool Installed
        {
            get { return KeyExists(@"Software\Classes\TypeLib\{3F8E2A61-7C4B-4E9D-A215-6B0C9D8E7F14}") && File.Exists(Dll); }
        }
    }

    sealed class SageShortcuts : Component
    {
        static readonly string Link = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Startup), "SageShortcuts.lnk");

        public SageShortcuts(string root)
        {
            Name = "Atalhos de teclado do VBE";
            Folder = Path.Combine(root, "SageShortcuts");
        }

        public override bool Installed { get { return File.Exists(Link); } }
        public override DateTime InstalledAt { get { return Installed ? File.GetLastWriteTime(Link) : DateTime.MinValue; } }

        public static bool Running
        {
            get
            {
                EventWaitHandle ev;
                if (!EventWaitHandle.TryOpenExisting(@"Local\SageShortcuts.Stop", out ev)) return false;
                ev.Dispose();
                return true;
            }
        }

        public void Start()
        {
            ProcessStartInfo psi = new ProcessStartInfo("powershell.exe",
                "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + Path.Combine(Folder, "SageShortcuts.ps1") + "\"");
            psi.WindowStyle = ProcessWindowStyle.Hidden;
            psi.CreateNoWindow = true;
            Process.Start(psi);
        }
    }

    sealed class Framework
    {
        public readonly SageEditor Addin;
        public readonly SageShortcuts Shortcuts;
        public readonly List<Component> Components = new List<Component>();
        public readonly string Root;

        public Framework(string root)
        {
            Root = root;
            Addin = new SageEditor(root);
            Shortcuts = new SageShortcuts(root);
            Components.Add(Addin);
            Components.Add(Shortcuts);
            Components.Add(new SageTypes(root));
        }

        public const string Name = "Sage Framework";
        public const string Description =
            "Menu Sage e temas de cores para todo o editor do VBA, abas das janelas abertas, tela de Configurações " +
            "no estilo do VS Code, atalhos de teclado (Ctrl+K, Ctrl+C para comentar) e tipos para o VBA no estilo do Python (StringS, ListS, DictionaryS, DataFrame, DateTimeS, Json, Requests e SqlEngine).";

        public bool AnyInstalled { get { foreach (Component c in Components) if (c.Installed) return true; return false; } }
        public bool AllInstalled { get { foreach (Component c in Components) if (!c.Installed) return false; return true; } }
        public bool Incomplete { get { return AnyInstalled && !AllInstalled; } }
        public bool UpdateAvailable { get { foreach (Component c in Components) if (c.UpdateAvailable) return true; return false; } }
        public bool NeedsExcelClosed { get { return true; } }

        public string Status
        {
            get
            {
                if (!AnyInstalled) return "Não instalado";
                if (Incomplete)
                {
                    List<string> missing = new List<string>();
                    foreach (Component c in Components) if (!c.Installed) missing.Add(c.Name.ToLowerInvariant());
                    return "Instalação incompleta: falta " + string.Join(" e ", missing.ToArray());
                }
                DateTime at = DateTime.MinValue;
                foreach (Component c in Components) if (c.InstalledAt > at) at = c.InstalledAt;
                return "Instalado em " + at.ToString("dd/MM/yyyy HH:mm") +
                    " · atalhos " + (SageShortcuts.Running ? "em execução" : "parados");
            }
        }
    }

    // ------------------------------------------------------------------
    // Janela
    // ------------------------------------------------------------------
    sealed class InstallerForm : Form
    {
        internal const string Title = "Sage Framework Installer";
        internal static readonly Color Back = Color.FromArgb(0xF3, 0xF4, 0xF1);
        internal static readonly Color Surface = Color.White;
        internal static readonly Color Border = Color.FromArgb(0xDE, 0xE2, 0xDA);
        internal static readonly Color Text1 = Color.FromArgb(0x1E, 0x22, 0x1C);
        internal static readonly Color Text2 = Color.FromArgb(0x5E, 0x66, 0x5A);
        internal static readonly Color Accent = Color.FromArgb(0x5F, 0x7A, 0x58);       // sálvia escura
        internal static readonly Color AccentHover = Color.FromArgb(0x4E, 0x66, 0x49);
        internal static readonly Font UiFont = new Font("Segoe UI", 9.75f);
        internal static Icon AppIcon;

        readonly Framework framework;
        readonly Label installedTab = new Label(), availableTab = new Label();
        readonly Panel tabLine = new Panel();
        readonly FlowLayoutPanel list = new FlowLayoutPanel();
        readonly Label empty = new Label();
        bool showInstalled;
        bool busy;

        public InstallerForm(string root)
        {
            framework = new Framework(root);
            showInstalled = framework.AnyInstalled;

            string iconFile = Path.Combine(Path.Combine(root, "SageInstaller"), "sage.ico");
            if (File.Exists(iconFile)) AppIcon = new Icon(iconFile);

            Text = Title;
            Font = UiFont;
            BackColor = Back;
            ClientSize = new Size(960, 560);
            MinimumSize = new Size(760, 440);
            StartPosition = FormStartPosition.CenterScreen;
            if (AppIcon != null) Icon = AppIcon;

            // Cabeçalho branco com título e abas, como no VS Installer
            Panel header = new Panel { Dock = DockStyle.Top, Height = 96, BackColor = Surface };
            header.Paint += delegate(object s, PaintEventArgs e)
            {
                using (Pen p = new Pen(Border)) e.Graphics.DrawLine(p, 0, header.Height - 1, header.Width, header.Height - 1);
            };
            header.Controls.Add(new Label
            {
                Text = Title, AutoSize = true, ForeColor = Text1,
                Font = new Font("Segoe UI Semilight", 18f), Location = new Point(28, 14)
            });

            SetupTab(installedTab, "Instalado", 32);
            SetupTab(availableTab, "Disponível", 132);
            header.Controls.Add(installedTab);
            header.Controls.Add(availableTab);
            tabLine.BackColor = Accent;
            header.Controls.Add(tabLine);
            installedTab.Click += delegate { showInstalled = true; RefreshCards(); };
            availableTab.Click += delegate { showInstalled = false; RefreshCards(); };

            Label footer = new Label
            {
                Dock = DockStyle.Bottom, Height = 30, ForeColor = Text2, TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 28, 0), Text = "Instalação por usuário, sem administrador · " + root
            };

            list.Dock = DockStyle.Fill;
            list.FlowDirection = FlowDirection.TopDown;
            list.WrapContents = false;
            list.AutoScroll = true;
            list.Padding = new Padding(28, 20, 28, 20);
            list.Resize += delegate { FitCards(); };

            empty.AutoSize = true;
            empty.ForeColor = Text2;
            empty.Margin = new Padding(4, 8, 0, 0);

            Controls.Add(list);
            Controls.Add(footer);
            Controls.Add(header);

            RefreshCards();
        }

        void SetupTab(Label tab, string text, int x)
        {
            tab.Text = text;
            tab.AutoSize = true;
            tab.Cursor = Cursors.Hand;
            tab.Font = new Font("Segoe UI", 10.5f);
            tab.Location = new Point(x, 62);
        }

        internal void RefreshCards()
        {
            installedTab.ForeColor = showInstalled ? Accent : Text2;
            availableTab.ForeColor = showInstalled ? Text2 : Accent;
            Label on = showInstalled ? installedTab : availableTab;
            tabLine.SetBounds(on.Left, on.Bottom + 4, on.Width, 3);

            list.SuspendLayout();
            foreach (Control c in list.Controls) if (c != empty) c.Dispose();
            list.Controls.Clear();
            // Instalação incompleta aparece em "Instalado", com a opção de reparar
            if (framework.AnyInstalled == showInstalled)
                list.Controls.Add(new ProductCard(this, framework));
            else
            {
                empty.Text = showInstalled ? "O Sage Framework não está instalado. Veja a aba Disponível." : "O Sage Framework já está instalado.";
                list.Controls.Add(empty);
            }
            list.ResumeLayout();
            FitCards();
        }

        void FitCards()
        {
            int w = list.ClientSize.Width - list.Padding.Horizontal - 8;
            foreach (Control c in list.Controls)
                if (c is ProductCard) c.Width = Math.Max(500, w);
        }

        internal bool Busy { get { return busy; } }

        // Executa o install.ps1 de cada componente, em sequência
        internal void Run(ProductCard card, bool uninstall, string verb)
        {
            if (busy) return;
            if (!EnsureExcelClosed()) return;

            busy = true;
            card.SetBusy(verb + "...");
            List<Component> steps = new List<Component>(framework.Components);
            if (uninstall) steps.Reverse();
            StringBuilder log = new StringBuilder();

            Thread worker = new Thread(delegate()
            {
                string failed = null;
                foreach (Component c in steps)
                {
                    Component current = c;
                    BeginInvoke((MethodInvoker)delegate { card.SetBusy(verb + ": " + current.Name.ToLowerInvariant() + "..."); });
                    if (RunScript(current, uninstall, card, log) != 0) { failed = current.Name; break; }
                }

                BeginInvoke((MethodInvoker)delegate
                {
                    busy = false;
                    if (failed != null)
                    {
                        string text;
                        lock (log) text = log.ToString().Trim();
                        if (text.Length > 1500) text = "..." + text.Substring(text.Length - 1500);
                        MessageBox.Show(this, verb + " falhou em \"" + failed + "\".\n\n" + text, Title,
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    showInstalled = framework.AnyInstalled;
                    RefreshCards();
                });
            });
            worker.IsBackground = true;
            worker.Start();
        }

        int RunScript(Component component, bool uninstall, ProductCard card, StringBuilder log)
        {
            ProcessStartInfo psi = new ProcessStartInfo("powershell.exe",
                "-NoProfile -ExecutionPolicy Bypass -File \"" + component.Script + "\"" + (uninstall ? " -Uninstall" : ""));
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            Encoding oem = Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage);
            psi.StandardOutputEncoding = oem;
            psi.StandardErrorEncoding = oem;

            try
            {
                using (Process p = Process.Start(psi))
                {
                    p.ErrorDataReceived += delegate(object s, DataReceivedEventArgs e) { if (e.Data != null) lock (log) log.AppendLine(e.Data); };
                    p.BeginErrorReadLine();
                    string line;
                    while ((line = p.StandardOutput.ReadLine()) != null)
                    {
                        lock (log) log.AppendLine(line);
                        string shown = line.Trim();
                        if (shown.Length > 0) BeginInvoke((MethodInvoker)delegate { card.SetBusy(shown); });
                    }
                    p.WaitForExit();
                    return p.ExitCode;
                }
            }
            catch (Exception ex)
            {
                lock (log) log.AppendLine(ex.Message);
                return -1;
            }
        }

        bool EnsureExcelClosed()
        {
            while (Process.GetProcessesByName("EXCEL").Length > 0)
            {
                DialogResult r = MessageBox.Show(this,
                    "Feche o Excel para continuar.\n\nO add-in fica carregado dentro do Excel e o arquivo não pode ser trocado com ele aberto.",
                    Title, MessageBoxButtons.RetryCancel, MessageBoxIcon.Information);
                if (r != DialogResult.Retry) return false;
            }
            return true;
        }

        internal Framework Framework { get { return framework; } }
    }

    // ------------------------------------------------------------------
    // Cartão do produto
    // ------------------------------------------------------------------
    sealed class ProductCard : Panel
    {
        readonly InstallerForm owner;
        readonly Framework framework;
        readonly Label name = new Label(), status = new Label(), description = new Label(), progressText = new Label();
        readonly ProgressBar progress = new ProgressBar();
        readonly FlowLayoutPanel buttons = new FlowLayoutPanel();

        public ProductCard(InstallerForm owner, Framework framework)
        {
            this.owner = owner;
            this.framework = framework;
            Height = 132;
            BackColor = InstallerForm.Surface;
            Margin = new Padding(0, 0, 0, 14);
            DoubleBuffered = true;

            name.Text = Framework.Name;
            name.Font = new Font("Segoe UI Semibold", 12f);
            name.ForeColor = InstallerForm.Text1;
            name.AutoSize = true;
            name.Location = new Point(96, 16);

            bool warn = framework.Incomplete || framework.UpdateAvailable;
            status.Text = framework.Status + (framework.UpdateAvailable ? "  ·  Atualização disponível" : "");
            status.ForeColor = warn ? InstallerForm.Accent : InstallerForm.Text2;
            status.AutoSize = true;
            status.Location = new Point(97, 44);

            description.Text = Framework.Description;
            description.ForeColor = InstallerForm.Text2;
            description.Location = new Point(97, 68);
            description.Height = 48;

            progress.Style = ProgressBarStyle.Marquee;
            progress.MarqueeAnimationSpeed = 30;
            progress.Visible = false;
            progressText.ForeColor = InstallerForm.Text2;
            progressText.Visible = false;

            buttons.FlowDirection = FlowDirection.RightToLeft;
            buttons.WrapContents = false;
            buttons.BackColor = InstallerForm.Surface;

            Controls.Add(name);
            Controls.Add(status);
            Controls.Add(description);
            Controls.Add(progress);
            Controls.Add(progressText);
            Controls.Add(buttons);
            BuildButtons();
            Resize += delegate { Relayout(); };
        }

        void BuildButtons()
        {
            if (!framework.AnyInstalled)
            {
                buttons.Controls.Add(MakeButton("Instalar", true, delegate { owner.Run(this, false, "Instalando"); }));
                return;
            }

            Button more = MakeButton("Mais  ▾", false, null);
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("Reparar", null, delegate { owner.Run(this, false, "Reparando"); });
            menu.Items.Add("Desinstalar", null, delegate
            {
                if (MessageBox.Show(owner, "Desinstalar o Sage Framework?", InstallerForm.Title,
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                    owner.Run(this, true, "Desinstalando");
            });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Abrir pasta", null, delegate { Process.Start("explorer.exe", "\"" + framework.Root + "\""); });
            more.Click += delegate { menu.Show(more, new Point(0, more.Height)); };
            buttons.Controls.Add(more);

            if (framework.Incomplete)
            {
                buttons.Controls.Add(MakeButton("Reparar", true, delegate { owner.Run(this, false, "Reparando"); }));
                return;
            }

            if (!SageShortcuts.Running)
                buttons.Controls.Add(MakeButton("Iniciar atalhos", false, delegate
                {
                    try { framework.Shortcuts.Start(); } catch (Exception ex) { MessageBox.Show(owner, ex.Message, InstallerForm.Title); }
                    Thread.Sleep(2500);
                    owner.RefreshCards();
                }));

            buttons.Controls.Add(MakeButton("Abrir Excel", !framework.UpdateAvailable, delegate
            {
                try { Process.Start("excel.exe"); } catch (Exception ex) { MessageBox.Show(owner, ex.Message, InstallerForm.Title); }
            }));

            if (framework.UpdateAvailable)
                buttons.Controls.Add(MakeButton("Atualizar", true, delegate { owner.Run(this, false, "Atualizando"); }));
        }

        Button MakeButton(string text, bool primary, EventHandler click)
        {
            Button b = new Button();
            b.Text = text;
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderColor = primary ? InstallerForm.Accent : Color.FromArgb(0xC8, 0xCC, 0xC4);
            b.FlatAppearance.MouseOverBackColor = primary ? InstallerForm.AccentHover : Color.FromArgb(0xEC, 0xEE, 0xEA);
            b.BackColor = primary ? InstallerForm.Accent : InstallerForm.Surface;
            b.ForeColor = primary ? Color.White : InstallerForm.Text1;
            b.Size = new Size(Math.Max(96, TextRenderer.MeasureText(text, InstallerForm.UiFont).Width + 32), 32);
            b.Margin = new Padding(8, 0, 0, 0);
            b.Cursor = Cursors.Hand;
            if (click != null) b.Click += delegate(object s, EventArgs e) { if (!owner.Busy) click(s, e); };
            return b;
        }

        public void SetBusy(string text)
        {
            buttons.Enabled = false;
            progress.Visible = true;
            progressText.Visible = true;
            progressText.Text = text;
            description.Visible = false;
            Relayout();
        }

        void Relayout()
        {
            int right = Width - 24;
            int bw = 0;
            foreach (Control c in buttons.Controls) bw += c.Width + c.Margin.Horizontal;
            buttons.SetBounds(right - bw, 18, bw, 34);
            description.Width = Math.Max(200, right - 97);
            progress.SetBounds(97, 76, Math.Max(200, right - 97), 6);
            progressText.SetBounds(97, 90, Math.Max(200, right - 97), 22);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using (Pen p = new Pen(InstallerForm.Border)) e.Graphics.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
            if (InstallerForm.AppIcon != null)
                using (Icon big = new Icon(InstallerForm.AppIcon, 64, 64))
                    e.Graphics.DrawIcon(big, new Rectangle(20, 18, 56, 56));
        }
    }
}
