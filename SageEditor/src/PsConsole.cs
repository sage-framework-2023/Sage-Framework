using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace SageEditor
{
    // Aba Terminal: o PowerShell 5.1 rodando dentro do Excel, numa thread própria (o Excel
    // continua respondendo enquanto um comando roda).
    //   VBA <procedimento> [argumentos]   executa um procedimento de um módulo (Application.Run)
    //                                     e mostra o retorno; Tab completa os nomes
    //   $Excel                            o Application do Excel
    //   cls / clear                       limpa a tela
    // Enter executa, Seta para cima/baixo navega no histórico, Tab completa, Esc apaga a linha
    // e Ctrl+C interrompe o comando (ou copia, com texto selecionado).
    sealed class PsConsole : UserControl
    {
        readonly RichTextBox output = new RichTextBox();
        readonly Label prompt = new Label();
        readonly InputBox input = new InputBox();
        readonly Panel inputRow = new Panel();
        readonly List<string> history = new List<string>();
        int historyIndex;
        Runspace runspace;
        PowerShell running;
        bool starting;
        string location = "";
        Theme theme;
        Color errorColor, warningColor;

        const int MaxLength = 2000000; // caracteres guardados na tela

        public PsConsole()
        {
            Font font = new Font("Consolas", 10f);
            output.ReadOnly = true;
            output.BorderStyle = BorderStyle.None;
            output.Font = font;
            output.DetectUrls = false;
            output.HideSelection = false;
            output.WordWrap = true; // as linhas já vêm na largura da tela (Out-String -Width)
            output.ScrollBars = RichTextBoxScrollBars.Vertical;
            output.Visible = false;

            prompt.AutoSize = true;
            prompt.Font = font;
            prompt.Dock = DockStyle.Left;
            prompt.Padding = new Padding(0, 1, 0, 0);
            prompt.Margin = Padding.Empty;
            prompt.Visible = false; // aparece quando o PowerShell estiver pronto
            input.BorderStyle = BorderStyle.None;
            input.Font = font;
            input.Dock = DockStyle.Fill;
            input.KeyDown += InputKeyDown;
            inputRow.Height = font.Height + 6;
            inputRow.Padding = new Padding(0, 1, 4, 0);
            inputRow.Controls.Add(input);
            inputRow.Controls.Add(prompt);

            Controls.Add(output);
            Controls.Add(inputRow);
            output.Click += delegate { if (output.SelectionLength == 0) input.Focus(); };
            MouseDown += delegate { input.Focus(); };
            inputRow.MouseDown += delegate { input.Focus(); };
            UpdatePrompt();
        }

        // Como num terminal comum: a saída cresce de cima para baixo e a linha de comando fica
        // logo depois dela; quando a tela enche, a saída rola e a linha de comando fica embaixo
        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            Arrange();
        }

        void Arrange()
        {
            int width = ClientSize.Width, rowHeight = inputRow.Height;
            int room = Math.Max(0, ClientSize.Height - rowHeight);
            int height = Math.Min(room, ContentHeight());
            output.Visible = height > 0;
            output.SetBounds(0, 0, width, Math.Max(1, height));
            inputRow.SetBounds(0, height, width, rowHeight);
            if (height >= room && output.TextLength > 0)
            {
                output.SelectionStart = output.TextLength;
                output.ScrollToCaret();
            }
        }

        // Altura do texto da saída (a última linha é a vazia depois da quebra final)
        int ContentHeight()
        {
            if (output.TextLength == 0) return 0;
            int lines = output.GetLineFromCharIndex(output.TextLength) + 1;
            int lineHeight = output.Font.Height;
            if (lines >= 2)
            {
                int first = output.GetPositionFromCharIndex(0).Y;
                int last = output.GetPositionFromCharIndex(output.GetFirstCharIndexFromLine(lines - 1)).Y;
                if (last > first) lineHeight = (last - first) / (lines - 1);
            }
            return lines * lineHeight + 2;
        }

        public void FocusInput()
        {
            input.Focus();
        }

        public void ApplyTheme(Theme t)
        {
            theme = t;
            bool dark = t.Background.GetBrightness() < 0.5f;
            errorColor = dark ? Color.FromArgb(0xF4, 0x87, 0x71) : Color.FromArgb(0xC5, 0x22, 0x22);
            warningColor = dark ? Color.FromArgb(0xCC, 0xA7, 0x00) : Color.FromArgb(0x9A, 0x6A, 0x00);
            BackColor = t.Background;
            output.BackColor = t.Background;
            output.ForeColor = t.Foreground;
            inputRow.BackColor = t.Background;
            input.BackColor = t.Background;
            input.ForeColor = t.Foreground;
            input.Dark = dark;
            input.Recolor();
            prompt.BackColor = t.Background;
            prompt.ForeColor = t.Accent;
            // Barra de rolagem escura no tema escuro (Windows 10 e 11)
            if (output.IsHandleCreated) SetWindowTheme(output.Handle, dark ? "DarkMode_Explorer" : "Explorer", null);
            else output.HandleCreated += delegate { SetWindowTheme(output.Handle, dark ? "DarkMode_Explorer" : "Explorer", null); };
        }

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        static extern int SetWindowTheme(IntPtr hwnd, string appName, string idList);

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (Visible) EnsureRunspace();
        }

        // ------------------------------------------------------------------
        // PowerShell
        // ------------------------------------------------------------------

        void EnsureRunspace()
        {
            if (runspace != null || starting) return;
            starting = true;
            input.Enabled = false;
            object excel = null;
            try { excel = ExcelBridge.Application(); }
            catch (Exception ex) { Log.Error(ex); }
            string[] folders = StartFolders(excel);
            ExcelBridge bridge = new ExcelBridge(this);
            ThreadPool.QueueUserWorkItem(delegate
            {
                Runspace rs = null;
                string error = null;
                try
                {
                    rs = RunspaceFactory.CreateRunspace();
                    rs.ThreadOptions = PSThreadOptions.ReuseThread;
                    rs.Open();
                    rs.SessionStateProxy.SetVariable("Excel", excel);
                    rs.SessionStateProxy.SetVariable("SageBridge", bridge);
                    rs.SessionStateProxy.SetVariable("SageStartFolders", folders);
                    using (PowerShell ps = PowerShell.Create())
                    {
                        ps.Runspace = rs;
                        ps.AddScript(InitScript).Invoke();
                    }
                }
                catch (Exception ex) { error = ex.Message; Log.Error(ex); }
                UI(delegate
                {
                    starting = false;
                    input.Enabled = true;
                    if (error != null) { Write(error + Environment.NewLine, errorColor); return; }
                    runspace = rs;
                    location = CurrentLocation();
                    UpdatePrompt();
                    prompt.Visible = true;
                    if (input.ContainsFocus || ContainsFocus || Visible) input.Focus();
                });
            });
        }

        const string InitScript = @"
function VBA {
    param(
        [Parameter(Mandatory = $true, Position = 0)][string]$Name,
        [Parameter(ValueFromRemainingArguments = $true)]$Arguments
    )
    $values = if ($null -eq $Arguments) { @() } else { @($Arguments) }
    $r = $SageBridge.Run($Name, [object[]]$values)
    if ($null -ne $r) { Write-Output -NoEnumerate $r }
}
Register-ArgumentCompleter -CommandName VBA -ParameterName Name -ScriptBlock {
    param($command, $parameter, $word)
    $SageBridge.Procedures() | Where-Object { $_ -like ""$word*"" } | ForEach-Object {
        [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', $_)
    }
}
$start = @($SageStartFolders) + $HOME | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Container) } | Select-Object -First 1
Set-Location -LiteralPath $start
Remove-Variable start, SageStartFolders
";

        // Onde o terminal começa: a pasta da pasta de trabalho ativa do Excel e, se ela não foi
        // salva (ou está na nuvem, com endereço https), a pasta padrão do Excel (Opções > Salvar).
        // O PowerShell usa a primeira que existir, senão a pasta do usuário.
        static string[] StartFolders(object excel)
        {
            List<string> folders = new List<string>();
            if (excel == null) return folders.ToArray();
            try
            {
                dynamic app = excel;
                dynamic book = app.ActiveWorkbook;
                if (book != null) folders.Add((string)book.Path);
            }
            catch (Exception ex) { Log.Error(ex); }
            try { folders.Add((string)((dynamic)excel).DefaultFilePath); }
            catch (Exception ex) { Log.Error(ex); }
            return folders.Where(f => !string.IsNullOrEmpty(f) && !f.Contains("://")).ToArray();
        }

        void Execute(string line)
        {
            WriteCommand(line, "");
            if (line.Trim().Length == 0) return;
            if (history.Count == 0 || history[history.Count - 1] != line) history.Add(line);
            historyIndex = history.Count;
            string trimmed = line.Trim().ToLowerInvariant();
            if (trimmed == "cls" || trimmed == "clear" || trimmed == "clear-host") { ClearScreen(); return; }
            if (runspace == null) return;

            int columns = Math.Max(40, output.ClientSize.Width / Math.Max(1, TextRenderer.MeasureText("M", output.Font).Width) - 2);
            PowerShell ps = PowerShell.Create();
            ps.Runspace = runspace;
            // ". { }" roda no escopo do terminal (as variáveis ficam); os tipos Sage aparecem
            // pelo ToString deles, e o resto pela formatação do PowerShell
            ps.AddScript(". {\n" + line + "\n} | ForEach-Object { if ($null -ne $_ -and $_.GetType().Assembly.GetName().Name -eq 'SageTypes') { $_.ToString() } else { $_ } } | Out-String -Stream -Width " + columns);
            PSDataCollection<PSObject> results = new PSDataCollection<PSObject>();
            results.DataAdded += delegate(object s, DataAddedEventArgs e) { WriteLine(Convert.ToString(results[e.Index]), null); };
            ps.Streams.Error.DataAdded += delegate(object s, DataAddedEventArgs e) { WriteLine(ErrorText(ps.Streams.Error[e.Index]), errorColor); };
            ps.Streams.Warning.DataAdded += delegate(object s, DataAddedEventArgs e) { WriteLine(Strings.ConsoleWarning + ps.Streams.Warning[e.Index].Message, warningColor); };
            ps.Streams.Information.DataAdded += delegate(object s, DataAddedEventArgs e) { WriteLine(Convert.ToString(ps.Streams.Information[e.Index].MessageData), null); };

            running = ps;
            SetBusy(true);
            try
            {
                ps.BeginInvoke<PSObject, PSObject>(null, results, null, delegate(IAsyncResult ar)
                {
                    try { ps.EndInvoke(ar); }
                    catch (PipelineStoppedException) { WriteLine("^C", warningColor); }
                    catch (Exception ex) { WriteLine(ex.Message, errorColor); }
                    string where = CurrentLocation();
                    UI(delegate
                    {
                        running = null;
                        ps.Dispose();
                        if (where.Length > 0) location = where;
                        SetBusy(false);
                    });
                }, null);
            }
            catch (Exception ex)
            {
                running = null;
                ps.Dispose();
                WriteLine(ex.Message, errorColor);
                SetBusy(false);
            }
        }

        static string ErrorText(ErrorRecord e)
        {
            Exception ex = e.Exception;
            // Erros do VBA e do Excel vêm embrulhados pelo PowerShell (MethodInvocationException)
            while (ex is MethodInvocationException || ex is TargetInvocationException)
            {
                if (ex.InnerException == null) break;
                ex = ex.InnerException;
            }
            if (ex != null && !(ex is RuntimeException) && !(ex is ParseException)) return ex.Message;
            string text = e.ToString();
            if (e.InvocationInfo != null && e.InvocationInfo.PositionMessage != null && !(ex is ParseException))
                text += Environment.NewLine + e.InvocationInfo.PositionMessage;
            return text;
        }

        string CurrentLocation()
        {
            try
            {
                Runspace rs = runspace;
                if (rs == null || rs.RunspaceAvailability != RunspaceAvailability.Available) return "";
                return rs.SessionStateProxy.Path.CurrentLocation.Path;
            }
            catch (Exception) { return ""; }
        }

        string PromptText { get { return "PS " + location + "> "; } }

        void UpdatePrompt()
        {
            prompt.Text = PromptText;
        }

        void SetBusy(bool busy)
        {
            prompt.Visible = !busy;
            input.ReadOnly = busy;
            if (!busy) { UpdatePrompt(); input.Focus(); }
        }

        void Stop()
        {
            PowerShell ps = running;
            if (ps == null) return;
            try { ps.BeginStop(null, null); }
            catch (Exception ex) { Log.Error(ex); }
        }

        // Tab: completa pelo próprio PowerShell (comandos, caminhos, parâmetros e os nomes do VBA)
        void Complete()
        {
            if (runspace == null || running != null || runspace.RunspaceAvailability != RunspaceAvailability.Available) return;
            string text = input.Text;
            int caret = input.SelectionStart;
            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = runspace;
                    CommandCompletion c = CommandCompletion.CompleteInput(text, caret, null, ps);
                    if (c == null || c.CompletionMatches.Count == 0) return;
                    if (completionText != text || completionCaret != caret) completionIndex = -1;
                    completionIndex = (completionIndex + 1) % c.CompletionMatches.Count;
                    string replaced = text.Substring(0, c.ReplacementIndex) + c.CompletionMatches[completionIndex].CompletionText +
                        text.Substring(c.ReplacementIndex + c.ReplacementLength);
                    // Tab de novo: a próxima opção, a partir do texto original
                    input.Text = replaced;
                    input.SelectionStart = c.ReplacementIndex + c.CompletionMatches[completionIndex].CompletionText.Length;
                    completionBase = text;
                    completionBaseCaret = caret;
                    completionText = input.Text;
                    completionCaret = input.SelectionStart;
                }
            }
            catch (Exception ex) { Log.Error(ex); }
        }

        string completionText, completionBase;
        int completionCaret, completionBaseCaret, completionIndex = -1;

        void InputKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    e.SuppressKeyPress = true;
                    if (running != null) return;
                    string line = input.Text;
                    input.Clear();
                    completionIndex = -1;
                    Execute(line);
                    break;
                case Keys.Tab:
                    e.SuppressKeyPress = true;
                    if (completionIndex >= 0 && input.Text == completionText && input.SelectionStart == completionCaret)
                    {
                        // ciclo: volta ao texto original e pede a próxima opção
                        input.Text = completionBase;
                        input.SelectionStart = completionBaseCaret;
                        completionText = completionBase;
                        completionCaret = completionBaseCaret;
                    }
                    else completionIndex = -1;
                    Complete();
                    break;
                case Keys.Up:
                case Keys.Down:
                    e.SuppressKeyPress = true;
                    if (history.Count == 0) return;
                    historyIndex = Math.Max(0, Math.Min(history.Count, historyIndex + (e.KeyCode == Keys.Up ? -1 : 1)));
                    input.Text = historyIndex < history.Count ? history[historyIndex] : "";
                    input.SelectionStart = input.Text.Length;
                    break;
                case Keys.Escape:
                    e.SuppressKeyPress = true;
                    input.Clear();
                    break;
                case Keys.C:
                    if (e.Control && !e.Shift && !e.Alt && input.SelectionLength == 0)
                    {
                        e.SuppressKeyPress = true;
                        if (running != null) Stop();
                        else { WriteCommand(input.Text, "^C"); input.Clear(); }
                    }
                    break;
                case Keys.L:
                    if (e.Control && !e.Shift && !e.Alt) { e.SuppressKeyPress = true; ClearScreen(); }
                    break;
                case Keys.V:
                    if (e.Control && !e.Alt) { e.SuppressKeyPress = true; input.PastePlain(); }
                    break;
                case Keys.Insert:
                    if (e.Shift && !e.Control && !e.Alt) { e.SuppressKeyPress = true; input.PastePlain(); }
                    break;
            }
        }

        // ------------------------------------------------------------------
        // Tela
        // ------------------------------------------------------------------

        void WriteLine(string text, Color? color)
        {
            Write((text ?? "") + Environment.NewLine, color);
        }

        // De qualquer thread
        void Write(string text, Color? color)
        {
            UI(delegate
            {
                if (output.TextLength > MaxLength)
                {
                    output.Select(0, output.TextLength - MaxLength / 2);
                    output.SelectedText = "";
                }
                // A quebra de linha do fim fica para a próxima escrita: sem uma linha vazia no
                // fim, a caixa tem a altura exata do texto (sem barra de rolagem nem espaço)
                if (pendingNewLine && text.Length > 0) { Append("\n", null); pendingNewLine = false; }
                if (text.EndsWith("\n")) { text = text.Substring(0, text.Length - 1); pendingNewLine = true; }
                if (text.EndsWith("\r")) text = text.Substring(0, text.Length - 1);
                Append(text, color);
                Arrange();
                output.ScrollToCaret();
            });
        }

        bool pendingNewLine;

        // O comando executado fica na tela com as mesmas cores da linha de comando
        void WriteCommand(string line, string suffix)
        {
            Write(PromptText, prompt.ForeColor);
            Color?[] map = PsColors.Map(line, input.Dark);
            int i = 0;
            while (i < line.Length)
            {
                int j = i;
                while (j < line.Length && map[j] == map[i]) j++;
                Write(line.Substring(i, j - i), map[i]);
                i = j;
            }
            Write(suffix + Environment.NewLine, theme != null ? theme.Muted : ForeColor);
        }

        void Append(string text, Color? color)
        {
            if (text.Length == 0) return;
            output.SelectionStart = output.TextLength;
            output.SelectionLength = 0;
            output.SelectionColor = color ?? output.ForeColor;
            output.AppendText(text);
            output.SelectionStart = output.TextLength;
        }

        void ClearScreen()
        {
            output.Clear();
            pendingNewLine = false;
            Arrange();
        }

        void UI(MethodInvoker action)
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                try { BeginInvoke(action); }
                catch (InvalidOperationException) { } // fechando
                return;
            }
            try { action(); }
            catch (Exception ex) { Log.Error(ex); }
        }

        // Chamadas do PowerShell que precisam da thread do Excel (o VBA e o modelo de objetos)
        internal T OnUiThread<T>(Func<T> f)
        {
            if (!InvokeRequired) return f();
            T result = default(T);
            Exception error = null;
            Invoke((MethodInvoker)delegate
            {
                try { result = f(); }
                catch (Exception ex) { error = ex; }
            });
            if (error != null) throw error;
            return result;
        }

        public void Shutdown()
        {
            Stop();
            Runspace rs = runspace;
            runspace = null;
            if (rs != null)
            {
                try
                {
                    if (rs.RunspaceAvailability == RunspaceAvailability.Available)
                    {
                        rs.SessionStateProxy.SetVariable("Excel", null);
                        rs.SessionStateProxy.SetVariable("SageBridge", null);
                    }
                    rs.Dispose();
                }
                catch (Exception ex) { Log.Error(ex); }
            }
            ExcelBridge.Release();
        }

        // Linha de comando: uma linha só, colorida enquanto se digita (PsColors), que fica com
        // Enter, Tab e as setas (o VBE e o WinForms as usariam para navegar)
        sealed class InputBox : RichTextBox
        {
            public bool Dark = true;
            bool coloring;

            public InputBox()
            {
                Multiline = false;
                DetectUrls = false;
                ScrollBars = RichTextBoxScrollBars.None;
            }

            protected override void OnTextChanged(EventArgs e)
            {
                base.OnTextChanged(e);
                Recolor();
            }

            public void Recolor()
            {
                if (coloring || !IsHandleCreated) return;
                coloring = true;
                int start = SelectionStart, length = SelectionLength;
                SendMessage(Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
                try
                {
                    string text = Text;
                    Color?[] map = PsColors.Map(text, Dark);
                    int i = 0;
                    while (i < text.Length)
                    {
                        int j = i;
                        while (j < text.Length && map[j] == map[i]) j++;
                        Select(i, j - i);
                        SelectionColor = map[i] ?? ForeColor;
                        i = j;
                    }
                    Select(start, length);
                }
                finally
                {
                    SendMessage(Handle, WM_SETREDRAW, (IntPtr)1, IntPtr.Zero);
                    Invalidate();
                    coloring = false;
                }
            }

            // Colar: só o texto, numa linha (as quebras viram "; ")
            public void PastePlain()
            {
                if (ReadOnly || !Clipboard.ContainsText()) return;
                string text = Clipboard.GetText().Replace("\r\n", "\n").Trim('\n').Replace("\n", "; ");
                SelectedText = text;
            }

            const int WM_SETREDRAW = 0x000B;
            [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

            protected override bool IsInputKey(Keys keyData)
            {
                Keys key = keyData & Keys.KeyCode;
                if (key == Keys.Enter || key == Keys.Tab || key == Keys.Up || key == Keys.Down || key == Keys.Escape) return true;
                return base.IsInputKey(keyData);
            }

            protected override bool ProcessDialogKey(Keys keyData)
            {
                Keys key = keyData & Keys.KeyCode;
                if (key == Keys.Tab || key == Keys.Enter || key == Keys.Escape) return false;
                return base.ProcessDialogKey(keyData);
            }

            // O VBE trata as teclas no laço de mensagens dele (aceleradores); WM_GETDLGCODE
            // diz que esta caixa quer todas
            protected override void WndProc(ref Message m)
            {
                const int WM_GETDLGCODE = 0x0087, DLGC_WANTALLKEYS = 0x0004, DLGC_WANTCHARS = 0x0080, DLGC_WANTARROWS = 0x0001, DLGC_WANTTAB = 0x0002;
                if (m.Msg == WM_GETDLGCODE)
                {
                    m.Result = (IntPtr)(DLGC_WANTALLKEYS | DLGC_WANTCHARS | DLGC_WANTARROWS | DLGC_WANTTAB);
                    return;
                }
                base.WndProc(ref m);
            }
        }
    }

    // O que o terminal enxerga do Excel: o comando VBA e os nomes dos procedimentos.
    // Chamado na thread do PowerShell; o trabalho é feito na thread do Excel.
    [ComVisible(false)]
    public sealed class ExcelBridge
    {
        readonly PsConsole console;
        internal ExcelBridge(PsConsole console) { this.console = console; }

        // Application.Run "<procedimento>", argumentos...
        public object Run(string name, object[] args)
        {
            object[] values = (args ?? new object[0]).Select(Unwrap).ToArray();
            return console.OnUiThread(delegate
            {
                object app = Application();
                object[] all = new object[values.Length + 1];
                all[0] = name;
                Array.Copy(values, 0, all, 1, values.Length);
                try { return app.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, app, all); }
                catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
            });
        }

        // "Modulo.Procedimento" de todos os módulos padrão dos projetos abertos (sem senha)
        public string[] Procedures()
        {
            return console.OnUiThread(delegate
            {
                SortedSet<string> names = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
                dynamic vbe = TerminalWindow.Vbe;
                foreach (dynamic project in vbe.VBProjects)
                {
                    try
                    {
                        if ((int)project.Protection != 0) continue;
                        foreach (dynamic component in project.VBComponents)
                        {
                            if ((int)component.Type != 1) continue; // vbext_ct_StdModule
                            string module = component.Name;
                            dynamic code = component.CodeModule;
                            int total = code.CountOfLines;
                            int line = (int)code.CountOfDeclarationLines + 1;
                            while (line <= total)
                            {
                                int kind = 0;
                                string proc = code.ProcOfLine(line, ref kind);
                                if (string.IsNullOrEmpty(proc)) { line++; continue; }
                                names.Add(module + "." + proc);
                                line = (int)code.ProcStartLine(proc, kind) + (int)code.ProcCountLines(proc, kind);
                            }
                        }
                    }
                    catch (Exception ex) { Log.Error(ex); }
                }
                return names.ToArray();
            });
        }

        static object Unwrap(object value)
        {
            PSObject ps = value as PSObject;
            return ps != null ? ps.BaseObject : value;
        }

        // ------------------------------------------------------------------
        // O Application do Excel deste processo, pela janela de uma pasta de trabalho
        // (EXCEL7): o objeto Window que ela expõe leva ao Application
        // ------------------------------------------------------------------

        static object application;

        // Ao desligar: uma referência presa ao Application impede o Excel de fechar
        internal static void Release()
        {
            object app = application;
            application = null;
            if (app != null && Marshal.IsComObject(app))
            {
                try { Marshal.ReleaseComObject(app); }
                catch (Exception) { }
            }
        }

        [DllImport("oleacc.dll")]
        static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object obj);
        const uint OBJID_NATIVEOM = 0xFFFFFFF0;

        internal static object Application()
        {
            if (application != null) return application;
            IntPtr sheet = IntPtr.Zero;
            Native.EnumThreadWindows(Native.GetCurrentThreadId(), delegate(IntPtr top, IntPtr l)
            {
                if (Native.ClassName(top) != "XLMAIN") return true;
                Native.EnumChildWindows(top, delegate(IntPtr h, IntPtr l2)
                {
                    if (Native.ClassName(h) == "EXCEL7") { sheet = h; return false; }
                    return true;
                }, IntPtr.Zero);
                return sheet == IntPtr.Zero;
            }, IntPtr.Zero);
            if (sheet != IntPtr.Zero)
            {
                Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
                object window;
                if (AccessibleObjectFromWindow(sheet, OBJID_NATIVEOM, ref dispatch, out window) >= 0 && window != null)
                    return application = ((dynamic)window).Application;
            }
            throw new InvalidOperationException(Strings.ConsoleNoWorkbook);
        }
    }
}
