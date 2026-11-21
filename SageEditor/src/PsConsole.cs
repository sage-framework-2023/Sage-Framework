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
    //   exit                              fecha a janela e recomeça a sessão, com a tela limpa
    // Enter executa, Seta para cima/baixo navega no histórico, Tab completa, Esc apaga a linha
    // e Ctrl+C interrompe o comando (ou copia, com texto selecionado).
    //
    // A tela é uma caixa só, como num terminal comum: a saída e, na última linha, o prompt e a
    // linha de comando. Tudo pode ser selecionado e copiado (inclusive da saída até a linha de
    // comando), mas só o que vem depois do prompt pode ser editado: digitar com o cursor na saída
    // leva o cursor para a linha de comando, e Backspace, Recortar e Colar não alcançam o resto.
    // promptStart: onde começa o prompt atual (-1: sem prompt, com um comando rodando ou o
    // PowerShell iniciando); inputStart: onde começa a linha de comando, logo depois do prompt.
    sealed class PsConsole : UserControl
    {
        readonly TerminalBox box = new TerminalBox();
        readonly List<string> history = new List<string>();
        int historyIndex;
        Runspace runspace;
        PowerShell running;
        bool starting;
        int promptStart = -1, inputStart;
        bool endsWithNewLine = true; // a tela termina numa quebra de linha (ou está vazia)
        bool writing, coloring, dark = true;
        string location = "";
        Theme theme;
        Color errorColor, warningColor, promptColor;

        const int MaxLength = 2000000; // caracteres guardados na tela
        const int WM_SETREDRAW = 0x000B, EM_GETTEXTRANGE = 0x044B, EM_SETUNDOLIMIT = 0x0452;

        public PsConsole()
        {
            box.BorderStyle = BorderStyle.None;
            box.Font = new Font("Consolas", 10f);
            box.DetectUrls = false;
            box.HideSelection = false;
            box.WordWrap = true; // as linhas já vêm na largura da tela (Out-String -Width)
            box.ScrollBars = RichTextBoxScrollBars.Vertical;
            box.Dock = DockStyle.Fill;
            box.KeyDown += BoxKeyDown;
            box.KeyPress += BoxKeyPress;
            box.TextChanged += delegate { if (!writing && !coloring) RecolorInput(); };
            box.MouseUp += delegate(object s, MouseEventArgs e) { if (e.Button == MouseButtons.Right) ShowMenu(e.Location); };
            // Sem Desfazer: ele traria de volta a saída e o prompt, que não são editáveis
            box.HandleCreated += delegate { SendMessage(box.Handle, EM_SETUNDOLIMIT, IntPtr.Zero, IntPtr.Zero); };
            Controls.Add(box);
            MouseDown += delegate { FocusInput(); };
        }

        public void FocusInput()
        {
            box.Focus();
            if (promptStart >= 0 && box.SelectionLength == 0 && box.SelectionStart < inputStart)
            {
                box.Select(box.TextLength, 0);
                box.ScrollToCaret();
            }
        }

        public void ApplyTheme(Theme t)
        {
            theme = t;
            dark = t.Background.GetBrightness() < 0.5f;
            errorColor = dark ? Color.FromArgb(0xF4, 0x87, 0x71) : Color.FromArgb(0xC5, 0x22, 0x22);
            warningColor = dark ? Color.FromArgb(0xCC, 0xA7, 0x00) : Color.FromArgb(0x9A, 0x6A, 0x00);
            promptColor = t.Foreground; // o prompt na cor do texto, como num terminal comum
            BackColor = t.Background;
            box.BackColor = t.Background;
            box.ForeColor = t.Foreground;
            RecolorInput();
            // Barra de rolagem escura no tema escuro (Windows 10 e 11)
            bool d = dark;
            if (box.IsHandleCreated) SetWindowTheme(box.Handle, d ? "DarkMode_Explorer" : "Explorer", null);
            else box.HandleCreated += delegate { SetWindowTheme(box.Handle, d ? "DarkMode_Explorer" : "Explorer", null); };
        }

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        static extern int SetWindowTheme(IntPtr hwnd, string appName, string idList);
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wParam, ref TEXTRANGE range);

        [StructLayout(LayoutKind.Sequential)]
        struct TEXTRANGE { public int cpMin, cpMax; public IntPtr text; }

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
                    rs.SessionStateProxy.SetVariable("SageNative", new NativeLauncher());
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
                    if (error != null) { WriteLine(error, errorColor); return; }
                    runspace = rs;
                    location = CurrentLocation();
                    NewPrompt();
                    if (ContainsFocus || Visible) FocusInput();
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
# Programas de console (ping, git, python): iniciados pelo terminal, sem janela e com a
# saída decodificada linha a linha (NativeProcess.cs)
$ExecutionContext.InvokeCommand.PostCommandLookupAction = {
    param($name, $e)
    $cmd = $e.Command
    if ($cmd.CommandType -ne 'Application' -or -not $SageNative.IsConsole($cmd.Path)) { return }
    $path = $cmd.Path
    $e.CommandScriptBlock = {
        $p = $SageNative.Start($path, [object[]]$args, (Get-Location -PSProvider FileSystem).ProviderPath)
        try {
            if ($MyInvocation.ExpectingInput) { foreach ($item in $input) { $p.WriteInput([string]$item) } }
            $p.CloseInput()
            while ($true) {
                $line = $p.Next(200)
                if ($null -eq $line) { if ($p.Done) { break } else { continue } }
                if ($line.IsError) { Write-Information -MessageData $line.Text -Tags 'sage-stderr' } else { $line.Text }
            }
        }
        finally { $p.Stop() }
        $global:LASTEXITCODE = $p.ExitCode
    }.GetNewClosure()
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
            if (line.Trim().Length == 0) { NewPrompt(); return; }
            if (history.Count == 0 || history[history.Count - 1] != line) history.Add(line);
            historyIndex = history.Count;
            string trimmed = line.Trim().ToLowerInvariant();
            if (trimmed == "cls" || trimmed == "clear" || trimmed == "clear-host") { ClearScreen(); NewPrompt(); return; }
            if (trimmed == "exit") { Exit(); return; }
            if (runspace == null) { NewPrompt(); return; }

            int columns = Math.Max(40, box.ClientSize.Width / Math.Max(1, TextRenderer.MeasureText("M", box.Font).Width) - 2);
            PowerShell ps = PowerShell.Create();
            ps.Runspace = runspace;
            // ". { }" roda no escopo do terminal (as variáveis ficam); os tipos Sage aparecem
            // pelo ToString deles, e o resto pela formatação do PowerShell
            ps.AddScript(". {\n" + line + "\n} | ForEach-Object { if ($null -ne $_ -and $_.GetType().Assembly.GetName().Name -eq 'SageTypes') { $_.ToString() } else { $_ } } | Out-String -Stream -Width " + columns);
            PSDataCollection<PSObject> results = new PSDataCollection<PSObject>();
            results.DataAdded += delegate(object s, DataAddedEventArgs e) { WriteLine(Convert.ToString(results[e.Index]), null); };
            ps.Streams.Error.DataAdded += delegate(object s, DataAddedEventArgs e) { WriteLine(ErrorText(ps.Streams.Error[e.Index]), errorColor); };
            ps.Streams.Warning.DataAdded += delegate(object s, DataAddedEventArgs e) { WriteLine(Strings.ConsoleWarning + ps.Streams.Warning[e.Index].Message, warningColor); };
            ps.Streams.Information.DataAdded += delegate(object s, DataAddedEventArgs e)
            {
                InformationRecord info = ps.Streams.Information[e.Index];
                bool stderr = info.Tags != null && info.Tags.Contains("sage-stderr");
                WriteLine(Convert.ToString(info.MessageData), stderr ? errorColor : (Color?)null);
            };

            running = ps;
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
                        NewPrompt();
                        FocusInput();
                    });
                }, null);
            }
            catch (Exception ex)
            {
                running = null;
                ps.Dispose();
                WriteLine(ex.Message, errorColor);
                NewPrompt();
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

        // ------------------------------------------------------------------
        // Prompt e linha de comando
        // ------------------------------------------------------------------

        // Prompt novo, numa linha própria no fim da tela
        void NewPrompt()
        {
            if (promptStart >= 0) return;
            writing = true;
            if (!endsWithNewLine) AppendAtEnd("\n", null);
            promptStart = box.TextLength;
            AppendAtEnd(PromptText, promptColor);
            inputStart = box.TextLength;
            box.Select(box.TextLength, 0);
            box.SelectionColor = box.ForeColor;
            writing = false;
            box.ScrollToCaret();
        }

        // Fim da linha de comando (Enter, Ctrl+C): o que foi digitado fica na tela como está
        void EndPrompt(string suffix)
        {
            if (promptStart < 0) return;
            writing = true;
            if (!string.IsNullOrEmpty(suffix)) AppendAtEnd(suffix, theme != null ? theme.Muted : box.ForeColor);
            AppendAtEnd("\n", null);
            promptStart = -1;
            box.Select(box.TextLength, 0);
            writing = false;
            box.ScrollToCaret();
        }

        void AppendAtEnd(string text, Color? color)
        {
            if (text.Length == 0) return;
            box.Select(box.TextLength, 0);
            box.SelectionColor = color ?? box.ForeColor;
            box.SelectedText = text;
            endsWithNewLine = text.EndsWith("\n");
        }

        // O que está digitado na linha de comando (atribuir troca o texto todo)
        string InputText
        {
            get { return promptStart < 0 ? "" : TextRange(inputStart, box.TextLength); }
            set
            {
                if (promptStart < 0) return;
                box.Select(inputStart, box.TextLength - inputStart);
                box.SelectionColor = box.ForeColor;
                box.SelectedText = value ?? ""; // o TextChanged recolore
            }
        }

        // Posição do cursor dentro da linha de comando
        int Caret { get { return Math.Max(0, box.SelectionStart - inputStart); } }

        string TextRange(int start, int end)
        {
            if (end <= start) return "";
            if (!box.IsHandleCreated) return box.Text.Substring(start, end - start);
            IntPtr buffer = Marshal.AllocHGlobal((end - start + 1) * 2);
            try
            {
                TEXTRANGE range = new TEXTRANGE { cpMin = start, cpMax = end, text = buffer };
                SendMessage(box.Handle, EM_GETTEXTRANGE, IntPtr.Zero, ref range);
                return Marshal.PtrToStringUni(buffer);
            }
            finally { Marshal.FreeHGlobal(buffer); }
        }

        // As cores do PowerShell (PsColors) na linha de comando, enquanto se digita
        void RecolorInput()
        {
            if (promptStart < 0 || coloring || !box.IsHandleCreated) return;
            coloring = true;
            int start = box.SelectionStart, length = box.SelectionLength;
            SendMessage(box.Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
            try
            {
                string text = InputText;
                Color?[] map = PsColors.Map(text, dark);
                int i = 0;
                while (i < text.Length)
                {
                    int j = i;
                    while (j < text.Length && map[j] == map[i]) j++;
                    box.Select(inputStart + i, j - i);
                    box.SelectionColor = map[i] ?? box.ForeColor;
                    i = j;
                }
                box.Select(start, length);
            }
            finally
            {
                SendMessage(box.Handle, WM_SETREDRAW, (IntPtr)1, IntPtr.Zero);
                box.Invalidate();
                coloring = false;
            }
        }

        // Antes de uma edição: a seleção (ou o cursor) só dentro da linha de comando. Fora dela,
        // o cursor vai para o fim; uma seleção que começa antes fica só com a parte de dentro.
        // false: não há linha de comando (um comando rodando)
        bool PrepareEdit()
        {
            if (promptStart < 0) return false;
            int s = box.SelectionStart, e = s + box.SelectionLength;
            if (s >= inputStart) return true;
            if (e <= inputStart) box.Select(box.TextLength, 0);
            else box.Select(inputStart, e - inputStart);
            return true;
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
            string text = InputText;
            int caret = Caret;
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
                    InputText = replaced;
                    int newCaret = c.ReplacementIndex + c.CompletionMatches[completionIndex].CompletionText.Length;
                    box.Select(inputStart + newCaret, 0);
                    completionBase = text;
                    completionBaseCaret = caret;
                    completionText = replaced;
                    completionCaret = newCaret;
                }
            }
            catch (Exception ex) { Log.Error(ex); }
        }

        string completionText, completionBase;
        int completionCaret, completionBaseCaret, completionIndex = -1;

        void BoxKeyDown(object sender, KeyEventArgs e)
        {
            bool ctrl = e.Control, shift = e.Shift, alt = e.Alt;
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    Suppress(e);
                    if (running != null || promptStart < 0) return;
                    {
                        string line = InputText;
                        completionIndex = -1;
                        EndPrompt(null);
                        Execute(line);
                    }
                    return;
                case Keys.Tab:
                    Suppress(e);
                    if (!PrepareEdit()) return;
                    if (completionIndex >= 0 && InputText == completionText && Caret == completionCaret)
                    {
                        // ciclo: volta ao texto original e pede a próxima opção
                        InputText = completionBase;
                        box.Select(inputStart + completionBaseCaret, 0);
                        completionText = completionBase;
                        completionCaret = completionBaseCaret;
                    }
                    else completionIndex = -1;
                    Complete();
                    return;
                case Keys.Up:
                case Keys.Down:
                    if (shift || ctrl || alt) return; // seleção pelo teclado, como sempre
                    Suppress(e);
                    if (promptStart < 0 || history.Count == 0) return;
                    historyIndex = Math.Max(0, Math.Min(history.Count, historyIndex + (e.KeyCode == Keys.Up ? -1 : 1)));
                    InputText = historyIndex < history.Count ? history[historyIndex] : "";
                    box.Select(box.TextLength, 0);
                    box.ScrollToCaret();
                    return;
                case Keys.Escape:
                    Suppress(e);
                    if (promptStart < 0) return;
                    InputText = "";
                    box.Select(box.TextLength, 0);
                    return;
                case Keys.Home:
                    // Na linha de comando, o começo é logo depois do prompt
                    if (ctrl || alt || promptStart < 0 || box.SelectionStart < inputStart) return;
                    Suppress(e);
                    if (shift) box.Select(inputStart, box.SelectionStart + box.SelectionLength - inputStart);
                    else box.Select(inputStart, 0);
                    return;
                case Keys.Left:
                    if (!shift && !ctrl && !alt && promptStart >= 0 && box.SelectionLength == 0 && box.SelectionStart == inputStart) Suppress(e);
                    return;
                case Keys.Back:
                    if (alt) { Suppress(e); return; }
                    if (!PrepareEdit()) { Suppress(e); return; }
                    if (box.SelectionLength > 0) return; // apaga a seleção, já só dentro da linha de comando
                    if (box.SelectionStart <= inputStart) { Suppress(e); return; }
                    if (ctrl) { Suppress(e); DeleteWordBack(); }
                    return;
                case Keys.Delete:
                    if (shift && !ctrl && !alt) { Suppress(e); Cut(); return; }
                    if (!PrepareEdit()) Suppress(e);
                    return;
                case Keys.C:
                    if (!ctrl || alt || shift) break;
                    Suppress(e);
                    if (box.SelectionLength > 0) { CopyText(box.SelectedText); return; }
                    if (running != null) Stop();
                    else if (promptStart >= 0) { completionIndex = -1; EndPrompt("^C"); NewPrompt(); }
                    return;
                case Keys.Insert:
                    if (ctrl && !shift && !alt) { Suppress(e); if (box.SelectionLength > 0) CopyText(box.SelectedText); return; }
                    if (shift && !ctrl && !alt) { Suppress(e); PastePlain(); return; }
                    break;
                case Keys.X:
                    if (ctrl && !alt && !shift) { Suppress(e); Cut(); return; }
                    break;
                case Keys.V:
                    if (ctrl && !alt) { Suppress(e); PastePlain(); return; }
                    break;
                case Keys.A:
                    if (ctrl && !alt && !shift) { Suppress(e); box.SelectAll(); return; }
                    break;
                case Keys.L:
                    if (ctrl && !alt && !shift) { Suppress(e); ClearScreen(); return; }
                    break;
            }
            // Os outros atalhos com Ctrl da caixa (Desfazer, alinhamento, espaçamento, maiúsculas...)
            // mudariam o texto ou a formatação: só a navegação passa (Ctrl+setas, Ctrl+Home/End)
            if (ctrl && !alt && !IsNavigation(e.KeyCode)) Suppress(e);
        }

        static bool IsNavigation(Keys key)
        {
            return key == Keys.Left || key == Keys.Right || key == Keys.Up || key == Keys.Down || key == Keys.Home ||
                key == Keys.End || key == Keys.PageUp || key == Keys.PageDown || key == Keys.ControlKey || key == Keys.ShiftKey;
        }

        static void Suppress(KeyEventArgs e)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        // Caracteres digitados: só na linha de comando
        void BoxKeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\b')
            {
                // Backspace (também Ctrl+H): nunca antes do começo da linha de comando
                if (!PrepareEdit() || (box.SelectionLength == 0 && box.SelectionStart <= inputStart)) e.Handled = true;
                return;
            }
            if (e.KeyChar < ' ' || e.KeyChar == (char)127) { e.Handled = true; return; } // Ctrl+M, Ctrl+Backspace...: tratados no KeyDown
            if (!PrepareEdit()) e.Handled = true;
        }

        void DeleteWordBack()
        {
            string text = InputText;
            int caret = Math.Min(Caret, text.Length), i = caret;
            while (i > 0 && text[i - 1] == ' ') i--;
            while (i > 0 && text[i - 1] != ' ') i--;
            box.Select(inputStart + i, caret - i);
            box.SelectedText = "";
        }

        // ------------------------------------------------------------------
        // Tela
        // ------------------------------------------------------------------

        // Um erro do VBA (ErrorToTerminal), em vermelho, acima do prompt
        public void PrintError(string text)
        {
            Color color = errorColor.IsEmpty ? Color.FromArgb(0xF4, 0x87, 0x71) : errorColor;
            Write((text ?? "").TrimEnd('\n') + "\n", color);
        }

        void WriteLine(string text, Color? color)
        {
            Write((text ?? "") + "\n", color);
        }

        // De qualquer thread
        void Write(string text, Color? color)
        {
            UI(delegate { WriteNow(text, color); });
        }

        void WriteNow(string text, Color? color)
        {
            if (string.IsNullOrEmpty(text)) return;
            text = text.Replace("\r\n", "\n").Replace('\r', '\n'); // o RichEdit guarda cada quebra como um caractere
            writing = true;
            int start = box.SelectionStart, length = box.SelectionLength;
            // Com o cursor no fim (ou na linha de comando), a tela acompanha o texto novo; com uma
            // seleção ou o cursor no meio da saída, fica onde está (para copiar com calma)
            bool follow = length == 0 && start >= (promptStart >= 0 ? inputStart : box.TextLength);
            int at;
            if (promptStart >= 0)
            {
                // Com o prompt na tela (o Console.PrintF com o terminal parado): o texto entra antes dele
                at = promptStart;
                if (!text.EndsWith("\n")) text += "\n";
            }
            else at = box.TextLength;
            box.Select(at, 0);
            box.SelectionColor = color ?? box.ForeColor;
            box.SelectedText = text;
            if (promptStart >= 0) { promptStart += text.Length; inputStart += text.Length; }
            else endsWithNewLine = text.EndsWith("\n");
            if (start >= at) start += text.Length;
            start = Trim(start);
            if (follow && promptStart < 0) box.Select(box.TextLength, 0);
            else box.Select(Math.Min(start, box.TextLength), length);
            writing = false;
            if (follow) box.ScrollToCaret();
        }

        // Tela cheia demais: apaga o começo (nunca o prompt)
        int Trim(int start)
        {
            if (box.TextLength <= MaxLength) return start;
            int cut = box.TextLength - MaxLength / 2;
            if (promptStart >= 0) cut = Math.Min(cut, promptStart);
            if (cut <= 0) return start;
            box.Select(0, cut);
            box.SelectedText = "";
            if (promptStart >= 0) { promptStart -= cut; inputStart -= cut; }
            return Math.Max(0, start - cut);
        }

        // Console.PrintF do VBA. A macro ocupa a thread da janela, então a tela é pintada aqui
        // mesmo; num laço rápido, os textos se juntam e saem de tempos em tempos: 50 ms, ou mais
        // se escrever na tela demorar, para a escrita não tomar mais de 1/5 do tempo da macro.
        // O que sobrar no fim sai no timer, assim que a thread ficar livre.
        const int PrintInterval = 50, PrintWeight = 4;
        const uint RDW_ALLCHILDREN = 0x0080, RDW_UPDATENOW = 0x0100;
        readonly StringBuilder printed = new StringBuilder();
        readonly System.Diagnostics.Stopwatch sincePrint = System.Diagnostics.Stopwatch.StartNew();
        System.Windows.Forms.Timer printTimer;
        long printWait = PrintInterval;

        public void Print(string text)
        {
            UI(delegate
            {
                printed.Append(text);
                if (sincePrint.ElapsedMilliseconds >= printWait) { FlushPrinted(); return; }
                if (printTimer == null)
                {
                    printTimer = new System.Windows.Forms.Timer { Interval = PrintInterval };
                    printTimer.Tick += delegate { FlushPrinted(); };
                }
                printTimer.Start();
            });
        }

        void FlushPrinted()
        {
            if (printTimer != null) printTimer.Stop();
            sincePrint.Restart();
            if (printed.Length == 0) return;
            string text = printed.ToString();
            printed.Length = 0;
            Write(text, null);
            if (IsHandleCreated) Native.RedrawWindow(Handle, IntPtr.Zero, IntPtr.Zero, RDW_UPDATENOW | RDW_ALLCHILDREN);
            printWait = Math.Max(PrintInterval, sincePrint.ElapsedMilliseconds * PrintWeight);
            sincePrint.Restart();
        }

        // ------------------------------------------------------------------
        // Copiar e colar
        // ------------------------------------------------------------------

        void ShowMenu(Point at)
        {
            Theme t = theme ?? Theme.Find(Settings.ColorTheme);
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Renderer = new ToolStripProfessionalRenderer(new EditorTabs.MenuColors(t));
            menu.BackColor = t.Sidebar;
            menu.ForeColor = t.Foreground;
            menu.ShowImageMargin = false;
            menu.ShowCheckMargin = false;
            menu.Padding = new Padding(2);
            bool canPaste = false;
            try { canPaste = promptStart >= 0 && Clipboard.ContainsText(); }
            catch (ExternalException) { } // área de transferência presa por outro programa
            AddMenuItem(menu, t, Strings.ConsoleCopy, "Ctrl+C", box.SelectionLength > 0, delegate { CopyText(box.SelectedText); });
            AddMenuItem(menu, t, Strings.ConsoleCopyAll, null, box.TextLength > 0, delegate { CopyText(box.Text); });
            menu.Items.Add(new ToolStripSeparator());
            AddMenuItem(menu, t, Strings.ConsolePaste, "Ctrl+V", canPaste, delegate { box.Focus(); PastePlain(); });
            AddMenuItem(menu, t, Strings.ConsoleSelectAll, "Ctrl+A", box.TextLength > 0, delegate { box.Focus(); box.SelectAll(); });
            menu.Closed += delegate { BeginInvoke((MethodInvoker)menu.Dispose); };
            menu.Show(box, at);
        }

        static void AddMenuItem(ContextMenuStrip menu, Theme t, string text, string shortcut, bool enabled, EventHandler click)
        {
            ToolStripMenuItem item = new ToolStripMenuItem(text);
            item.ForeColor = enabled ? t.Foreground : t.Muted;
            item.Padding = new Padding(0, 1, 0, 1);
            item.Enabled = enabled;
            if (shortcut != null) item.ShortcutKeyDisplayString = shortcut;
            item.Click += click;
            menu.Items.Add(item);
        }

        static void CopyText(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            text = text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");
            try { Clipboard.SetText(text); }
            catch (ExternalException ex) { Log.Error(ex); }
        }

        // Recortar: copia; apaga só se a seleção estiver dentro da linha de comando
        void Cut()
        {
            if (box.SelectionLength == 0) return;
            CopyText(box.SelectedText);
            if (promptStart >= 0 && box.SelectionStart >= inputStart) box.SelectedText = "";
        }

        // Colar: só o texto, numa linha (as quebras viram "; "), na linha de comando
        void PastePlain()
        {
            string text = null;
            try { if (Clipboard.ContainsText()) text = Clipboard.GetText(); }
            catch (ExternalException) { return; }
            if (string.IsNullOrEmpty(text) || !PrepareEdit()) return;
            box.SelectedText = text.Replace("\r\n", "\n").Replace('\r', '\n').Trim('\n').Replace("\n", "; ");
        }

        // exit: como fechar um terminal. A sessão do PowerShell acaba (variáveis, pasta atual), a tela
        // e o histórico de comandos são apagados e a janela fecha; uma sessão nova já começa, para
        // o Terminal reabrir limpo
        // Botão "Fechar Terminal" (a lixeira da faixa das abas): o mesmo que digitar exit
        public void Kill() { Exit(); }

        void Exit()
        {
            Shutdown();
            history.Clear();
            historyIndex = 0;
            ClearAll();
            TerminalWindow.Hide();
            EnsureRunspace();
        }

        // cls, clear, Ctrl+L: limpa a tela; com o prompt na tela, ele volta com o que estava digitado
        void ClearScreen()
        {
            bool hadPrompt = promptStart >= 0;
            string typed = InputText;
            ClearAll();
            if (!hadPrompt) return;
            NewPrompt();
            InputText = typed;
            box.Select(box.TextLength, 0);
        }

        void ClearAll()
        {
            writing = true;
            box.Clear();
            writing = false;
            promptStart = -1;
            endsWithNewLine = true;
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

        // A caixa da tela: fica com Enter, Tab, Esc e as setas (o VBE e o WinForms as usariam para navegar)
        sealed class TerminalBox : RichTextBox
        {
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
                int started = Environment.TickCount;
                try { return app.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, app, all); }
                catch (TargetInvocationException ex)
                {
                    // O erro já foi para o Terminal com o lugar e a linha (Erros no Terminal): sem repetir
                    if (ErrorToTerminal.ReportedSince(started)) return null;
                    throw ex.InnerException ?? ex;
                }
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
