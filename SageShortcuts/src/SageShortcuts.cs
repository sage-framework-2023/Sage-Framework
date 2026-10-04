// SageShortcuts - atalhos de teclado (com acordes, estilo VS Code) para o editor do VBA.
//
// Roda fora do Excel: instala um hook WH_KEYBOARD_LL, reconhece os atalhos de
// keybindings.txt quando o VBE é a janela ativa e executa a ação via COM.
// Como nenhum código VBA roda dentro do hook, o Excel não trava em modo de
// interrupção nem quando o projeto é resetado.
//
// Carregado por SageShortcuts.ps1 (Add-Type, C# 5), que chama Program.Main.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace SageShortcuts
{
    public static class Program
    {
        const string MutexName = @"Local\SageShortcuts.Instance";
        internal const string StopEventName = @"Local\SageShortcuts.Stop";

        [STAThread]
        public static int Main(string[] args)
        {
            int parentPid = 0;
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "keybindings.txt");
            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i].ToLowerInvariant();
                if (arg == "--stop")
                {
                    SignalStop();
                    return 0;
                }
                if (arg == "--parent" && i + 1 < args.Length)
                    int.TryParse(args[++i], out parentPid);
                else if (arg == "--config" && i + 1 < args.Length)
                    configPath = Path.GetFullPath(args[++i]);
            }

            bool created;
            using (Mutex mutex = new Mutex(true, MutexName, out created))
            {
                if (!created) return 1; // já existe uma instância rodando
                using (EventWaitHandle stopEvent = new EventWaitHandle(false, EventResetMode.ManualReset, StopEventName))
                {
                    Application.EnableVisualStyles();
                    Application.Run(new TrayContext(stopEvent, parentPid, configPath));
                }
            }
            return 0;
        }

        static void SignalStop()
        {
            try
            {
                using (EventWaitHandle ev = EventWaitHandle.OpenExisting(StopEventName))
                    ev.Set();
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                // nada rodando
            }
        }
    }

    // ------------------------------------------------------------------------
    // Ícone na bandeja + ciclo de vida
    // ------------------------------------------------------------------------
    sealed class TrayContext : ApplicationContext
    {
        readonly Control invoker;
        readonly NotifyIcon tray;
        readonly ToolStripMenuItem enabledItem;
        readonly KeyboardHook hook;
        readonly ShortcutEngine engine;
        readonly FileSystemWatcher watcher;
        readonly System.Windows.Forms.Timer reloadTimer;
        readonly string configPath;
        RegisteredWaitHandle stopWait;
        Process parent;

        public TrayContext(EventWaitHandle stopEvent, int parentPid, string configPath)
        {
            invoker = new Control();
            invoker.CreateControl(); // handle criado nesta thread, para BeginInvoke

            this.configPath = configPath;

            enabledItem = new ToolStripMenuItem("Ativo", null, delegate { SetEnabled(!engine.Enabled); });
            enabledItem.Checked = true;
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add(enabledItem);
            menu.Items.Add("Recarregar atalhos", null, delegate { Reload(true); });
            menu.Items.Add("Editar atalhos", null, delegate { OpenConfig(); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Sair", null, delegate { ExitThread(); });

            tray = new NotifyIcon();
            tray.Icon = SystemIcons.Application;
            tray.Text = "Atalhos do VBE";
            tray.ContextMenuStrip = menu;
            tray.Visible = true;

            engine = new ShortcutEngine(OnBindingMatched);
            Reload(false);

            hook = new KeyboardHook(engine.OnKey);
            hook.Start();

            // Recarrega sozinho quando keybindings.txt é salvo
            reloadTimer = new System.Windows.Forms.Timer();
            reloadTimer.Interval = 300;
            reloadTimer.Tick += delegate { reloadTimer.Stop(); Reload(true); };
            watcher = new FileSystemWatcher(Path.GetDirectoryName(configPath), Path.GetFileName(configPath));
            watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size;
            watcher.SynchronizingObject = invoker;
            watcher.Changed += delegate { reloadTimer.Stop(); reloadTimer.Start(); };
            watcher.Created += delegate { reloadTimer.Stop(); reloadTimer.Start(); };
            watcher.Renamed += delegate { reloadTimer.Stop(); reloadTimer.Start(); };
            watcher.EnableRaisingEvents = true;

            // Encerramento: "SageShortcuts.exe --stop" ou fim do processo do Excel
            stopWait = ThreadPool.RegisterWaitForSingleObject(stopEvent,
                delegate { invoker.BeginInvoke((MethodInvoker)ExitThread); }, null, -1, true);

            if (parentPid != 0)
            {
                try
                {
                    parent = Process.GetProcessById(parentPid);
                    parent.EnableRaisingEvents = true;
                    parent.Exited += delegate { invoker.BeginInvoke((MethodInvoker)ExitThread); };
                    if (parent.HasExited) invoker.BeginInvoke((MethodInvoker)ExitThread);
                }
                catch (ArgumentException)
                {
                    invoker.BeginInvoke((MethodInvoker)ExitThread);
                }
            }
        }

        void SetEnabled(bool value)
        {
            engine.Enabled = value;
            enabledItem.Checked = value;
            tray.Text = value ? "Atalhos do VBE" : "Atalhos do VBE (pausado)";
        }

        void Reload(bool notify)
        {
            List<string> errors;
            Bindings bindings = Bindings.Load(configPath, out errors);
            engine.Bindings = bindings;
            if (errors.Count > 0)
                Notify(ToolTipIcon.Warning, "keybindings.txt", string.Join("\n", errors.ToArray()));
            else if (notify)
                Notify(ToolTipIcon.Info, "Atalhos recarregados", bindings.Count + " atalho(s) ativo(s).");
        }

        void OpenConfig()
        {
            try { Process.Start("notepad.exe", "\"" + configPath + "\""); }
            catch (Exception ex) { Notify(ToolTipIcon.Error, "Editar atalhos", ex.Message); }
        }

        // Chamado na thread do hook: só agenda, nunca executa ali.
        void OnBindingMatched(Binding binding, IntPtr vbeWindow)
        {
            invoker.BeginInvoke((MethodInvoker)delegate
            {
                try
                {
                    VbeActions.Execute(binding, vbeWindow);
                }
                catch (Exception ex)
                {
                    Notify(ToolTipIcon.Warning, binding.Keys, ex.Message);
                }
            });
        }

        void Notify(ToolTipIcon icon, string title, string text)
        {
            if (text.Length > 250) text = text.Substring(0, 250) + "...";
            tray.ShowBalloonTip(4000, title, text, icon);
        }

        protected override void ExitThreadCore()
        {
            if (stopWait != null) { stopWait.Unregister(null); stopWait = null; }
            hook.Stop();
            watcher.Dispose();
            reloadTimer.Dispose();
            tray.Visible = false;
            tray.Dispose();
            if (parent != null) parent.Dispose();
            invoker.Dispose();
            base.ExitThreadCore();
        }
    }

    // ------------------------------------------------------------------------
    // Atalhos (keybindings.txt)
    // ------------------------------------------------------------------------
    enum ActionKind { VbeCommand, Macro, Comment, Uncomment, CopyLinesUp, CopyLinesDown, MoveLinesUp, MoveLinesDown, ToggleImmediate, ToggleWatch, ToggleLocals, SendKeys }

    sealed class Binding
    {
        public string Keys;        // texto original, para mensagens
        public ActionKind Kind;
        public int CommandId;      // vbe:<id>
        public string Macro;       // macro:<nome>
        public string SendKeys;    // keys:<tecla>: acordes codificados como no KeyChord ("mods:vk,...")
    }

    sealed class Bindings
    {
        readonly Dictionary<string, Binding> exact = new Dictionary<string, Binding>();
        readonly HashSet<string> prefixes = new HashSet<string>();

        public int Count { get { return exact.Count; } }

        public Binding Find(string sequence)
        {
            Binding b;
            return exact.TryGetValue(sequence, out b) ? b : null;
        }

        public bool IsPrefix(string sequence)
        {
            return prefixes.Contains(sequence);
        }

        // Formato:  <tecla>[, <tecla>...] = <ação>     # comentário
        //   tecla: [Ctrl+][Shift+][Alt+][Win+]<nome>   (ex.: Ctrl+K, Ctrl+Shift+F2)
        //   ação:  comment | uncomment | copyLinesUp | copyLinesDown | moveLinesUp | moveLinesDown |
        //          toggleImmediate | toggleWatch | toggleLocals | vbe:<id do controle> | macro:<nome para Application.Run> |
        //          keys:<tecla> (envia outra combinação ao VBE, ex.: keys:Shift+F2)
        public static Bindings Load(string path, out List<string> errors)
        {
            Bindings result = new Bindings();
            errors = new List<string>();
            if (!File.Exists(path))
            {
                errors.Add("Arquivo não encontrado: " + path);
                return result;
            }

            string[] lines;
            try { lines = File.ReadAllLines(path, Encoding.UTF8); }
            catch (IOException ex) { errors.Add(ex.Message); return result; }

            for (int n = 0; n < lines.Length; n++)
            {
                string line = lines[n];
                int hash = line.IndexOf('#');
                if (hash >= 0) line = line.Substring(0, hash);
                line = line.Trim();
                if (line.Length == 0) continue;

                string where = "Linha " + (n + 1) + ": ";
                int eq = line.IndexOf('=');
                if (eq < 0) { errors.Add(where + "falta '='"); continue; }

                string keysText = line.Substring(0, eq).Trim();
                string actionText = line.Substring(eq + 1).Trim();

                string sequence, error;
                if (!KeyChord.ParseSequence(keysText, out sequence, out error)) { errors.Add(where + error); continue; }

                Binding binding = ParseAction(actionText, out error);
                if (binding == null) { errors.Add(where + error); continue; }
                binding.Keys = keysText;

                if (result.exact.ContainsKey(sequence))
                    errors.Add(where + "'" + keysText + "' repetido; vale a última definição");
                result.exact[sequence] = binding;

                string[] parts = sequence.Split(',');
                for (int i = 1; i < parts.Length; i++)
                    result.prefixes.Add(string.Join(",", parts, 0, i));
            }
            return result;
        }

        static Binding ParseAction(string text, out string error)
        {
            error = null;
            switch (text.ToLowerInvariant())
            {
                case "comment": { Binding c = new Binding(); c.Kind = ActionKind.Comment; return c; }
                case "uncomment": { Binding u = new Binding(); u.Kind = ActionKind.Uncomment; return u; }
                case "copylinesup": { Binding up = new Binding(); up.Kind = ActionKind.CopyLinesUp; return up; }
                case "copylinesdown": { Binding down = new Binding(); down.Kind = ActionKind.CopyLinesDown; return down; }
                case "movelinesup": { Binding mu = new Binding(); mu.Kind = ActionKind.MoveLinesUp; return mu; }
                case "movelinesdown": { Binding md = new Binding(); md.Kind = ActionKind.MoveLinesDown; return md; }
                case "toggleimmediate": { Binding ti = new Binding(); ti.Kind = ActionKind.ToggleImmediate; return ti; }
                case "togglewatch": { Binding tw = new Binding(); tw.Kind = ActionKind.ToggleWatch; return tw; }
                case "togglelocals": { Binding tl = new Binding(); tl.Kind = ActionKind.ToggleLocals; return tl; }
            }
            int colon = text.IndexOf(':');
            string kind = colon < 0 ? "" : text.Substring(0, colon).Trim().ToLowerInvariant();
            string arg = colon < 0 ? "" : text.Substring(colon + 1).Trim();
            Binding b = new Binding();

            if (kind == "vbe")
            {
                int id;
                if (!int.TryParse(arg, out id)) { error = "id inválido em '" + text + "'"; return null; }
                b.Kind = ActionKind.VbeCommand;
                b.CommandId = id;
                return b;
            }
            if (kind == "keys" && arg.Length > 0)
            {
                string sequence;
                if (!KeyChord.ParseSequence(arg, out sequence, out error)) { error = error + " em '" + text + "'"; return null; }
                b.Kind = ActionKind.SendKeys;
                b.SendKeys = sequence;
                return b;
            }
            if (kind == "macro" && arg.Length > 0)
            {
                b.Kind = ActionKind.Macro;
                b.Macro = arg;
                return b;
            }
            error = "ação inválida '" + text + "' (use comment, uncomment, copyLinesUp, copyLinesDown, moveLinesUp, moveLinesDown, toggleImmediate, toggleWatch, toggleLocals, vbe:<id>, macro:<nome> ou keys:<tecla>)";
            return null;
        }
    }

    static class KeyChord
    {
        public const int Ctrl = 1, Shift = 2, Alt = 4, Win = 8;

        // Um acorde vira "mods:vk"; uma sequência, acordes separados por vírgula.
        public static string Encode(int vk, int mods)
        {
            return mods + ":" + vk;
        }

        public static bool ParseSequence(string text, out string sequence, out string error)
        {
            sequence = null;
            error = null;
            string[] chords = text.Split(new char[] { ',', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (chords.Length == 0) { error = "atalho vazio"; return false; }

            List<string> encoded = new List<string>();
            foreach (string chord in chords)
            {
                string[] parts = chord.Split('+');
                int mods = 0;
                for (int i = 0; i < parts.Length - 1; i++)
                {
                    switch (parts[i].Trim().ToLowerInvariant())
                    {
                        case "ctrl": case "control": mods |= Ctrl; break;
                        case "shift": mods |= Shift; break;
                        case "alt": mods |= Alt; break;
                        case "win": mods |= Win; break;
                        default: error = "modificador desconhecido '" + parts[i] + "'"; return false;
                    }
                }
                int vk = ParseKey(parts[parts.Length - 1].Trim());
                if (vk == 0) { error = "tecla desconhecida '" + parts[parts.Length - 1] + "'"; return false; }
                encoded.Add(Encode(vk, mods));
            }
            sequence = string.Join(",", encoded.ToArray());
            return true;
        }

        static int ParseKey(string name)
        {
            if (name.Length == 1)
            {
                char c = char.ToUpperInvariant(name[0]);
                if ((c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9')) return c;
            }
            switch (name.ToLowerInvariant())
            {
                case "esc": return (int)Keys.Escape;
                case "del": return (int)Keys.Delete;
                case "ins": return (int)Keys.Insert;
                case "pgup": return (int)Keys.PageUp;
                case "pgdn": return (int)Keys.PageDown;
                case "backspace": return (int)Keys.Back;
            }
            Keys key;
            if (Enum.TryParse<Keys>(name, true, out key) && (key & Keys.Modifiers) == 0)
                return (int)key;
            return 0;
        }
    }

    // ------------------------------------------------------------------------
    // Reconhecimento dos acordes (roda na thread do hook: tem de ser rápido)
    // ------------------------------------------------------------------------
    sealed class ShortcutEngine
    {
        const int ChordTimeoutMs = 3000;

        readonly Action<Binding, IntPtr> onMatch;
        readonly HashSet<int> swallowed = new HashSet<int>(); // teclas cujo keydown foi consumido
        volatile Bindings bindings = new Bindings();
        volatile bool enabled = true;
        string pending;       // prefixo já digitado (ex.: Ctrl+K)
        int pendingTick;

        public ShortcutEngine(Action<Binding, IntPtr> onMatch)
        {
            this.onMatch = onMatch;
        }

        public Bindings Bindings { get { return bindings; } set { bindings = value; } }
        public bool Enabled { get { return enabled; } set { enabled = value; } }

        // Devolve true para consumir a tecla (o VBE não a recebe).
        public bool OnKey(int vk, bool down)
        {
            if (!down)
                return swallowed.Remove(vk); // keyup de uma tecla consumida também é consumido

            if (IsModifier(vk)) return false;
            if (swallowed.Contains(vk)) return true; // auto-repetição de tecla consumida

            IntPtr vbe;
            if (!enabled || !Native.IsVbeForeground(out vbe))
            {
                pending = null;
                return false;
            }

            string chord = KeyChord.Encode(vk, CurrentModifiers());
            if (pending != null && unchecked(Environment.TickCount - pendingTick) > ChordTimeoutMs)
                pending = null;

            if (pending != null)
            {
                string sequence = pending + "," + chord;
                pending = null;
                if (Handle(sequence, vk, vbe)) return true;
                // O acorde não continuou nenhum atalho: trata a tecla como início de outro.
            }
            return Handle(chord, vk, vbe);
        }

        bool Handle(string sequence, int vk, IntPtr vbe)
        {
            Bindings current = bindings;
            if (current.IsPrefix(sequence))
            {
                pending = sequence;
                pendingTick = Environment.TickCount;
                swallowed.Add(vk);
                MaskAlt();
                return true;
            }
            Binding binding = current.Find(sequence);
            if (binding == null) return false;
            swallowed.Add(vk);
            MaskAlt();
            onMatch(binding, vbe);
            return true;
        }

        // Com a tecla consumida, o VBE veria só "Alt pressionado e solto" e ativaria a
        // barra de menus. Uma tecla neutra no meio evita isso (o hook ignora teclas
        // sintéticas, então ela não volta para cá).
        static void MaskAlt()
        {
            if (Native.IsDown(Keys.Menu) || Native.IsDown(Keys.LWin) || Native.IsDown(Keys.RWin))
                Native.TapMaskKey();
        }

        static bool IsModifier(int vk)
        {
            switch ((Keys)vk)
            {
                case Keys.ControlKey: case Keys.LControlKey: case Keys.RControlKey:
                case Keys.ShiftKey: case Keys.LShiftKey: case Keys.RShiftKey:
                case Keys.Menu: case Keys.LMenu: case Keys.RMenu:
                case Keys.LWin: case Keys.RWin:
                    return true;
            }
            return false;
        }

        static int CurrentModifiers()
        {
            int mods = 0;
            if (Native.IsDown(Keys.ControlKey)) mods |= KeyChord.Ctrl;
            if (Native.IsDown(Keys.ShiftKey)) mods |= KeyChord.Shift;
            if (Native.IsDown(Keys.Menu)) mods |= KeyChord.Alt;
            if (Native.IsDown(Keys.LWin) || Native.IsDown(Keys.RWin)) mods |= KeyChord.Win;
            return mods;
        }
    }

    // ------------------------------------------------------------------------
    // Hook de teclado de baixo nível, em thread própria com loop de mensagens
    // ------------------------------------------------------------------------
    sealed class KeyboardHook
    {
        readonly Func<int, bool, bool> handler;
        readonly ManualResetEvent ready = new ManualResetEvent(false);
        Native.LowLevelKeyboardProc proc; // mantém o delegate vivo enquanto o hook existe
        Thread thread;
        uint threadId;
        IntPtr hook;

        public KeyboardHook(Func<int, bool, bool> handler)
        {
            this.handler = handler;
        }

        public void Start()
        {
            thread = new Thread(Run);
            thread.IsBackground = true;
            thread.Name = "KeyboardHook";
            thread.Start();
            ready.WaitOne();
            if (hook == IntPtr.Zero)
                throw new Win32Exception(Marshal.GetLastWin32Error(), "SetWindowsHookEx falhou");
        }

        public void Stop()
        {
            if (thread == null) return;
            Native.PostThreadMessage(threadId, Native.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            thread.Join(2000);
            thread = null;
        }

        void Run()
        {
            threadId = Native.GetCurrentThreadId();
            Native.MSG msg;
            Native.PeekMessage(out msg, IntPtr.Zero, 0, 0, 0); // garante a fila de mensagens
            proc = HookProc;
            hook = Native.SetWindowsHookEx(Native.WH_KEYBOARD_LL, proc, Native.GetModuleHandle(null), 0);
            ready.Set();
            if (hook == IntPtr.Zero) return;

            while (Native.GetMessage(out msg, IntPtr.Zero, 0, 0) > 0)
            {
                Native.TranslateMessage(ref msg);
                Native.DispatchMessage(ref msg);
            }
            Native.UnhookWindowsHookEx(hook);
            hook = IntPtr.Zero;
        }

        IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                Native.KBDLLHOOKSTRUCT k = (Native.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Native.KBDLLHOOKSTRUCT));
                // Ignora teclas sintéticas (SendKeys, SendInput de outros programas)
                if ((k.flags & Native.LLKHF_INJECTED) == 0)
                {
                    int msg = wParam.ToInt32();
                    bool down = msg == Native.WM_KEYDOWN || msg == Native.WM_SYSKEYDOWN;
                    bool swallow = false;
                    try { swallow = handler((int)k.vkCode, down); }
                    catch { }
                    if (swallow) return (IntPtr)1;
                }
            }
            return Native.CallNextHookEx(hook, nCode, wParam, lParam);
        }
    }

    // ------------------------------------------------------------------------
    // Execução das ações no Excel dono do VBE (via COM, fora do processo)
    // ------------------------------------------------------------------------
    static class VbeActions
    {
        const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);
        const int RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A);

        public static void Execute(Binding binding, IntPtr vbeWindow)
        {
            // Só teclas: não precisa do Excel (o VBE recebe a combinação como se fosse digitada;
            // o hook ignora teclas sintéticas, então não há laço)
            if (binding.Kind == ActionKind.SendKeys)
            {
                foreach (string chord in binding.SendKeys.Split(','))
                {
                    string[] p = chord.Split(':');
                    Native.SendChord(int.Parse(p[1]), int.Parse(p[0]));
                }
                return;
            }
            for (int attempt = 0; ; attempt++)
            {
                try
                {
                    ExecuteOnce(binding, vbeWindow);
                    return;
                }
                catch (COMException ex)
                {
                    bool busy = ex.ErrorCode == RPC_E_CALL_REJECTED || ex.ErrorCode == RPC_E_SERVERCALL_RETRYLATER;
                    if (!busy || attempt >= 10) throw;
                    Thread.Sleep(50);
                }
                finally
                {
                    // Solta as referências COM para não segurar o Excel aberto
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        static void ExecuteOnce(Binding binding, IntPtr vbeWindow)
        {
            dynamic app = ExcelFromVbe(vbeWindow);
            if (app == null) throw new InvalidOperationException("Excel não encontrado.");

            switch (binding.Kind)
            {
                case ActionKind.VbeCommand:
                    dynamic control = app.VBE.CommandBars.FindControl(Type.Missing, binding.CommandId);
                    if (control == null)
                        throw new InvalidOperationException("Comando " + binding.CommandId + " não existe no VBE.");
                    if (control.Enabled) control.Execute();
                    break;
                case ActionKind.Macro:
                    app.Run(binding.Macro);
                    break;
                case ActionKind.Comment:
                case ActionKind.Uncomment:
                    // Edita pelo CodeModule (não usa os comandos nativos 192/2552): funciona em
                    // modo de interrupção e põe o comentário no recuo, como o VS Code.
                    // O Desfazer do VBE não registra essas edições.
                    ToggleComment(app.VBE.ActiveCodePane, binding.Kind == ActionKind.Comment);
                    break;
                case ActionKind.CopyLinesUp:
                case ActionKind.CopyLinesDown:
                    CopyLines(app.VBE.ActiveCodePane, binding.Kind == ActionKind.CopyLinesDown);
                    break;
                case ActionKind.MoveLinesUp:
                case ActionKind.MoveLinesDown:
                    MoveLines(app.VBE.ActiveCodePane, binding.Kind == ActionKind.MoveLinesDown);
                    break;
                case ActionKind.ToggleImmediate:
                    ToggleWindow(app.VBE, vbext_wt_Immediate, 2554);
                    break;
                case ActionKind.ToggleWatch:
                    ToggleWindow(app.VBE, vbext_wt_Watch, 2556);
                    break;
                case ActionKind.ToggleLocals:
                    ToggleWindow(app.VBE, vbext_wt_Locals, 2555);
                    break;
            }
        }

        const int vbext_wt_Watch = 3, vbext_wt_Locals = 4, vbext_wt_Immediate = 5;

        // Fecha a janela (Verificação imediata, Inspeção de variáveis, Variáveis locais) se
        // estiver aberta; senão abre e põe o foco nela, como Ctrl+J no VS Code (que alterna o
        // painel). Se ela ainda não existir em VBE.Windows, abre pelo comando do menu Exibir.
        static void ToggleWindow(dynamic vbe, int type, int commandId)
        {
            foreach (dynamic window in vbe.Windows)
            {
                if ((int)window.Type != type) continue;
                if ((bool)window.Visible) window.Visible = false;
                else
                {
                    window.Visible = true;
                    window.SetFocus();
                }
                return;
            }
            dynamic control = vbe.CommandBars.FindControl(Type.Missing, commandId);
            if (control != null && (bool)control.Enabled) control.Execute();
        }

        // Troca as linhas da seleção com a de cima ou a de baixo, como Alt+Seta no VS Code;
        // a seleção acompanha. Também pelo CodeModule (sem Desfazer).
        static void MoveLines(dynamic pane, bool down)
        {
            if (pane == null) return;
            dynamic module = pane.CodeModule;
            int total = module.CountOfLines;
            if (total == 0) return;
            int startLine = 0, startCol = 0, endLine = 0, endCol = 0;
            pane.GetSelection(ref startLine, ref startCol, ref endLine, ref endCol);
            int lastLine = (endLine > startLine && endCol == 1) ? endLine - 1 : endLine;
            if (down ? lastLine >= total : startLine <= 1) return;

            if (down)
            {
                string below = module.Lines(lastLine + 1, 1);
                module.DeleteLines(lastLine + 1, 1);
                module.InsertLines(startLine, below);
                pane.SetSelection(startLine + 1, startCol, endLine + 1, endCol);
            }
            else
            {
                string above = module.Lines(startLine - 1, 1);
                module.DeleteLines(startLine - 1, 1);
                module.InsertLines(lastLine, above); // o bloco subiu uma linha; a de cima vai para depois dele
                pane.SetSelection(startLine - 1, startCol, endLine - 1, endCol);
            }
        }

        // Duplica as linhas da seleção, como Shift+Alt+Seta no VS Code. Para baixo, a
        // seleção vai para a cópia; para cima, fica na cópia (que ocupa o lugar original).
        // Também pelo CodeModule: funciona em modo de interrupção, sem Desfazer.
        static void CopyLines(dynamic pane, bool down)
        {
            if (pane == null) return;
            dynamic module = pane.CodeModule;
            if (module.CountOfLines == 0) return;
            int startLine = 0, startCol = 0, endLine = 0, endCol = 0;
            pane.GetSelection(ref startLine, ref startCol, ref endLine, ref endCol);
            int lastLine = (endLine > startLine && endCol == 1) ? endLine - 1 : endLine;
            int count = lastLine - startLine + 1;
            string text = module.Lines(startLine, count);

            if (down)
            {
                module.InsertLines(lastLine + 1, text);
                pane.SetSelection(startLine + count, startCol, endLine + count, endCol);
            }
            else
            {
                module.InsertLines(startLine, text);
                pane.SetSelection(startLine, startCol, endLine, endCol);
            }
        }

        const string CommentToken = "' ";

        static void ToggleComment(dynamic pane, bool comment)
        {
            if (pane == null) return;
            int startLine = 0, startCol = 0, endLine = 0, endCol = 0;
            pane.GetSelection(ref startLine, ref startCol, ref endLine, ref endCol);
            // Como no VS Code: seleção que termina na coluna 1 não inclui essa linha
            int lastLine = (endLine > startLine && endCol == 1) ? endLine - 1 : endLine;

            dynamic module = pane.CodeModule;
            int count = lastLine - startLine + 1;
            string text = module.Lines(startLine, count);
            string[] lines = text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            if (lines.Length != count) return;

            // Coluna do comentário: o menor recuo entre as linhas não vazias
            int column = int.MaxValue;
            foreach (string line in lines)
                if (line.Trim().Length > 0) column = Math.Min(column, Indent(line));
            if (column == int.MaxValue) return; // só linhas vazias

            int[] at = new int[count];     // posição (base 0) da inserção/remoção
            int[] delta = new int[count];  // caracteres inseridos (+) ou removidos (-)
            for (int i = 0; i < count; i++)
            {
                string line = lines[i];
                string changed = line;
                if (comment)
                {
                    if (line.Trim().Length > 0)
                    {
                        changed = line.Insert(column, CommentToken);
                        at[i] = column;
                    }
                }
                else
                {
                    int pos = Indent(line);
                    if (pos < line.Length && line[pos] == '\'')
                    {
                        // Tira também o espaço depois do apóstrofo ("' x" -> "x"), exceto num
                        // comentário no estilo do VBE, com o apóstrofo na coluna 1 seguido do
                        // recuo original ("'    x" -> "    x").
                        int length = 1;
                        bool vbeStyle = pos == 0 && pos + 2 < line.Length && char.IsWhiteSpace(line[pos + 2]);
                        if (pos + 1 < line.Length && line[pos + 1] == ' ' && !vbeStyle)
                            length = 2;
                        changed = line.Remove(pos, length);
                        at[i] = pos;
                    }
                }
                delta[i] = changed.Length - line.Length;
                if (delta[i] != 0) module.ReplaceLine(startLine + i, changed);
            }

            startCol = ShiftColumn(startCol, 0, at, delta);
            if (endLine <= lastLine) endCol = ShiftColumn(endCol, endLine - startLine, at, delta);
            pane.SetSelection(startLine, startCol, endLine, endCol);
        }

        static int Indent(string line)
        {
            int i = 0;
            while (i < line.Length && (line[i] == ' ' || line[i] == '\t')) i++;
            return i;
        }

        // Ajusta uma coluna (base 1) da seleção à edição feita na linha.
        static int ShiftColumn(int col, int index, int[] at, int[] delta)
        {
            if (index < 0 || index >= at.Length || delta[index] == 0 || col - 1 < at[index]) return col;
            return Math.Max(at[index] + 1, col + delta[index]);
        }

        // Pega o Excel do mesmo processo do VBE (funciona com várias instâncias abertas).
        static object ExcelFromVbe(IntPtr vbeWindow)
        {
            uint pid;
            Native.GetWindowThreadProcessId(vbeWindow, out pid);

            IntPtr xlMain = IntPtr.Zero;
            Native.EnumWindows(delegate(IntPtr h, IntPtr l)
            {
                uint p;
                Native.GetWindowThreadProcessId(h, out p);
                if (p == pid && Native.ClassName(h) == "XLMAIN") { xlMain = h; return false; }
                return true;
            }, IntPtr.Zero);

            if (xlMain != IntPtr.Zero)
            {
                IntPtr desk = Native.FindWindowEx(xlMain, IntPtr.Zero, "XLDESK", null);
                IntPtr sheet = desk == IntPtr.Zero ? IntPtr.Zero : Native.FindWindowEx(desk, IntPtr.Zero, "EXCEL7", null);
                if (sheet != IntPtr.Zero)
                {
                    Guid iidDispatch = new Guid("00020400-0000-0000-C000-000000000046");
                    object window;
                    if (Native.AccessibleObjectFromWindow(sheet, Native.OBJID_NATIVEOM, ref iidDispatch, out window) == 0 && window != null)
                        return ((dynamic)window).Application;
                }
            }

            // Sem janela de pasta (só add-ins abertos): usa a instância registrada
            try { return Marshal.GetActiveObject("Excel.Application"); }
            catch (COMException) { return null; }
        }
    }

    // ------------------------------------------------------------------------
    // Win32
    // ------------------------------------------------------------------------
    static class Native
    {
        public const int WH_KEYBOARD_LL = 13;
        public const int WM_KEYDOWN = 0x0100, WM_SYSKEYDOWN = 0x0104;
        public const uint WM_QUIT = 0x0012;
        public const uint LLKHF_INJECTED = 0x10;
        public const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        const uint GA_ROOTOWNER = 3;

        const string VbeClass = "wndclass_desked_gsk";

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT
        {
            public uint vkCode, scanCode, flags, time;
            public UIntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam, lParam;
            public uint time;
            public int ptX, ptY;
        }

        public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetModuleHandle(string name);
        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        public static extern int GetMessage(out MSG msg, IntPtr hwnd, uint min, uint max);
        [DllImport("user32.dll")]
        public static extern bool PeekMessage(out MSG msg, IntPtr hwnd, uint min, uint max, uint remove);
        [DllImport("user32.dll")]
        public static extern bool TranslateMessage(ref MSG msg);
        [DllImport("user32.dll")]
        public static extern IntPtr DispatchMessage(ref MSG msg);
        [DllImport("user32.dll")]
        public static extern bool PostThreadMessage(uint threadId, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern short GetAsyncKeyState(int vk);
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern int GetClassName(IntPtr hwnd, StringBuilder name, int max);
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc proc, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string title);
        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objId, ref Guid riid,
            [MarshalAs(UnmanagedType.IDispatch)] out object obj);

        public static bool IsDown(Keys key)
        {
            return (GetAsyncKeyState((int)key) & 0x8000) != 0;
        }

        [DllImport("user32.dll")]
        static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extraInfo);
        const byte VK_MASK = 0xE8; // código de tecla sem uso
        const uint KEYEVENTF_KEYUP = 0x2;

        public static void TapMaskKey()
        {
            keybd_event(VK_MASK, 0, 0, UIntPtr.Zero);
            keybd_event(VK_MASK, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        }

        // Pressiona os modificadores, a tecla, e solta tudo na ordem inversa (keys:<tecla>)
        public static void SendChord(int vk, int mods)
        {
            List<byte> modifiers = new List<byte>();
            if ((mods & KeyChord.Ctrl) != 0) modifiers.Add((byte)Keys.ControlKey);
            if ((mods & KeyChord.Shift) != 0) modifiers.Add((byte)Keys.ShiftKey);
            if ((mods & KeyChord.Alt) != 0) modifiers.Add((byte)Keys.Menu);
            if ((mods & KeyChord.Win) != 0) modifiers.Add((byte)Keys.LWin);
            foreach (byte m in modifiers) keybd_event(m, 0, 0, UIntPtr.Zero);
            keybd_event((byte)vk, 0, 0, UIntPtr.Zero);
            keybd_event((byte)vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            for (int i = modifiers.Count - 1; i >= 0; i--) keybd_event(modifiers[i], 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        }

        public static string ClassName(IntPtr hwnd)
        {
            StringBuilder sb = new StringBuilder(64);
            GetClassName(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        // VBE ativo: a janela principal ou uma janela flutuante dele (não um diálogo).
        public static bool IsVbeForeground(out IntPtr vbe)
        {
            vbe = IntPtr.Zero;
            IntPtr fg = GetForegroundWindow();
            if (fg == IntPtr.Zero) return false;
            string cls = ClassName(fg);
            if (cls == VbeClass) { vbe = fg; return true; }
            if (cls == "#32770") return false; // Localizar, Referências, MsgBox...
            IntPtr root = GetAncestor(fg, GA_ROOTOWNER);
            if (root != IntPtr.Zero && ClassName(root) == VbeClass) { vbe = root; return true; }
            return false;
        }
    }
}
