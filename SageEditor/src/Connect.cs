// SageEditor - add-in do editor do VBA (VBE), carregado dentro do Excel.
//
// Roda código .NET no processo do Excel, independente do VBA: continua funcionando
// com o código em modo de interrupção ou resetado.
//   - Menu "Sage" > "Configurações..." (tela no estilo do VS Code)
//   - Temas de cores para todo o VBE (ThemeEngine)
//   - Abas das janelas abertas no topo da área de código (EditorTabs)
//   - Vários cursores na janela de código (MultiCursor)
//   - Comando Clear na Verificação imediata (ImmediateCommands)

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageEditor
{
    // IDTExtensibility2 da biblioteca "Microsoft Add-In Designer" (MSADDNDR.DLL)
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int connectMode, [MarshalAs(UnmanagedType.IDispatch)] object addInInst, ref Array custom);
        [DispId(2)] void OnDisconnection(int removeMode, ref Array custom);
        [DispId(3)] void OnAddInsUpdate(ref Array custom);
        [DispId(4)] void OnStartupComplete(ref Array custom);
        [DispId(5)] void OnBeginShutdown(ref Array custom);
    }

    [ComVisible(true), Guid("3C1D5E7A-9B2F-4A6C-8E41-7F0A2B9D6C53"), ProgId("Sage.Editor"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2
    {
        const int ext_cm_Startup = 1;

        dynamic vbe;
        Control ui; // criado na thread do Excel, para voltar a ela com BeginInvoke
        SageMenu menu;
        Timer lineTimer; // números de linha: acompanha a linha atual e a rolagem
        bool started;

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            try
            {
                vbe = application;
                ((dynamic)addInInst).Object = this;
                ui = new Control();
                ui.CreateControl();
                if (connectMode != ext_cm_Startup)
                    Start();
                else // normalmente o VBE chama OnStartupComplete; isto é só a garantia
                    ui.BeginInvoke((MethodInvoker)SafeStart);
            }
            catch (Exception ex) { Log.Error(ex); }
        }

        public void OnStartupComplete(ref Array custom)
        {
            SafeStart();
        }

        void SafeStart()
        {
            try { Start(); }
            catch (Exception ex) { Log.Error(ex); }
        }

        void Start()
        {
            if (started || vbe == null) return;
            started = true;

            Settings.Load();
            Strings.Load();
            LineNumbers.Vbe = vbe;
            IntPtr main = new IntPtr((long)vbe.MainWindow.HWnd);
            ThemeEngine.Initialize(main);
            ThemeEngine.Apply(Theme.Find(Settings.ColorTheme));
            try { EditorTabs.Start(main); }
            catch (Exception ex) { Log.Error(ex); }
            MultiCursor.Vbe = vbe;
            MultiCursor.Start(main);
            ImmediateCommands.Vbe = vbe;
            ImmediateCommands.Start(main);

            menu = new SageMenu((object)vbe, new Action(OpenSettings));
            try { Syntax.Scan(vbe); }
            catch (Exception ex) { Log.Error(ex); }

            lineTimer = new Timer();
            lineTimer.Interval = 150;
            lineTimer.Tick += delegate
            {
                try { LineNumbers.Poll(); }
                catch (Exception) { } // VBE ocupado (ex.: executando código)
                try { EditorTabs.Poll(); ThemeEngine.PollForms(); MultiCursor.Poll(); }
                catch (Exception ex) { Log.Error(ex); }
                try { AutoReference.Poll(vbe); ImmediateCommands.Poll(); }
                catch (Exception) { } // VBE ocupado; tenta no próximo ciclo
            };
            lineTimer.Start();
            Log.Info("Iniciado. Tema: " + Settings.ColorTheme + ". Idioma: " + Strings.Language);
        }

        public void OnDisconnection(int removeMode, ref Array custom)
        {
            try
            {
                SettingsForm.CloseWindow();
                if (menu != null) { menu.Dispose(); menu = null; }
                if (lineTimer != null) { lineTimer.Dispose(); lineTimer = null; }
                LineNumbers.Vbe = null;
                MultiCursor.Shutdown();
                MultiCursor.Vbe = null;
                ImmediateCommands.Shutdown();
                AutoReference.Shutdown();
                ImmediateCommands.Vbe = null;
                EditorTabs.Shutdown();
                ThemeEngine.Shutdown();
                if (ui != null) { ui.Dispose(); ui = null; }
                started = false;

                // Referências COM presas impedem o Excel de fechar
                vbe = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception ex) { Log.Error(ex); }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ------------------------------------------------------------------
        // Também acessíveis por automação: VBE.Addins("Sage.Editor").Object
        // ------------------------------------------------------------------

        public void OpenSettings()
        {
            SettingsForm.ShowWindow(ApplyTheme);
        }

        public string Diagnostics() { return ThemeEngine.Diagnostics(); }

        public string SetTheme(string name)
        {
            Theme theme = Theme.Find(name);
            Settings.ColorTheme = theme.Name;
            // Chamadas por automação chegam numa thread do COM; as janelas do VBE
            // só podem ser subclassificadas na thread delas.
            Control target = ui;
            if (target != null && target.InvokeRequired)
                target.Invoke((MethodInvoker)delegate { Apply(theme); });
            else
                Apply(theme);
            return theme.Name;
        }

        static void Apply(Theme theme)
        {
            ThemeEngine.Apply(theme);
            EditorTabs.Refresh(); // o tema padrão não passa pelo ThemeEngine
        }

        // Chamado pela thread da tela de Configurações
        void ApplyTheme(Theme theme)
        {
            Control target = ui;
            if (target == null) return;
            target.BeginInvoke((MethodInvoker)delegate
            {
                try { Apply(theme); }
                catch (Exception ex) { Log.Error(ex); }
            });
        }
    }
}
