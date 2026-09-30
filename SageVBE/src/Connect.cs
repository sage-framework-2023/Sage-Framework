// SageVBE - add-in do editor do VBA (VBE), carregado dentro do Excel.
//
// Roda código .NET no processo do Excel, independente do VBA: continua funcionando
// com o código em modo de interrupção ou resetado.
//   - Menu "Sage" > "Configurações..." (tela no estilo do VS Code)
//   - Temas de cores para todo o VBE (ThemeEngine)

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SageVBE
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

    [ComVisible(true), Guid("3C1D5E7A-9B2F-4A6C-8E41-7F0A2B9D6C53"), ProgId("Sage.VBE"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2
    {
        const int ext_cm_Startup = 1;

        dynamic vbe;
        Control ui; // criado na thread do Excel, para voltar a ela com BeginInvoke
        SageMenu menu;
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
            ThemeEngine.Initialize(new IntPtr((long)vbe.MainWindow.HWnd));
            ThemeEngine.Apply(Theme.Find(Settings.ColorTheme));

            menu = new SageMenu((object)vbe, new Action(OpenSettings));
            Log.Info("Iniciado. Tema: " + Settings.ColorTheme);
        }

        public void OnDisconnection(int removeMode, ref Array custom)
        {
            try
            {
                SettingsForm.CloseWindow();
                if (menu != null) { menu.Dispose(); menu = null; }
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
        // Também acessíveis por automação: VBE.Addins("Sage.VBE").Object
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
                target.Invoke((MethodInvoker)delegate { ThemeEngine.Apply(theme); });
            else
                ThemeEngine.Apply(theme);
            return theme.Name;
        }

        // Chamado pela thread da tela de Configurações
        void ApplyTheme(Theme theme)
        {
            Control target = ui;
            if (target == null) return;
            target.BeginInvoke((MethodInvoker)delegate
            {
                try { ThemeEngine.Apply(theme); }
                catch (Exception ex) { Log.Error(ex); }
            });
        }
    }
}
