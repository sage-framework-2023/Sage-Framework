using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using SYSKIND = System.Runtime.InteropServices.ComTypes.SYSKIND;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;

namespace SageSetup
{
    // Instalador do Sage Framework em um arquivo só (release do GitHub).
    //
    //   SageFramework-Setup.exe                instala (janela)
    //   SageFramework-Setup.exe /quiet         instala sem janela
    //   Uninstall.exe /uninstall [/quiet]      desinstala (é o mesmo programa, copiado na instalação)
    //
    // Os arquivos vêm num zip embutido (payload.zip, montado pelo build-release.ps1): os DLLs já
    // compilados, a duckdb.dll, a biblioteca de tipos e o SageShortcuts. Instala em
    // Arquivos de Programas\Sage e registra para todos os usuários (HKLM), por isso pede
    // administrador. Códigos de saída: 0 ok, 1 erro, 2 cancelado (Excel aberto ou o usuário desistiu).
    static class Program
    {
        [DllImport("user32.dll")] static extern bool SetProcessDPIAware();

        [STAThread]
        static int Main(string[] args)
        {
            HashSet<string> flags = new HashSet<string>(args.Select(a => a.TrimStart('/', '-').ToLowerInvariant()));
            bool quiet = flags.Contains("quiet") || flags.Contains("s");
            SetProcessDPIAware();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Log.Start(flags.Contains("uninstall") ? "uninstall" : "install");
            try
            {
                if (flags.Contains("uninstall")) return Uninstaller.Run(quiet, flags.Contains("confirmed"), flags.Contains("temp"));
                int dotnet = DotNet.Ensure(quiet);
                if (dotnet != 0) return dotnet;
                if (quiet) return Installer.RunQuiet();
                Application.Run(new SetupForm());
                return SetupForm.ExitCode;
            }
            catch (Exception ex)
            {
                Log.Write("ERRO: " + ex);
                if (!quiet) MessageBox.Show(Texts.Failed + "\n\n" + ex.Message, Texts.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }
    }

    // ----------------------------------------------------------------------
    // Textos (português ou inglês, pelo idioma do Windows)
    // ----------------------------------------------------------------------

    static class Texts
    {
        static readonly bool Pt = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "pt";
        static string T(string pt, string en) { return Pt ? pt : en; }

        public static string Title { get { return "Sage Framework " + BuildInfo.Version; } }
        public static string Description { get { return T(
            "Ferramentas para o editor do VBA no Excel: temas de cores, números de linha, abas, vários cursores, " +
            "atalhos de teclado e tipos no estilo do Python (StringS, ListS, DictionaryS, DataFrame, DateTimeS, Json, Requests e SqlEngine).",
            "Tools for the Excel VBA editor: color themes, line numbers, tabs, multiple cursors, keyboard shortcuts " +
            "and Python-style types (StringS, ListS, DictionaryS, DataFrame, DateTimeS, Json, Requests and SqlEngine)."); } }
        public static string Folder { get { return T("Pasta de instalação:", "Install folder:"); } }
        public static string Requirements { get { return T("Requer o Excel de 64 bits e o .NET Framework 4.8 (já vem no Windows 10 e 11).",
            "Requires 64-bit Excel and .NET Framework 4.8 (included in Windows 10 and 11)."); } }
        public static string Installed(string v) { return T("Versão instalada: " + v, "Installed version: " + v); }
        public static string Install { get { return T("Instalar", "Install"); } }
        public static string Update { get { return T("Atualizar", "Update"); } }
        public static string Cancel { get { return T("Cancelar", "Cancel"); } }
        public static string Close { get { return T("Concluir", "Finish"); } }
        public static string Done { get { return T("Instalado. Abra o Excel e o editor do VBA (Alt+F11): o menu Sage já aparece.",
            "Installed. Open Excel and the VBA editor (Alt+F11): the Sage menu is there."); } }
        public static string CloseExcel { get { return T("Feche o Excel antes de continuar (salve o seu trabalho).",
            "Close Excel before continuing (save your work)."); } }
        public static string Failed { get { return T("A instalação falhou.", "Setup failed."); } }
        public static string ConfirmUninstall { get { return T("Remover o Sage Framework e todos os seus componentes?\n\n" +
            "As configurações do Sage (tema, atalhos de teclado) também serão apagadas.",
            "Remove Sage Framework and all of its components?\n\nSage settings (theme, keyboard shortcuts) will also be deleted."); } }
        public static string Uninstalled { get { return T("O Sage Framework foi removido.", "Sage Framework was removed."); } }
        public static string UninstallFailed { get { return T("A desinstalação falhou.", "Uninstall failed."); } }

        // .NET Framework
        public static string DotNetMissing(string found) { return T(
            "O Sage Framework precisa do .NET Framework 4.8 (encontrado: " + found + ").\n\n" +
            "Baixar e instalar agora? O instalador é da Microsoft (cerca de 70 MB baixados durante a instalação).",
            "Sage Framework requires .NET Framework 4.8 (found: " + found + ").\n\n" +
            "Download and install it now? The installer is from Microsoft (about 70 MB downloaded during setup)."); }
        public static string DotNetDownloading { get { return T("Baixando o instalador do .NET Framework 4.8...", "Downloading the .NET Framework 4.8 installer..."); } }
        public static string DotNetBadSignature { get { return T("O arquivo baixado não tem a assinatura digital da Microsoft e não foi executado.",
            "The downloaded file is not signed by Microsoft and was not run."); } }
        public static string DotNetFailed(int code) { return T("A instalação do .NET Framework 4.8 não terminou (código " + code + ").",
            "The .NET Framework 4.8 installation did not finish (code " + code + ")."); }
        public static string DotNetRestart { get { return T("O .NET Framework 4.8 foi instalado. Reinicie o computador e rode este instalador de novo.",
            "The .NET Framework 4.8 was installed. Restart the computer and run this setup again."); } }
        public static string DotNetManual { get { return T("Instale o .NET Framework 4.8 pelo site da Microsoft e rode este instalador de novo:\n",
            "Install .NET Framework 4.8 from Microsoft's website and run this setup again:\n"); } }

        // Etapas
        public static string StepPrepare { get { return T("Preparando...", "Preparing..."); } }
        public static string StepFiles { get { return T("Copiando arquivos...", "Copying files..."); } }
        public static string StepRegister { get { return T("Registrando os componentes...", "Registering components..."); } }
        public static string StepShortcuts { get { return T("Criando os atalhos...", "Creating shortcuts..."); } }
        public static string StepStart { get { return T("Iniciando o SageShortcuts...", "Starting SageShortcuts..."); } }

        // Menu Iniciar
        public static string MenuFolder { get { return "Sage Framework"; } }
        public static string LinkUninstall { get { return T("Desinstalar o Sage Framework", "Uninstall Sage Framework"); } }
        public static string LinkShortcuts { get { return T("SageShortcuts (atalhos de teclado do VBA)", "SageShortcuts (VBA keyboard shortcuts)"); } }
    }

    // ----------------------------------------------------------------------
    // Registro de atividade (%TEMP%\SageFramework-Setup.log), para diagnóstico
    // ----------------------------------------------------------------------

    static class Log
    {
        static string path;

        public static void Start(string mode)
        {
            path = Path.Combine(Path.GetTempPath(), "SageFramework-Setup.log");
            Write("---- " + mode + " " + BuildInfo.Version + " " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
        }

        public static void Write(string line)
        {
            try { File.AppendAllText(path, line + Environment.NewLine, Encoding.UTF8); }
            catch (Exception) { }
        }
    }

    // ----------------------------------------------------------------------
    // .NET Framework 4.8: o Sage precisa dele (TLS 1.3 no Requests, entre outros). O Windows 10
    // e o 11 já o trazem; em versões antigas, oferece baixar e instalar o da Microsoft.
    // (Sem nenhum .NET 4, este instalador nem abre: o próprio Windows pede para instalar.)
    // ----------------------------------------------------------------------

    static class DotNet
    {
        const int Release48 = 528040;   // valor "Release" do .NET Framework 4.8
        public const string Url = "https://go.microsoft.com/fwlink/?linkid=2085155";   // ndp48-web.exe
        const int ErrorCancelled = 1602, RestartRequired = 3010, RestartStarted = 1641;

        public static int Release()
        {
            using (RegistryKey k = Common.Machine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
            {
                object v = k == null ? null : k.GetValue("Release");
                return v is int ? (int)v : 0;
            }
        }

        static string Describe(int release)
        {
            if (release >= 533320) return "4.8.1";
            if (release >= Release48) return "4.8";
            if (release >= 461808) return "4.7.2";
            if (release >= 461308) return "4.7.1";
            if (release >= 460798) return "4.7";
            if (release >= 394802) return "4.6.2";
            if (release >= 394254) return "4.6.1";
            if (release >= 393295) return "4.6";
            if (release >= 379893) return "4.5.2";
            if (release >= 378675) return "4.5.1";
            if (release >= 378389) return "4.5";
            return "4.0";
        }

        // 0: pode seguir; 1: falhou; 2: o usuário não quis
        public static int Ensure(bool quiet)
        {
            int release = Release();
            Log.Write(".NET Framework: Release " + release + " (" + Describe(release) + ")");
            if (release >= Release48) return 0;
            if (!quiet && MessageBox.Show(Texts.DotNetMissing(Describe(release)), Texts.Title, MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning) != DialogResult.Yes)
                return 2;

            string file = Path.Combine(Path.GetTempPath(), "ndp48-web.exe");
            try
            {
                using (new WaitCursor()) Download(file);
                if (!IsSignedByMicrosoft(file))
                {
                    Log.Write("Assinatura inválida: " + file);
                    Common.DeleteFile(file);
                    if (!quiet) MessageBox.Show(Texts.DotNetBadSignature, Texts.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1;
                }
                // /passive: só a barra de progresso; /q: nada (modo silencioso)
                Process p = Process.Start(new ProcessStartInfo(file, (quiet ? "/q" : "/passive") + " /norestart") { UseShellExecute = false });
                p.WaitForExit();
                int code = p.ExitCode;
                Log.Write(".NET Framework 4.8: código " + code);
                Common.DeleteFile(file);
                if (code == RestartRequired || code == RestartStarted)
                {
                    if (!quiet) MessageBox.Show(Texts.DotNetRestart, Texts.Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return 2;
                }
                if (code == ErrorCancelled) return 2;
                if (code != 0 || Release() < Release48)
                {
                    if (!quiet) MessageBox.Show(Texts.DotNetFailed(code) + "\n\n" + Texts.DotNetManual + Url, Texts.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1;
                }
                return 0;
            }
            catch (Exception ex)
            {
                Log.Write(".NET Framework: " + ex);
                if (!quiet) MessageBox.Show(ex.Message + "\n\n" + Texts.DotNetManual + Url, Texts.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        static void Download(string file)
        {
            Log.Write(Texts.DotNetDownloading);
            // TLS 1.2 (3072): versões antigas do .NET usariam o TLS 1.0, recusado pela Microsoft
            ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072;
            using (WebClient web = new WebClient()) web.DownloadFile(Url, file);
        }

        // Assinatura Authenticode válida (WinVerifyTrust) e emitida para a Microsoft
        internal static bool IsSignedByMicrosoft(string file)
        {
            if (!WinTrust.Verify(file)) return false;
            try
            {
                string subject = new System.Security.Cryptography.X509Certificates.X509Certificate(
                    System.Security.Cryptography.X509Certificates.X509Certificate.CreateFromSignedFile(file)).Subject;
                Log.Write("Assinado por: " + subject);
                return subject.Contains("O=Microsoft Corporation");
            }
            catch (Exception) { return false; }
        }

        sealed class WaitCursor : IDisposable
        {
            public WaitCursor() { Cursor.Current = Cursors.WaitCursor; }
            public void Dispose() { Cursor.Current = Cursors.Default; }
        }
    }

    static class WinTrust
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct FileInfo
        {
            public uint Size;
            public string Path;
            public IntPtr Handle, KnownSubject;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct Data
        {
            public uint Size;
            public IntPtr PolicyCallbackData, SipClientData;
            public uint UiChoice, RevocationChecks, UnionChoice;
            public IntPtr File;
            public uint StateAction;
            public IntPtr StateData, UrlReference;
            public uint ProvFlags, UiContext;
            public IntPtr SignatureSettings;
        }

        [DllImport("wintrust.dll", CharSet = CharSet.Unicode)]
        static extern int WinVerifyTrust(IntPtr hwnd, ref Guid action, ref Data data);

        static readonly Guid GenericVerifyV2 = new Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");

        public static bool Verify(string path)
        {
            FileInfo info = new FileInfo { Size = (uint)Marshal.SizeOf(typeof(FileInfo)), Path = path };
            IntPtr pInfo = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(FileInfo)));
            try
            {
                Marshal.StructureToPtr(info, pInfo, false);
                Data data = new Data
                {
                    Size = (uint)Marshal.SizeOf(typeof(Data)),
                    UiChoice = 2,          // WTD_UI_NONE
                    RevocationChecks = 0,  // WTD_REVOKE_NONE
                    UnionChoice = 1,       // WTD_CHOICE_FILE
                    File = pInfo,
                };
                Guid action = GenericVerifyV2;
                int result = WinVerifyTrust(IntPtr.Zero, ref action, ref data);
                Log.Write("WinVerifyTrust " + path + ": 0x" + result.ToString("X8"));
                return result == 0;
            }
            finally
            {
                Marshal.DestroyStructure(pInfo, typeof(FileInfo));
                Marshal.FreeHGlobal(pInfo);
            }
        }
    }

    // ----------------------------------------------------------------------
    // Conteúdo embutido
    // ----------------------------------------------------------------------

    sealed class ComClass
    {
        public string Dll, Class, Clsid, ProgId;
    }

    static class Payload
    {
        public static ZipArchive Open()
        {
            Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("payload.zip");
            if (s == null) throw new InvalidOperationException("payload.zip não está embutido no instalador.");
            return new ZipArchive(s, ZipArchiveMode.Read);
        }

        // Linhas "dll|classe|CLSID|ProgId" dos com.txt de cada componente; dll relativo à pasta do com.txt
        public static List<ComClass> Classes()
        {
            List<ComClass> result = new List<ComClass>();
            using (ZipArchive zip = Open())
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    if (!entry.FullName.EndsWith("com.txt", StringComparison.OrdinalIgnoreCase)) continue;
                    string folder = Path.GetDirectoryName(entry.FullName.Replace('/', '\\'));
                    using (StreamReader r = new StreamReader(entry.Open(), Encoding.UTF8))
                    {
                        string line;
                        while ((line = r.ReadLine()) != null)
                        {
                            string[] p = line.Trim().Split('|');
                            if (p.Length != 4) continue;
                            result.Add(new ComClass { Dll = Path.Combine(folder, p[0]), Class = p[1], Clsid = p[2], ProgId = p[3] });
                        }
                    }
                }
            return result;
        }

        public static long Size()
        {
            using (ZipArchive zip = Open()) return zip.Entries.Sum(e => e.Length);
        }
    }

    // ----------------------------------------------------------------------
    // Caminhos, registro e comuns a instalar e desinstalar
    // ----------------------------------------------------------------------

    static class Paths
    {
        public static string InstallDir { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Sage"); } }
        public static string Uninstaller { get { return Path.Combine(InstallDir, "Uninstall.exe"); } }
        public static string MenuDir { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms), Texts.MenuFolder); } }
        public static string StartupLink { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartup), "SageShortcuts.lnk"); } }
        public static string ShortcutsScript { get { return Path.Combine(InstallDir, @"Shortcuts\SageShortcuts.ps1"); } }
        public static string PowerShell { get { return Path.Combine(Environment.SystemDirectory, @"WindowsPowerShell\v1.0\powershell.exe"); } }
        public static string Conhost { get { return Path.Combine(Environment.SystemDirectory, "conhost.exe"); } }
        public const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\SageFramework";
        public const string AddinKey = @"Software\Microsoft\VBA\VBE\6.0\Addins64\";
        public const string EditorProgId = "Sage.Editor";
    }

    static class Common
    {
        public static RegistryKey Machine { get { return RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64); } }
        public static RegistryKey User { get { return RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64); } }

        public static string InstalledVersion()
        {
            using (RegistryKey k = Machine.OpenSubKey(Paths.UninstallKey))
                return k == null ? null : k.GetValue("DisplayVersion") as string;
        }

        public static bool ExcelRunning() { return Process.GetProcessesByName("EXCEL").Length > 0; }

        // Pede para fechar o Excel até ele fechar ou o usuário desistir
        public static bool WaitForExcel(bool quiet)
        {
            while (ExcelRunning())
            {
                if (quiet) { Log.Write("Excel aberto"); return false; }
                if (MessageBox.Show(Texts.CloseExcel, Texts.Title, MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                    return false;
            }
            return true;
        }

        // Encerra o SageShortcuts (o mesmo sinal do "SageShortcuts.ps1 --stop")
        public static void StopShortcuts()
        {
            foreach (string name in new[] { "SageShortcuts", "VBEShortcuts" })
            {
                try
                {
                    using (EventWaitHandle ev = EventWaitHandle.OpenExisting(@"Local\" + name + ".Stop"))
                    {
                        ev.Set();
                        Log.Write("Sinal de parada: " + name);
                    }
                }
                catch (WaitHandleCannotBeOpenedException) { }
                catch (Exception ex) { Log.Write("Parar " + name + ": " + ex.Message); }
            }
            Thread.Sleep(1000);
        }

        public static void DeleteKey(RegistryKey root, string path)
        {
            try { root.DeleteSubKeyTree(path, false); }
            catch (Exception ex) { Log.Write("Chave " + path + ": " + ex.Message); }
        }

        public static void DeleteFile(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (Exception ex) { Log.Write("Arquivo " + path + ": " + ex.Message); }
        }

        public static void DeleteDir(string path)
        {
            for (int attempt = 0; attempt < 5; attempt++)
            {
                try
                {
                    if (Directory.Exists(path)) Directory.Delete(path, true);
                    return;
                }
                catch (Exception ex)
                {
                    Log.Write("Pasta " + path + ": " + ex.Message);
                    Thread.Sleep(500);
                }
            }
        }

        // Instalações por usuário feitas pelos install.ps1 do repositório (HKCU): ficariam na
        // frente das do instalador (o COM procura primeiro no HKCU)
        public static void RemoveUserInstall(List<ComClass> classes)
        {
            using (RegistryKey user = User)
            {
                foreach (ComClass c in classes)
                {
                    DeleteKey(user, @"Software\Classes\CLSID\" + c.Clsid);
                    DeleteKey(user, @"Software\Classes\" + c.ProgId);
                }
                foreach (string legacy in new[] { "Sage.VBE", "SageTypes.StringS" }) DeleteKey(user, @"Software\Classes\" + legacy);
                foreach (string addin in new[] { Paths.EditorProgId, "Sage.VBE" }) DeleteKey(user, Paths.AddinKey + addin);
            }
            TypeLibs.UnregisterUser(Path.Combine(Paths.InstallDir, @"Types\SageTypes.tlb"));
            string startup = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            DeleteFile(Path.Combine(startup, "SageShortcuts.lnk"));
            DeleteFile(Path.Combine(startup, "VBEShortcuts.lnk"));
            string local = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Sage");
            foreach (string sub in new[] { "Editor", "Types", "VBE" }) DeleteDir(Path.Combine(local, sub));
        }

        // Atalho .lnk (WScript.Shell)
        public static void Shortcut(string path, string target, string arguments, string workingDir, string icon, string description, int windowStyle)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            dynamic shell = Activator.CreateInstance(Type.GetTypeFromProgID("WScript.Shell"));
            dynamic link = shell.CreateShortcut(path);
            link.TargetPath = target;
            if (arguments != null) link.Arguments = arguments;
            if (workingDir != null) link.WorkingDirectory = workingDir;
            if (icon != null) link.IconLocation = icon;
            if (description != null) link.Description = description;
            link.WindowStyle = windowStyle;
            link.Save();
            Marshal.FinalReleaseComObject(link);
            Marshal.FinalReleaseComObject(shell);
        }

        public static string ShortcutsArguments
        {
            // Pelo conhost sem janela: com o Windows Terminal como terminal padrão (Windows 11),
            // o powershell.exe abriria numa aba dele, e o -WindowStyle Hidden não a esconde
            get { return "--headless \"" + Paths.PowerShell + "\" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + Paths.ShortcutsScript + "\""; }
        }
    }

    // ----------------------------------------------------------------------
    // Registro de cada usuário do Windows (HKEY_CURRENT_USER de cada um): os perfis com sessão
    // aberta (já carregados em HKEY_USERS), os outros (NTUSER.DAT carregado e descarregado) e o
    // perfil padrão, copiado para os usuários criados depois
    // ----------------------------------------------------------------------

    static class UserHives
    {
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)] static extern int RegLoadKey(IntPtr hkey, string subKey, string file);
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)] static extern int RegUnLoadKey(IntPtr hkey, string subKey);
        [DllImport("advapi32.dll", SetLastError = true)] static extern bool OpenProcessToken(IntPtr process, int access, out IntPtr token);
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern bool LookupPrivilegeValue(string system, string name, out long luid);
        [DllImport("advapi32.dll", SetLastError = true)] static extern bool AdjustTokenPrivileges(IntPtr token, bool disableAll, ref TokenPrivilege state, int length, IntPtr previous, IntPtr returnLength);
        [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        struct TokenPrivilege { public int Count; public long Luid; public int Attributes; }

        static readonly IntPtr HKEY_USERS = new IntPtr(unchecked((int)0x80000003));
        const string TempHive = "SageSetupUser";

        public static void ForEach(Action<RegistryKey> action)
        {
            // Carregar o NTUSER.DAT de outro usuário exige estes privilégios (o administrador tem)
            EnablePrivilege("SeRestorePrivilege");
            EnablePrivilege("SeBackupPrivilege");
            using (RegistryKey users = RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Registry64))
            using (RegistryKey list = Common.Machine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"))
            {
                if (list == null) return;
                foreach (string sid in list.GetSubKeyNames())
                {
                    if (!sid.StartsWith("S-1-5-21-")) continue; // contas de pessoas (não as do sistema)
                    try
                    {
                        using (RegistryKey loaded = users.OpenSubKey(sid, true))
                            if (loaded != null) { action(loaded); Log.Write("Usuário " + sid); continue; }
                        using (RegistryKey profile = list.OpenSubKey(sid))
                        {
                            string dir = profile == null ? null : profile.GetValue("ProfileImagePath") as string;
                            if (dir != null) Offline(users, Path.Combine(Environment.ExpandEnvironmentVariables(dir), "NTUSER.DAT"), action);
                        }
                    }
                    catch (Exception ex) { Log.Write("Usuário " + sid + ": " + ex.Message); }
                }
                string defaultDir = list.GetValue("Default") as string;
                if (defaultDir != null)
                {
                    try { Offline(users, Path.Combine(Environment.ExpandEnvironmentVariables(defaultDir), "NTUSER.DAT"), action); }
                    catch (Exception ex) { Log.Write("Perfil padrão: " + ex.Message); }
                }
            }
        }

        static void Offline(RegistryKey users, string hive, Action<RegistryKey> action)
        {
            if (!File.Exists(hive)) return;
            int result = RegLoadKey(HKEY_USERS, TempHive, hive);
            if (result != 0) { Log.Write("RegLoadKey " + hive + ": " + result); return; }
            try
            {
                using (RegistryKey root = users.OpenSubKey(TempHive, true)) action(root);
                Log.Write("Perfil " + hive);
            }
            finally
            {
                // As chaves abertas precisam estar fechadas para descarregar
                GC.Collect();
                GC.WaitForPendingFinalizers();
                result = RegUnLoadKey(HKEY_USERS, TempHive);
                if (result != 0) Log.Write("RegUnLoadKey " + hive + ": " + result);
            }
        }

        static void EnablePrivilege(string name)
        {
            IntPtr token;
            if (!OpenProcessToken(Process.GetCurrentProcess().Handle, 0x28 /* ADJUST_PRIVILEGES | QUERY */, out token)) return;
            try
            {
                TokenPrivilege tp = new TokenPrivilege { Count = 1, Attributes = 2 /* SE_PRIVILEGE_ENABLED */ };
                if (LookupPrivilegeValue(null, name, out tp.Luid))
                    AdjustTokenPrivileges(token, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
            }
            finally { CloseHandle(token); }
        }
    }

    // ----------------------------------------------------------------------
    // Biblioteca de tipos (oleaut32)
    // ----------------------------------------------------------------------

    static class TypeLibs
    {
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        static extern void LoadTypeLibEx(string file, int regKind, out ITypeLib lib);
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        static extern void RegisterTypeLib(ITypeLib lib, string fullPath, string helpDir);
        [DllImport("oleaut32.dll")]
        static extern int UnRegisterTypeLib(ref Guid libId, short major, short minor, int lcid, SYSKIND syskind);
        [DllImport("oleaut32.dll")]
        static extern int UnRegisterTypeLibForUser(ref Guid libId, short major, short minor, int lcid, SYSKIND syskind);

        public static void Register(string tlb)
        {
            ITypeLib lib;
            LoadTypeLibEx(tlb, 2 /* REGKIND_NONE */, out lib);
            try { RegisterTypeLib(lib, tlb, Path.GetDirectoryName(tlb)); }
            finally { Marshal.FinalReleaseComObject(lib); }
        }

        public static void Unregister(string tlb) { Remove(tlb, false); }
        public static void UnregisterUser(string tlb) { Remove(tlb, true); }

        // O LIBID e a versão vêm do próprio .tlb
        static void Remove(string tlb, bool user)
        {
            if (!File.Exists(tlb)) return;
            try
            {
                ITypeLib lib;
                LoadTypeLibEx(tlb, 2, out lib);
                IntPtr pa;
                lib.GetLibAttr(out pa);
                TYPELIBATTR a = (TYPELIBATTR)Marshal.PtrToStructure(pa, typeof(TYPELIBATTR));
                lib.ReleaseTLibAttr(pa);
                Marshal.FinalReleaseComObject(lib);
                Guid id = a.guid;
                int hr = user
                    ? UnRegisterTypeLibForUser(ref id, a.wMajorVerNum, a.wMinorVerNum, a.lcid, a.syskind)
                    : UnRegisterTypeLib(ref id, a.wMajorVerNum, a.wMinorVerNum, a.lcid, a.syskind);
                Log.Write("TypeLib " + id + (user ? " (usuário)" : "") + ": 0x" + hr.ToString("X8"));
                // A API falha se o arquivo registrado não existe mais (instalação antiga apagada), e
                // o registro do usuário, que tem prioridade, faria o VBA dar "Erro ao carregar DLL"
                if (user)
                    using (RegistryKey root = Common.User)
                        Common.DeleteKey(root, @"Software\Classes\TypeLib\" + id.ToString("B").ToUpperInvariant());
            }
            catch (Exception ex) { Log.Write("TypeLib " + tlb + ": " + ex.Message); }
        }
    }

    // ----------------------------------------------------------------------
    // Instalação
    // ----------------------------------------------------------------------

    static class Installer
    {
        public static int RunQuiet()
        {
            if (!Common.WaitForExcel(true)) return 2;
            Install(delegate(int p, string s) { Log.Write(p + "% " + s); });
            return 0;
        }

        public static void Install(Action<int, string> progress)
        {
            progress(5, Texts.StepPrepare);
            List<ComClass> classes = Payload.Classes();
            Common.StopShortcuts();
            // Versão anterior: tira os registros antes de trocar os arquivos
            TypeLibs.Unregister(Path.Combine(Paths.InstallDir, @"Types\SageTypes.tlb"));

            progress(15, Texts.StepFiles);
            string dir = Paths.InstallDir;
            Common.DeleteDir(dir);
            Directory.CreateDirectory(dir);
            using (ZipArchive zip = Payload.Open())
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    if (entry.FullName.EndsWith("/")) continue;
                    string target = Path.Combine(dir, entry.FullName.Replace('/', '\\'));
                    Directory.CreateDirectory(Path.GetDirectoryName(target));
                    using (Stream from = entry.Open())
                    using (FileStream to = File.Create(target))
                        from.CopyTo(to);
                }
            File.Copy(Application.ExecutablePath, Paths.Uninstaller, true);
            Log.Write("Arquivos em " + dir);
            // Instalação por usuário (install.ps1 do repositório): depois da cópia, porque o LIBID
            // da biblioteca de tipos é lido do .tlb que acabou de ser instalado
            Common.RemoveUserInstall(classes);

            progress(55, Texts.StepRegister);
            using (RegistryKey machine = Common.Machine)
            {
                foreach (ComClass c in classes) RegisterClass(machine, c);
                // Versões anteriores do instalador registravam o suplemento aqui (o VBE ignora)
                Common.DeleteKey(machine, Paths.AddinKey + Paths.EditorProgId);
            }
            // Suplemento do VBE (3 = carregar ao iniciar): o VBE só lê os suplementos do registro
            // do usuário, então vai no de cada um (inclusive no perfil padrão, para os usuários novos)
            UserHives.ForEach(delegate(RegistryKey user)
            {
                using (RegistryKey k = user.CreateSubKey(Paths.AddinKey + Paths.EditorProgId))
                {
                    k.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    k.SetValue("FriendlyName", "Sage");
                    k.SetValue("Description", "Menu Sage e temas de cores para o editor do VBA");
                }
            });
            foreach (string tlb in Directory.GetFiles(dir, "*.tlb", SearchOption.AllDirectories))
            {
                TypeLibs.Register(tlb);
                Log.Write("TypeLib registrada: " + tlb);
            }

            progress(75, Texts.StepShortcuts);
            string menu = Paths.MenuDir;
            Common.DeleteDir(menu);
            // No menu Iniciar, só o desinstalador
            Common.Shortcut(Path.Combine(menu, Texts.LinkUninstall + ".lnk"), Paths.Uninstaller, "/uninstall", dir, Paths.Uninstaller + ",0", Texts.LinkUninstall, 1);
            // Inicia com o Windows, para todos os usuários
            Common.Shortcut(Paths.StartupLink, Paths.Conhost, Common.ShortcutsArguments,
                Path.GetDirectoryName(Paths.ShortcutsScript), Paths.Uninstaller + ",0", Texts.LinkShortcuts, 7);
            WriteUninstallEntry(dir);

            progress(90, Texts.StepStart);
            // Pelo Explorer, para rodar como o usuário (sem os privilégios de administrador do instalador)
            try { Process.Start(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "explorer.exe"), "\"" + Paths.StartupLink + "\""); }
            catch (Exception ex) { Log.Write("Iniciar SageShortcuts: " + ex.Message); }

            progress(100, Texts.Done);
            Log.Write("Instalado");
        }

        static void RegisterClass(RegistryKey machine, ComClass c)
        {
            string dll = Path.Combine(Paths.InstallDir, c.Dll);
            string assembly = AssemblyName.GetAssemblyName(dll).FullName;
            using (RegistryKey key = machine.CreateSubKey(@"Software\Classes\CLSID\" + c.Clsid))
            {
                key.SetValue("", c.Class);
                using (RegistryKey p = key.CreateSubKey("ProgId")) p.SetValue("", c.ProgId);
                using (RegistryKey inproc = key.CreateSubKey("InprocServer32"))
                {
                    inproc.SetValue("", "mscoree.dll");
                    inproc.SetValue("ThreadingModel", "Both");
                    inproc.SetValue("Class", c.Class);
                    inproc.SetValue("Assembly", assembly);
                    inproc.SetValue("RuntimeVersion", "v4.0.30319");
                    // Sem escapar os espaços: o mscoree não decodifica "%20" e procuraria
                    // "Program%20Files" (Uri.AbsoluteUri escaparia)
                    inproc.SetValue("CodeBase", "file:///" + dll.Replace('\\', '/'));
                }
                // Categoria "componentes .NET"
                key.CreateSubKey(@"Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}").Close();
            }
            using (RegistryKey key = machine.CreateSubKey(@"Software\Classes\" + c.ProgId))
            {
                key.SetValue("", c.Class);
                using (RegistryKey id = key.CreateSubKey("CLSID")) id.SetValue("", c.Clsid);
            }
            Log.Write("Classe " + c.ProgId + " " + c.Clsid);
        }

        // Aparece em Configurações > Aplicativos (e no Painel de Controle)
        static void WriteUninstallEntry(string dir)
        {
            long kb = Directory.GetFiles(dir, "*", SearchOption.AllDirectories).Sum(f => new FileInfo(f).Length) / 1024;
            using (RegistryKey machine = Common.Machine)
            using (RegistryKey k = machine.CreateSubKey(Paths.UninstallKey))
            {
                k.SetValue("DisplayName", "Sage Framework");
                k.SetValue("DisplayVersion", BuildInfo.Version);
                k.SetValue("Publisher", "Sage Framework");
                k.SetValue("DisplayIcon", Paths.Uninstaller + ",0");
                k.SetValue("InstallLocation", dir);
                k.SetValue("UninstallString", "\"" + Paths.Uninstaller + "\" /uninstall");
                k.SetValue("QuietUninstallString", "\"" + Paths.Uninstaller + "\" /uninstall /quiet");
                k.SetValue("URLInfoAbout", "https://github.com/sage-framework-2023/Sage-Framework");
                k.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
                k.SetValue("EstimatedSize", (int)Math.Min(int.MaxValue, kb), RegistryValueKind.DWord);
                k.SetValue("NoModify", 1, RegistryValueKind.DWord);
                k.SetValue("NoRepair", 1, RegistryValueKind.DWord);
            }
        }
    }

    // ----------------------------------------------------------------------
    // Desinstalação
    // ----------------------------------------------------------------------

    static class Uninstaller
    {
        public static int Run(bool quiet, bool confirmed, bool fromTemp)
        {
            if (!quiet && !confirmed &&
                MessageBox.Show(Texts.ConfirmUninstall, Texts.Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return 2;

            // Rodando de dentro da pasta instalada: continua de uma cópia na pasta temporária,
            // para poder apagar a pasta (e o próprio Uninstall.exe)
            string self = Application.ExecutablePath;
            if (!fromTemp && self.StartsWith(Paths.InstallDir + "\\", StringComparison.OrdinalIgnoreCase))
            {
                string copy = Path.Combine(Path.GetTempPath(), "SageFramework-Uninstall.exe");
                File.Copy(self, copy, true);
                Process p = Process.Start(new ProcessStartInfo(copy, "/uninstall /confirmed /temp" + (quiet ? " /quiet" : "")) { UseShellExecute = false });
                if (!quiet) return 0;
                p.WaitForExit();
                return p.ExitCode;
            }

            if (!Common.WaitForExcel(quiet)) return 2;
            try
            {
                using (new CursorScope())
                    Uninstall();
            }
            catch (Exception ex)
            {
                Log.Write("ERRO: " + ex);
                if (!quiet) MessageBox.Show(Texts.UninstallFailed + "\n\n" + ex.Message, Texts.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
            finally
            {
                if (fromTemp) DeleteSelfLater();
            }
            if (!quiet) MessageBox.Show(Texts.Uninstalled, Texts.Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return 0;
        }

        static void Uninstall()
        {
            List<ComClass> classes = Payload.Classes();
            Common.StopShortcuts();
            string dir = Paths.InstallDir;

            TypeLibs.Unregister(Path.Combine(dir, @"Types\SageTypes.tlb"));
            using (RegistryKey machine = Common.Machine)
            {
                foreach (ComClass c in classes)
                {
                    Common.DeleteKey(machine, @"Software\Classes\CLSID\" + c.Clsid);
                    Common.DeleteKey(machine, @"Software\Classes\" + c.ProgId);
                }
                Common.DeleteKey(machine, Paths.AddinKey + Paths.EditorProgId);
                Common.DeleteKey(machine, Paths.UninstallKey);
            }
            UserHives.ForEach(delegate(RegistryKey user) { Common.DeleteKey(user, Paths.AddinKey + Paths.EditorProgId); });
            Common.RemoveUserInstall(classes);

            Common.DeleteFile(Paths.StartupLink);
            Common.DeleteDir(Paths.MenuDir);
            Common.DeleteDir(dir);

            // Dados do usuário que desinstala: configurações, log, atalhos e o banco temporário do DataFrame
            Common.DeleteDir(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Sage"));
            Common.DeleteDir(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Sage"));
            Common.DeleteDir(Path.Combine(Path.GetTempPath(), "SageDuckDB"));
            Log.Write("Desinstalado");
        }

        // A cópia temporária se apaga depois de terminar
        static void DeleteSelfLater()
        {
            try
            {
                string cmd = "/c ping 127.0.0.1 -n 3 > nul & del /f /q \"" + Application.ExecutablePath + "\"";
                Process.Start(new ProcessStartInfo("cmd.exe", cmd) { WindowStyle = ProcessWindowStyle.Hidden, CreateNoWindow = true, UseShellExecute = false });
            }
            catch (Exception ex) { Log.Write("Apagar cópia: " + ex.Message); }
        }

        sealed class CursorScope : IDisposable
        {
            public CursorScope() { Cursor.Current = Cursors.WaitCursor; }
            public void Dispose() { Cursor.Current = Cursors.Default; }
        }
    }

    // ----------------------------------------------------------------------
    // Janela do instalador
    // ----------------------------------------------------------------------

    sealed class SetupForm : Form
    {
        public static int ExitCode = 2;

        readonly Button install, cancel;
        readonly ProgressBar bar;
        readonly Label status;
        bool busy, finished;

        public SetupForm()
        {
            Text = Texts.Title;
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96f, 96f);
            ClientSize = new Size(560, 330);
            BackColor = Color.White;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch (Exception) { }

            // Faixa do título, nas cores do tema Sage
            Panel header = new Panel { Dock = DockStyle.Top, Height = 78, BackColor = Color.FromArgb(0x2B, 0x33, 0x2E) };
            PictureBox logo = new PictureBox { Location = new Point(20, 15), Size = new Size(48, 48), SizeMode = PictureBoxSizeMode.Zoom };
            try { logo.Image = new Icon(Icon, 48, 48).ToBitmap(); } catch (Exception) { }
            Label title = new Label { Text = "Sage Framework", Location = new Point(80, 14), AutoSize = true, ForeColor = Color.FromArgb(0xC8, 0xD8, 0xCC),
                Font = new Font("Segoe UI Semibold", 16f) };
            Label version = new Label { Text = BuildInfo.Version, Location = new Point(83, 47), AutoSize = true, ForeColor = Color.FromArgb(0x8F, 0xA8, 0x96) };
            header.Controls.Add(logo);
            header.Controls.Add(title);
            header.Controls.Add(version);

            Label description = new Label { Text = Texts.Description, Location = new Point(20, 94), Size = new Size(520, 56) };
            Label folderLabel = new Label { Text = Texts.Folder, Location = new Point(20, 156), AutoSize = true, ForeColor = Color.DimGray };
            TextBox folder = new TextBox { Text = Paths.InstallDir, Location = new Point(20, 176), Width = 520, ReadOnly = true };
            string installed = Common.InstalledVersion();
            Label info = new Label { Text = (installed != null ? Texts.Installed(installed) + "    " : "") + Texts.Requirements,
                Location = new Point(20, 208), Size = new Size(520, 36), ForeColor = Color.DimGray };

            bar = new ProgressBar { Location = new Point(20, 250), Width = 520, Height = 16, Visible = false };
            status = new Label { Location = new Point(20, 270), Size = new Size(520, 20), ForeColor = Color.DimGray };

            install = new Button { Text = installed != null ? Texts.Update : Texts.Install, Location = new Point(354, 292), Size = new Size(90, 28) };
            cancel = new Button { Text = Texts.Cancel, Location = new Point(450, 292), Size = new Size(90, 28) };
            install.Click += delegate { if (finished) Close(); else Start(); };
            cancel.Click += delegate { Close(); };
            AcceptButton = install;
            CancelButton = cancel;

            Controls.AddRange(new Control[] { header, description, folderLabel, folder, info, bar, status, install, cancel });
            FormClosing += (s, e) => { if (busy) e.Cancel = true; };
        }

        void Start()
        {
            if (!Common.WaitForExcel(false)) return;
            busy = true;
            install.Enabled = cancel.Enabled = false;
            bar.Visible = true;
            Thread worker = new Thread(delegate()
            {
                Exception error = null;
                try { Installer.Install(Report); }
                catch (Exception ex) { error = ex; Log.Write("ERRO: " + ex); }
                BeginInvoke((Action)delegate { Finish(error); });
            });
            worker.SetApartmentState(ApartmentState.STA); // WScript.Shell
            worker.Start();
        }

        void Report(int percent, string text)
        {
            BeginInvoke((Action)delegate { bar.Value = percent; status.Text = text; });
        }

        void Finish(Exception error)
        {
            busy = false;
            if (error != null)
            {
                ExitCode = 1;
                status.Text = Texts.Failed;
                status.ForeColor = Color.Firebrick;
                cancel.Enabled = true;
                MessageBox.Show(this, Texts.Failed + "\n\n" + error.Message, Texts.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            ExitCode = 0;
            finished = true;
            install.Text = Texts.Close;
            install.Enabled = true;
            cancel.Visible = false;
        }
    }
}
