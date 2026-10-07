using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace SageEditor
{
    // Programas de console (ping, git, python, .bat) chamados na aba Terminal.
    //
    // O Excel do Office não pode ter console (AllocConsole e AttachConsole falham). Sem console,
    // o PowerShell iniciaria cada programa num console novo, visível, separado do VBE, e leria a
    // saída sem saber a codificação (acentos trocados). Por isso o terminal inicia esses programas
    // ele mesmo (PostCommandLookupAction, em PsConsole.InitScript): com um console sem janela
    // (CreateNoWindow), lendo a saída em bytes e decodificando cada linha. UTF-8 válido é UTF-8
    // (git, python moderno); o resto, a página de código do console do Windows (850 em português:
    // ping, ipconfig, dir). O erro do programa (stderr) vem separado, para aparecer em vermelho.
    [ComVisible(false)]
    public sealed class NativeLauncher
    {
        // Só os programas de console: os de janela (notepad, explorer) seguem o caminho normal
        // do PowerShell, que não espera por eles
        public bool IsConsole(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".bat" || ext == ".cmd" || ext == ".com") return true;
            if (ext != ".exe") return false;
            try
            {
                using (FileStream f = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (BinaryReader r = new BinaryReader(f))
                {
                    f.Position = 0x3C;
                    int pe = r.ReadInt32();
                    f.Position = pe + 4 + 20 + 68; // assinatura, cabeçalho do arquivo, subsistema no opcional
                    return r.ReadUInt16() == 3;    // IMAGE_SUBSYSTEM_WINDOWS_CUI
                }
            }
            catch (Exception) { return false; }
        }

        public NativeProcess Start(string path, object[] args, string directory)
        {
            return new NativeProcess(path, args ?? new object[0], directory);
        }
    }

    [ComVisible(false)]
    public sealed class NativeLine
    {
        public string Text;
        public bool IsError;
    }

    [ComVisible(false)]
    public sealed class NativeProcess
    {
        readonly Process process;
        readonly BlockingCollection<NativeLine> lines = new BlockingCollection<NativeLine>();
        readonly Encoding oem = Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage);
        int readers = 2;

        internal NativeProcess(string path, object[] args, string directory)
        {
            ProcessStartInfo info = new ProcessStartInfo();
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".bat" || ext == ".cmd")
            {
                info.FileName = Path.Combine(Environment.SystemDirectory, "cmd.exe");
                info.Arguments = "/d /c \"" + Quote(path) + (args.Length > 0 ? " " + CommandLine(args) : "") + "\"";
            }
            else
            {
                info.FileName = path;
                info.Arguments = CommandLine(args);
            }
            info.UseShellExecute = false;
            info.CreateNoWindow = true; // console sem janela: nenhuma janela abre
            info.RedirectStandardInput = true;
            info.RedirectStandardOutput = true;
            info.RedirectStandardError = true;
            if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory)) info.WorkingDirectory = directory;
            process = Process.Start(info);
            Read(process.StandardOutput.BaseStream, false);
            Read(process.StandardError.BaseStream, true);
        }

        // Linhas do PowerShell para a entrada do programa ("texto" | programa)
        public void WriteInput(string text)
        {
            byte[] bytes = oem.GetBytes((text ?? "") + "\r\n");
            process.StandardInput.BaseStream.Write(bytes, 0, bytes.Length);
        }

        public void CloseInput()
        {
            try { process.StandardInput.Close(); } catch (Exception) { }
        }

        // A próxima linha; null se nada chegou no tempo dado (o PowerShell confere o Ctrl+C
        // entre uma chamada e outra) ou se acabou (Done)
        public NativeLine Next(int milliseconds)
        {
            NativeLine line;
            return lines.TryTake(out line, milliseconds) ? line : null;
        }

        public bool Done { get { return lines.IsCompleted; } }

        public int ExitCode
        {
            get
            {
                try { process.WaitForExit(2000); return process.HasExited ? process.ExitCode : -1; }
                catch (Exception) { return -1; }
            }
        }

        // Ctrl+C ou fim: encerra o programa e os que ele abriu
        public void Stop()
        {
            try
            {
                if (process.HasExited) return;
                ProcessStartInfo kill = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "taskkill.exe"),
                    "/PID " + process.Id + " /T /F");
                kill.UseShellExecute = false;
                kill.CreateNoWindow = true;
                using (Process k = Process.Start(kill)) k.WaitForExit(5000);
                if (!process.HasExited) process.Kill();
            }
            catch (Exception) { }
        }

        void Read(Stream stream, bool isError)
        {
            Thread reader = new Thread(delegate()
            {
                MemoryStream current = new MemoryStream();
                byte[] buffer = new byte[4096];
                try
                {
                    int n;
                    while ((n = stream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        for (int i = 0; i < n; i++)
                        {
                            if (buffer[i] == (byte)'\n') { Emit(current, isError); current.SetLength(0); }
                            else current.WriteByte(buffer[i]);
                        }
                    }
                    if (current.Length > 0) Emit(current, isError);
                }
                catch (Exception) { }
                finally
                {
                    if (Interlocked.Decrement(ref readers) == 0) lines.CompleteAdding();
                }
            });
            reader.IsBackground = true;
            reader.Start();
        }

        void Emit(MemoryStream bytes, bool isError)
        {
            byte[] b = bytes.ToArray();
            int length = b.Length;
            if (length > 0 && b[length - 1] == (byte)'\r') length--;
            string text = Decode(b, length);
            try { lines.Add(new NativeLine { Text = text, IsError = isError }); }
            catch (InvalidOperationException) { }
        }

        string Decode(byte[] b, int length)
        {
            bool ascii = true;
            for (int i = 0; i < length; i++) if (b[i] >= 0x80) { ascii = false; break; }
            if (ascii) return Encoding.ASCII.GetString(b, 0, length);
            try { return new UTF8Encoding(false, true).GetString(b, 0, length); } // UTF-8 válido
            catch (DecoderFallbackException) { return oem.GetString(b, 0, length); }
        }

        // Linha de comando do Windows a partir dos argumentos (aspas e barras como o CRT espera)
        static string CommandLine(object[] args)
        {
            List<string> parts = new List<string>();
            foreach (object a in Flatten(args)) parts.Add(Quote(Convert.ToString(a, CultureInfo.InvariantCulture)));
            return string.Join(" ", parts.ToArray());
        }

        static IEnumerable<object> Flatten(IEnumerable<object> args)
        {
            foreach (object a in args)
            {
                object value = a is System.Management.Automation.PSObject ? ((System.Management.Automation.PSObject)a).BaseObject : a;
                System.Collections.IEnumerable list = value as System.Collections.IEnumerable;
                if (list != null && !(value is string))
                {
                    List<object> inner = new List<object>();
                    foreach (object x in list) inner.Add(x);
                    foreach (object x in Flatten(inner)) yield return x;
                }
                else yield return value;
            }
        }

        // Como o Windows PowerShell 5.1 passa os argumentos (é o que cmd /c "..." e os scripts
        // existentes esperam): com espaço e sem aspas em volta, ganha aspas; as aspas de dentro
        // vão como estão; barras no fim, antes da aspa que fecha, são dobradas
        static string Quote(string arg)
        {
            if (arg == null) arg = "";
            if (arg.Length == 0) return "\"\"";
            if (arg.IndexOfAny(new[] { ' ', '\t' }) < 0) return arg;
            if (arg.Length > 1 && arg[0] == '"' && arg[arg.Length - 1] == '"') return arg;
            int slashes = 0;
            for (int i = arg.Length - 1; i >= 0 && arg[i] == '\\'; i--) slashes++;
            return "\"" + arg + new string('\\', slashes) + "\"";
        }
    }
}
