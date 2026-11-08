using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using IOPath = System.IO.Path;

namespace SageTypes
{
    // Path: um caminho de arquivo ou pasta, e tudo o que se faz com ele (no lugar do
    // FileSystemObject da biblioteca Scripting):
    //
    //   Dim pasta As New Sage.Path
    //   pasta = "dados"                                   ' sem pasta: a da pasta de trabalho ativa
    //   Set arquivo = pasta.Join("vendas", "2026.csv")
    //   If arquivo.Exists Then texto = arquivo.ReadText()
    //   For Each f In pasta.Glob("**\*.csv"): Debug.Print f.Name, f.Size: Next
    //   Set area = PathOf(sgDesktop)                      ' global: um caminho numa linha
    //
    // Um caminho sem pasta é completado na hora em que é atribuído, como nos Read do DataFrame:
    // a partir da pasta da pasta de trabalho ativa (Excel.ResolvePath). O texto é lido em
    // UTF-8 (ou na codificação da marca no início do arquivo; sem marca e sem UTF-8 válido, na do
    // Windows) e gravado em UTF-8, com as quebras de linha do Windows.
    //
    // Interface dual: acrescente membros só no fim, com o próximo DispId (ver String).
    [ComVisible(true), Guid("D96DE1A1-BEE0-4A4C-84E5-455AA1040991"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _Path
    {
        [DispId(0)] string Value { get; set; }
        [DispId(1)] string Name { get; }
        [DispId(2)] string Stem { get; }
        [DispId(3)] string Extension { get; }
        [DispId(4)] Path Parent { get; }
        [DispId(5)] bool Exists { get; }
        [DispId(6)] bool IsFile { get; }
        [DispId(7)] bool IsFolder { get; }
        [DispId(8)] double Size { get; }
        [DispId(9)] Datetime Modified { get; }
        [DispId(10)] Datetime Created { get; }
        [DispId(11)] Path Join([Optional] object P1, [Optional] object P2, [Optional] object P3, [Optional] object P4, [Optional] object P5,
            [Optional] object P6, [Optional] object P7, [Optional] object P8, [Optional] object P9, [Optional] object P10);
        [DispId(12)] Path WithName(string Name);
        [DispId(13)] Path WithExtension(string Extension);
        [DispId(14)] String ReadText([Optional] object Encoding);
        [DispId(15)] List ReadLines([Optional] object Encoding);
        [DispId(16)] void WriteText(object Text, [Optional] object Encoding);
        [DispId(17)] void WriteLines(object Lines, [Optional] object Encoding);
        [DispId(18)] void AppendText(object Text, [Optional] object Encoding);
        [DispId(19)] List Files([Optional] object Pattern);
        [DispId(20)] List Folders([Optional] object Pattern);
        [DispId(21)] List Glob(string Pattern);
        [DispId(22)] Path Mkdir();
        [DispId(23)] void Delete([Optional] object Recursive);
        [DispId(24)] Path CopyTo(object Destination, [Optional] object Overwrite);
        [DispId(25)] Path MoveTo(object Destination, [Optional] object Overwrite);
        [DispId(26)] string ToString();
    }

    // Pastas do Windows para o PathOf (e PathOf(sgCurrentFolder): de onde partem os caminhos sem pasta)
    [ComVisible(true), Guid("5FAFF811-3A8B-4E35-BF08-C5DF0F63C8DF")]
    public enum sgFolders
    {
        sgCurrentFolder = 0,
        sgTemp = 1,
        sgHome = 2,
        sgDesktop = 3,
        sgDocuments = 4,
        sgDownloads = 5,
        sgAppData = 6,
        sgLocalAppData = 7,
    }

    [ComVisible(true), Guid("F6B281D0-ED8D-4698-9C96-DFCCF0A7E1B1"), ProgId("Sage.Path")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_Path))]
    public sealed class Path : _Path
    {
        string full = "";

        public Path() { }

        internal static Path From(string path)
        {
            Path p = new Path();
            p.full = path ?? "";
            return p;
        }

        // Texto, String, outro Path ou uma das pastas do Windows (sgFolders)
        internal static Path Of(object source)
        {
            source = Interop.Unwrap(source);
            Path other = source as Path;
            if (other != null) return From(other.full);
            if (source is int || source is short || source is long || source is byte)
                return From(Folder((sgFolders)Convert.ToInt32(source, CultureInfo.InvariantCulture)));
            if (source is sgFolders) return From(Folder((sgFolders)source));
            return From(Full(Interop.Text(source)));
        }

        static string Folder(sgFolders which)
        {
            switch (which)
            {
                case sgFolders.sgCurrentFolder: return IOPath.GetFullPath(Excel.ResolvePath("."));
                case sgFolders.sgTemp: return IOPath.GetTempPath().TrimEnd('\\');
                case sgFolders.sgHome: return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                case sgFolders.sgDesktop: return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                case sgFolders.sgDocuments: return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                case sgFolders.sgDownloads: return KnownFolder(new Guid("374DE290-123F-4565-9164-39C4925E467B"),
                    IOPath.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"));
                case sgFolders.sgAppData: return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                case sgFolders.sgLocalAppData: return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            }
            throw Interop.Error(5, "ValueError: pasta desconhecida: " + (int)which);
        }

        // Caminho completo: sem pasta, a partir da pasta de trabalho ativa; "." e ".." resolvidos
        static string Full(string text)
        {
            string t = (text ?? "").Trim().Trim('"');
            if (t.Length == 0) return "";
            t = t.Replace('/', '\\');
            try { return IOPath.GetFullPath(Excel.ResolvePath(t)); }
            catch (Exception) { throw Interop.Error(52, "ValueError: caminho inválido: '" + text + "'"); }
        }

        string Need()
        {
            if (full.Length == 0) throw Interop.Error(5, "ValueError: o Path está vazio");
            return full;
        }

        // ------------------------------------------------------------------
        // Partes do caminho
        // ------------------------------------------------------------------

        public string Value { get { return full; } set { full = Full(value); } }

        public override string ToString() { return full; }

        public string Name { get { return IOPath.GetFileName(full.TrimEnd('\\')); } }
        public string Stem { get { return IOPath.GetFileNameWithoutExtension(full.TrimEnd('\\')); } }
        public string Extension { get { return IOPath.GetExtension(full.TrimEnd('\\')); } }

        public Path Parent
        {
            get
            {
                string parent = IOPath.GetDirectoryName(Need().TrimEnd('\\'));
                return From(string.IsNullOrEmpty(parent) ? full : parent);
            }
        }

        public bool Exists { get { return full.Length > 0 && (File.Exists(full) || Directory.Exists(full)); } }
        public bool IsFile { get { return full.Length > 0 && File.Exists(full); } }
        public bool IsFolder { get { return full.Length > 0 && Directory.Exists(full); } }

        // Bytes: o do arquivo, ou de tudo dentro da pasta
        public double Size
        {
            get
            {
                string p = Need();
                try
                {
                    if (File.Exists(p)) return new FileInfo(p).Length;
                    if (Directory.Exists(p)) return new DirectoryInfo(p).EnumerateFiles("*", SearchOption.AllDirectories).Sum(f => (double)f.Length);
                }
                catch (Exception ex) { throw Fail(ex, p); }
                throw NotFound(p);
            }
        }

        public Datetime Modified { get { return Stamp(true); } }
        public Datetime Created { get { return Stamp(false); } }

        Datetime Stamp(bool modified)
        {
            string p = Need();
            if (!Exists) throw NotFound(p);
            DateTime t = File.Exists(p) ? (modified ? File.GetLastWriteTime(p) : File.GetCreationTime(p))
                : (modified ? Directory.GetLastWriteTime(p) : Directory.GetCreationTime(p));
            return Datetime.From(t, TimeZoneInfo.Local);
        }

        // Partes acrescentadas ao caminho; uma parte com unidade (C:\...) recomeça dela
        public Path Join(object P1, object P2, object P3, object P4, object P5, object P6, object P7, object P8, object P9, object P10)
        {
            string result = Need();
            foreach (object part in new[] { P1, P2, P3, P4, P5, P6, P7, P8, P9, P10 })
            {
                if (Interop.IsMissing(part) || part == null) continue;
                string s = Interop.Text(Interop.Unwrap(part)).Replace('/', '\\');
                if (s.Length == 0) continue;
                try { result = IOPath.Combine(result, s); }
                catch (ArgumentException) { throw Interop.Error(52, "ValueError: parte de caminho inválida: '" + s + "'"); }
            }
            return From(IOPath.GetFullPath(result));
        }

        public Path WithName(string Name)
        {
            return From(IOPath.Combine(Parent.full, Name ?? ""));
        }

        // "csv" ou ".csv"; "" tira a extensão
        public Path WithExtension(string Extension)
        {
            string ext = (Extension ?? "").Trim();
            if (ext.Length > 0 && ext[0] != '.') ext = "." + ext;
            return From(IOPath.ChangeExtension(Need().TrimEnd('\\'), ext.Length == 0 ? null : ext));
        }

        // ------------------------------------------------------------------
        // Texto
        // ------------------------------------------------------------------

        public String ReadText(object Encoding)
        {
            return (String)Interop.Wrap(Read(Encoding));
        }

        // As linhas, sem as quebras (\r\n, \n ou \r); a quebra no fim não vira uma linha vazia
        public List ReadLines(object Encoding)
        {
            string text = Read(Encoding);
            List<object> lines = Regex.Split(text, "\r\n|\n|\r").Select(l => (object)l).ToList();
            if (lines.Count > 0 && text.Length > 0 && (text.EndsWith("\n") || text.EndsWith("\r"))) lines.RemoveAt(lines.Count - 1);
            if (text.Length == 0) lines.Clear();
            return List.From(lines);
        }

        public void WriteText(object Text, object Encoding)
        {
            Write(Interop.Text(Interop.Unwrap(Text)), Encoding, false);
        }

        // Uma linha por item (array, List ou Range), terminadas por \r\n
        public void WriteLines(object Lines, object Encoding)
        {
            StringBuilder sb = new StringBuilder();
            foreach (object line in Interop.Items(Lines)) sb.Append(Interop.Text(Interop.Unwrap(line))).Append("\r\n");
            Write(sb.ToString(), Encoding, false);
        }

        public void AppendText(object Text, object Encoding)
        {
            Write(Interop.Text(Interop.Unwrap(Text)), Encoding, true);
        }

        string Read(object encoding)
        {
            string p = Need();
            if (Directory.Exists(p)) throw Interop.Error(75, "IsADirectoryError: '" + p + "' é uma pasta");
            byte[] bytes;
            try { bytes = File.ReadAllBytes(p); }
            catch (Exception ex) { throw Fail(ex, p); }
            if (Given(encoding)) return EncodingOf(encoding, false).GetString(bytes).TrimStart('\uFEFF');
            // Sem codificação: a da marca no início; sem marca, UTF-8 se for válido; senão, a do Windows
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return new UTF8Encoding(false).GetString(bytes, 3, bytes.Length - 3);
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return System.Text.Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) return System.Text.Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
            try { return new UTF8Encoding(false, true).GetString(bytes); }
            catch (DecoderFallbackException) { return System.Text.Encoding.Default.GetString(bytes); }
        }

        void Write(string text, object encoding, bool append)
        {
            string p = Need();
            if (Directory.Exists(p)) throw Interop.Error(75, "IsADirectoryError: '" + p + "' é uma pasta");
            Encoding enc = Given(encoding) ? EncodingOf(encoding, true) : new UTF8Encoding(false);
            try
            {
                // Ao acrescentar num arquivo que já tem conteúdo, sem repetir a marca do início
                bool continuing = append && File.Exists(p) && new FileInfo(p).Length > 0;
                using (FileStream fs = new FileStream(p, append ? FileMode.Append : FileMode.Create, FileAccess.Write, FileShare.Read))
                {
                    if (!continuing) { byte[] bom = enc.GetPreamble(); fs.Write(bom, 0, bom.Length); }
                    byte[] data = enc.GetBytes(text);
                    fs.Write(data, 0, data.Length);
                }
            }
            catch (Exception ex) { throw Fail(ex, p); }
        }

        // "utf-8" (sem marca), "utf-8-sig" (com a marca, que o Excel reconhece ao abrir um CSV),
        // "utf-16", "ansi"/"windows-1252", "latin-1" ou outro nome do Windows
        static Encoding EncodingOf(object encoding, bool writing)
        {
            string name = Interop.Text(Interop.Unwrap(encoding)).Trim().ToLowerInvariant().Replace("_", "-");
            switch (name)
            {
                case "utf-8": case "utf8": return new UTF8Encoding(false);
                case "utf-8-sig": case "utf8-sig": return new UTF8Encoding(true);
                case "utf-16": case "utf16": case "unicode": return new UnicodeEncoding(false, writing);
                case "ansi": case "windows-1252": case "cp1252": return System.Text.Encoding.GetEncoding(1252);
                case "latin-1": case "latin1": case "iso-8859-1": return System.Text.Encoding.GetEncoding(28591);
            }
            try { return System.Text.Encoding.GetEncoding(name); }
            catch (ArgumentException) { throw Interop.Error(5, "LookupError: codificação desconhecida: '" + name + "'"); }
        }

        // ------------------------------------------------------------------
        // Pastas
        // ------------------------------------------------------------------

        // Arquivos da pasta (padrão: todos), em ordem de nome
        public List Files(object Pattern)
        {
            return Sorted(Entries(false, Given(Pattern) ? Interop.Text(Interop.Unwrap(Pattern)) : "*"));
        }

        public List Folders(object Pattern)
        {
            return Sorted(Entries(true, Given(Pattern) ? Interop.Text(Interop.Unwrap(Pattern)) : "*"));
        }

        IEnumerable<string> Entries(bool folders, string pattern)
        {
            string p = Folder();
            try
            {
                Regex match = Wildcard(pattern);
                IEnumerable<string> all = folders ? Directory.EnumerateDirectories(p) : Directory.EnumerateFiles(p);
                return all.Where(e => match.IsMatch(IOPath.GetFileName(e))).ToList();
            }
            catch (Exception ex) { throw Fail(ex, p); }
        }

        // Arquivos e pastas pelo padrão: "*.csv", "dados\*.csv", "**\*.csv" (** = esta pasta e todas
        // as de dentro), "relatorio_??.xlsx"
        public List Glob(string Pattern)
        {
            string p = Folder();
            string[] parts = (Pattern ?? "").Replace('/', '\\').Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) throw Interop.Error(5, "ValueError: o padrão do Glob está vazio");
            HashSet<string> found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try { Walk(p, parts, 0, found); }
            catch (Exception ex) { throw Fail(ex, p); }
            return Sorted(found);
        }

        static void Walk(string dir, string[] parts, int index, HashSet<string> found)
        {
            string part = parts[index];
            bool last = index == parts.Length - 1;
            if (part == "**")
            {
                if (last) { foreach (string d in Directory.EnumerateDirectories(dir, "*", SearchOption.AllDirectories)) found.Add(d); return; }
                Walk(dir, parts, index + 1, found);                                  // nenhuma pasta
                foreach (string d in Directory.EnumerateDirectories(dir)) Walk(d, parts, index, found); // uma ou mais
                return;
            }
            Regex match = Wildcard(part);
            if (last)
            {
                foreach (string e in Directory.EnumerateFileSystemEntries(dir)) if (match.IsMatch(IOPath.GetFileName(e))) found.Add(e);
                return;
            }
            foreach (string d in Directory.EnumerateDirectories(dir)) if (match.IsMatch(IOPath.GetFileName(d))) Walk(d, parts, index + 1, found);
        }

        static Regex Wildcard(string pattern)
        {
            string re = "^" + Regex.Escape(pattern).Replace("\\*", ".*").Replace("\\?", ".") + "$";
            return new Regex(re, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        static List Sorted(IEnumerable<string> paths)
        {
            return List.From(paths.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).Select(x => (object)From(x)));
        }

        string Folder()
        {
            string p = Need();
            if (File.Exists(p)) throw Interop.Error(75, "NotADirectoryError: '" + p + "' é um arquivo, não uma pasta");
            if (!Directory.Exists(p)) throw NotFound(p);
            return p;
        }

        // Cria a pasta e as de cima que faltarem (já existir não é erro); devolve o próprio Path
        public Path Mkdir()
        {
            string p = Need();
            if (File.Exists(p)) throw Interop.Error(58, "FileExistsError: já existe um arquivo em '" + p + "'");
            try { Directory.CreateDirectory(p); }
            catch (Exception ex) { throw Fail(ex, p); }
            return this;
        }

        // Apaga o arquivo, ou a pasta: vazia, ou com tudo dentro com Recursive:=True
        public void Delete(object Recursive)
        {
            string p = Need();
            try
            {
                if (File.Exists(p)) { File.SetAttributes(p, FileAttributes.Normal); File.Delete(p); return; }
                if (!Directory.Exists(p)) throw NotFound(p);
                bool all = Given(Recursive) && Convert.ToBoolean(Interop.Unwrap(Recursive), CultureInfo.InvariantCulture);
                if (!all && Directory.EnumerateFileSystemEntries(p).Any())
                    throw Interop.Error(75, "OSError: a pasta '" + p + "' não está vazia (use Delete True para apagar com tudo dentro)");
                Directory.Delete(p, all);
            }
            catch (COMException) { throw; }
            catch (Exception ex) { throw Fail(ex, p); }
        }

        // Copia para Destination (um caminho, ou uma pasta que já existe: a cópia vai para dentro dela)
        public Path CopyTo(object Destination, object Overwrite)
        {
            string from = Need(), to = Target(Destination);
            bool overwrite = Given(Overwrite) && Convert.ToBoolean(Interop.Unwrap(Overwrite), CultureInfo.InvariantCulture);
            try
            {
                if (File.Exists(from))
                {
                    if (File.Exists(to) && !overwrite) throw Interop.Error(58, "FileExistsError: '" + to + "' já existe (use Overwrite:=True)");
                    Directory.CreateDirectory(IOPath.GetDirectoryName(to));
                    File.Copy(from, to, overwrite);
                }
                else if (Directory.Exists(from))
                {
                    if (Directory.Exists(to) && !overwrite) throw Interop.Error(58, "FileExistsError: '" + to + "' já existe (use Overwrite:=True)");
                    CopyFolder(from, to);
                }
                else throw NotFound(from);
            }
            catch (COMException) { throw; }
            catch (Exception ex) { throw Fail(ex, to); }
            return From(to);
        }

        public Path MoveTo(object Destination, object Overwrite)
        {
            string from = Need(), to = Target(Destination);
            bool overwrite = Given(Overwrite) && Convert.ToBoolean(Interop.Unwrap(Overwrite), CultureInfo.InvariantCulture);
            try
            {
                bool file = File.Exists(from);
                if (!file && !Directory.Exists(from)) throw NotFound(from);
                if (File.Exists(to) || Directory.Exists(to))
                {
                    if (!overwrite) throw Interop.Error(58, "FileExistsError: '" + to + "' já existe (use Overwrite:=True)");
                    if (File.Exists(to)) File.Delete(to); else Directory.Delete(to, true);
                }
                Directory.CreateDirectory(IOPath.GetDirectoryName(to));
                if (file) File.Move(from, to);
                else if (string.Equals(IOPath.GetPathRoot(from), IOPath.GetPathRoot(to), StringComparison.OrdinalIgnoreCase)) Directory.Move(from, to);
                else { CopyFolder(from, to); Directory.Delete(from, true); } // outra unidade: o Directory.Move não atravessa
            }
            catch (COMException) { throw; }
            catch (Exception ex) { throw Fail(ex, to); }
            return From(to);
        }

        string Target(object destination)
        {
            string to = Of(destination).full;
            if (to.Length == 0) throw Interop.Error(5, "ValueError: o destino está vazio");
            // Uma pasta que já existe: o arquivo (ou a pasta) vai para dentro dela, com o mesmo nome
            if (Directory.Exists(to) && !string.Equals(to, full, StringComparison.OrdinalIgnoreCase) &&
                (File.Exists(full) || !string.Equals(IOPath.GetFileName(to), Name, StringComparison.OrdinalIgnoreCase)))
                to = IOPath.Combine(to, Name);
            return to;
        }

        static void CopyFolder(string from, string to)
        {
            Directory.CreateDirectory(to);
            foreach (string f in Directory.EnumerateFiles(from)) File.Copy(f, IOPath.Combine(to, IOPath.GetFileName(f)), true);
            foreach (string d in Directory.EnumerateDirectories(from)) CopyFolder(d, IOPath.Combine(to, IOPath.GetFileName(d)));
        }

        // ------------------------------------------------------------------
        // Erros com nome, e o caminho na mensagem
        // ------------------------------------------------------------------

        static bool Given(object v) { return v != null && !Interop.IsMissing(v) && !(v is DBNull); }

        static COMException NotFound(string p)
        {
            return Interop.Error(53, "FileNotFoundError: '" + p + "' não existe");
        }

        static COMException Fail(Exception ex, string p)
        {
            COMException com = ex as COMException;
            if (com != null) return com;
            if (ex is FileNotFoundException || ex is DirectoryNotFoundException) return NotFound(p);
            if (ex is UnauthorizedAccessException) return Interop.Error(70, "PermissionError: sem permissão para '" + p + "'");
            if (ex is PathTooLongException) return Interop.Error(52, "OSError: caminho longo demais: '" + p + "'");
            if (ex is IOException) return Interop.Error(75, "OSError: '" + p + "': " + ex.Message);
            return Interop.Error(75, "OSError: '" + p + "': " + ex.Message);
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHGetKnownFolderPath(ref Guid id, int flags, IntPtr token, out IntPtr path);

        static string KnownFolder(Guid id, string fallback)
        {
            IntPtr p;
            if (SHGetKnownFolderPath(ref id, 0, IntPtr.Zero, out p) != 0) return fallback;
            try { return Marshal.PtrToStringUni(p); }
            finally { Marshal.FreeCoTaskMem(p); }
        }
    }
}
