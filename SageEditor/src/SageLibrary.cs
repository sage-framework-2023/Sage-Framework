using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace SageEditor
{
    // Opções das Configurações sobre a biblioteca Sage Framework no VBA

    // A referência à biblioteca (Ferramentas > Referências) no projeto ativo do VBE
    static class SageReference
    {
        const int vbext_pp_locked = 1, vbext_vm_Design = 2;

        public enum Status { Missing, Present, NoProject, NotInstalled }

        // Na thread do Excel. project: o nome do projeto ativo, para a descrição.
        public static Status Get(dynamic vbe, out string project)
        {
            project = null;
            if (!AutoReference.Installed) return Status.NotInstalled;
            dynamic p = Active(vbe);
            if (p == null) return Status.NoProject;
            project = Name(p);
            return Find(p) != null ? Status.Present : Status.Missing;
        }

        // Na thread do Excel. Devolve null quando deu certo, ou a mensagem para o usuário.
        public static string Set(dynamic vbe, bool on)
        {
            try
            {
                if (!AutoReference.Installed) return Strings.SageReferenceNotInstalled;
                dynamic p = Active(vbe);
                if (p == null) return Strings.SageReferenceNoProject;
                if ((int)p.Protection == vbext_pp_locked) return Strings.SageReferenceLocked;
                if ((int)p.Mode != vbext_vm_Design) return Strings.SageReferenceRunning;
                dynamic r = Find(p);
                if (on && r == null)
                {
                    p.References.AddFromGuid(AutoReference.LibId, 1, 0);
                    Log.Info("Referência ao Sage Framework adicionada em " + Name(p));
                }
                else if (!on && r != null)
                {
                    p.References.Remove(r);
                    Log.Info("Referência ao Sage Framework removida de " + Name(p));
                }
                return null;
            }
            catch (Exception ex)
            {
                Log.Error(new Exception("Referência ao Sage Framework", ex));
                return ex.Message;
            }
        }

        static dynamic Active(dynamic vbe)
        {
            try { return vbe.ActiveVBProject; }
            catch (Exception) { return null; }
        }

        static string Name(dynamic p)
        {
            try
            {
                string file = p.FileName;
                if (!string.IsNullOrEmpty(file)) return Path.GetFileName(file);
            }
            catch (Exception) { } // pasta ainda não salva
            try { return (string)p.Name; }
            catch (Exception) { return ""; }
        }

        static dynamic Find(dynamic p)
        {
            foreach (dynamic r in p.References)
            {
                string guid = "";
                try { guid = r.Guid; }
                catch (Exception) { } // referência quebrada ("FALTANDO")
                if (string.Equals(guid, AutoReference.LibId, StringComparison.OrdinalIgnoreCase)) return r;
            }
            return null;
        }
    }

    // Os controles do Sage (Sage TextBox) na Caixa de ferramentas dos UserForms. Quem mexe na
    // Caixa é o SageShortcuts, com o Excel fechado (o VBE mantém o arquivo dela aberto); daqui
    // só sai o pedido, em %APPDATA%\Sage\toolbox.txt (o mesmo arquivo em que o SageShortcuts
    // anota o que já fez):
    //   offered=<CLSIDs>   controles já postos uma vez (não voltam sozinhos se desmarcados)
    //   state=on|off       o último estado aplicado
    //   request=add|remove o pedido ainda não aplicado
    static class ToolboxControls
    {
        const string ControlClsid = "{0AD72DAF-5E3F-47CA-8C4C-3E5050C6DFC5}"; // Sage.TextBox

        static string FilePath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Sage\toolbox.txt"); } }

        public static bool Installed
        {
            get
            {
                using (Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(@"CLSID\" + ControlClsid + @"\Control"))
                    return k != null;
            }
        }

        // O que vale (ou valerá, quando o Excel fechar)
        public static bool On
        {
            get
            {
                Dictionary<string, string> v = Read();
                string request;
                if (v.TryGetValue("request", out request)) return request == "add";
                string state;
                if (v.TryGetValue("state", out state)) return state == "on";
                return Installed; // ainda não aplicado: o SageShortcuts os põe na primeira vez
            }
        }

        public static void Set(bool on)
        {
            Dictionary<string, string> v = Read();
            string state;
            bool applied = v.TryGetValue("state", out state) ? state == "on" : false;
            if (v.ContainsKey("state") && applied == on) v.Remove("request");
            else v["request"] = on ? "add" : "remove";
            Write(v);
        }

        static Dictionary<string, string> Read()
        {
            Dictionary<string, string> v = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (File.Exists(FilePath))
                    foreach (string line in File.ReadAllLines(FilePath))
                    {
                        int i = line.IndexOf('=');
                        if (i > 0) v[line.Substring(0, i).Trim()] = line.Substring(i + 1).Trim();
                    }
            }
            catch (Exception ex) { Log.Error(ex); }
            return v;
        }

        static void Write(Dictionary<string, string> v)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                List<string> lines = new List<string>();
                foreach (KeyValuePair<string, string> kv in v) lines.Add(kv.Key + "=" + kv.Value);
                File.WriteAllLines(FilePath, lines.ToArray());
            }
            catch (Exception ex) { Log.Error(ex); }
        }
    }
}
