using System;
using System.Runtime.CompilerServices;
using Microsoft.Win32;

namespace SageEditor
{
    // Adiciona a referência à biblioteca do Sage Framework (Sage.StringS...) ao projeto
    // VBA cujo código está aberto, para ela já vir marcada em Ferramentas > Referências.
    // No VBA as referências ficam em cada pasta de trabalho; não há como marcar uma
    // biblioteca para todas. Cada projeto é tratado uma vez por sessão.
    //
    // Fica como está: projeto protegido; em execução ou depuração (adicionar ali faria
    // o VBA pedir para resetar o projeto, então tenta de novo depois); projeto que usa
    // o Sage.xlam (legado), cujo nome "Sage" colidiria com o da biblioteca.
    static class AutoReference
    {
        const string LibId = "{3F8E2A61-7C4B-4E9D-A215-6B0C9D8E7F14}"; // SageTypes\src\AssemblyInfo.cs
        const string LibraryName = "Sage";
        const int vbext_pp_locked = 1, vbext_vm_Design = 2;
        const int PollEvery = 7; // ticks do timer de 150 ms: ~1 s

        // Por identidade do objeto COM, sem segurá-lo (referências presas impedem o Excel de fechar)
        static readonly ConditionalWeakTable<object, object> handled = new ConditionalWeakTable<object, object>();
        static readonly object Marker = new object();
        static int tick;

        static bool Installed
        {
            get
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Classes\TypeLib\" + LibId))
                    return k != null;
            }
        }

        public static void Poll(dynamic vbe)
        {
            if (++tick % PollEvery != 0 || !Settings.AutoReference) return;
            dynamic pane = vbe.ActiveCodePane;
            if (pane == null) return;
            object project = pane.CodeModule.Parent.Collection.Parent; // CodeModule > VBComponent > VBComponents > VBProject
            object ignored;
            if (project == null || handled.TryGetValue(project, out ignored)) return;

            dynamic p = project;
            if ((int)p.Mode != vbext_vm_Design) return; // tenta de novo quando o código parar
            handled.Add(project, Marker);
            if ((int)p.Protection == vbext_pp_locked || (string)p.Name == LibraryName || !Installed) return;

            foreach (dynamic r in p.References)
            {
                string guid = "", name = "";
                try { guid = r.Guid; name = r.Name; }
                catch (Exception) { } // referência quebrada ("FALTANDO")
                if (string.Equals(guid, LibId, StringComparison.OrdinalIgnoreCase)) return; // já tem
                if (name == LibraryName) return;                                            // Sage.xlam
            }

            try
            {
                p.References.AddFromGuid(LibId, 1, 0);
                Log.Info("Referência ao Sage Framework adicionada ao projeto " + (string)p.Name);
            }
            catch (Exception ex) { Log.Error(new Exception("Referência ao Sage Framework no projeto " + (string)p.Name, ex)); }
        }
    }
}
