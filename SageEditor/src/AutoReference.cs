using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;

namespace SageEditor
{
    // Pastas de trabalho novas (Ctrl+N, Arquivo > Novo, a pasta em branco ao abrir o Excel) já
    // vêm com a referência à biblioteca do Sage Framework (Sage.StringS...) em Ferramentas >
    // Referências. No VBA as referências ficam em cada pasta; não há como marcar uma biblioteca
    // para todas.
    //
    // Cada pasta é tratada uma vez só, então quem não quer a referência num arquivo a desmarca e
    // ela não volta; arquivos de sessões anteriores nunca são alterados:
    // - quando o suplemento carrega (o VBA inicia com Alt+F11 ou ao abrir um arquivo com macros),
    //   uma passada pelas pastas abertas que são novas: ainda não salvas ou criadas nesta sessão
    //   do Excel (quem cria a pasta, salva como .xlsm e só depois abre o editor também a recebe);
    // - depois, o evento NewWorkbook do Excel, uma vez por pasta criada.
    // (O VBE devolve objetos diferentes para o mesmo projeto a cada consulta, então não dá para
    // lembrar de "projetos já tratados" verificando periodicamente.)
    //
    // Fica como está: projeto protegido; projeto que usa o Sage.xlam (legado), cujo nome "Sage"
    // colidiria com o da biblioteca.
    static class AutoReference
    {
        const string LibId = "{3F8E2A61-7C4B-4E9D-A215-6B0C9D8E7F14}"; // SageTypes\src\AssemblyInfo.cs
        const string LibraryName = "Sage";
        const int vbext_pp_locked = 1, vbext_vm_Design = 2, vbext_ct_Document = 100;
        const int PollEvery = 7; // ticks do timer de 150 ms: ~1 s

        static readonly DateTime SessionStart = System.Diagnostics.Process.GetCurrentProcess().StartTime;
        static readonly Guid AppEventsId = new Guid("00024413-0000-0000-C000-000000000046");

        static bool scanned;
        static IConnectionPoint point;
        static int cookie;
        static AutoReferenceEvents sink;
        // Pastas criadas (nome -> quando), tratadas quando o projeto VBA delas aparecer no VBE
        static readonly Dictionary<string, DateTime> pending = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
        static int tick;
        static bool errorLogged;

        // Instalada pelo instalador (para todos os usuários) ou pelo install.ps1 (só o usuário atual)
        static bool Installed
        {
            get
            {
                using (RegistryKey k = Registry.ClassesRoot.OpenSubKey(@"TypeLib\" + LibId))
                    return k != null;
            }
        }

        // Timer do Connect (thread de interface do Excel)
        public static void Poll(dynamic vbe)
        {
            if (++tick % PollEvery != 0 || !Settings.AutoReference) return;
            try
            {
                if (point == null) Hook(vbe);
                if (!scanned)
                {
                    scanned = true;
                    if (!Installed) return;
                    foreach (dynamic p in vbe.VBProjects)
                        if (IsNew(p)) AddReference(p);
                }
                if (pending.Count > 0) ProcessPending(vbe);
            }
            catch (Exception ex)
            {
                // VBE ocupado: tenta no próximo ciclo; registra só a primeira vez
                if (!errorLogged) { errorLogged = true; Log.Error(new Exception("Referência automática", ex)); }
            }
        }

        public static void Shutdown()
        {
            try { if (point != null) point.Unadvise(cookie); }
            catch (Exception) { }
            if (point != null) Marshal.ReleaseComObject(point);
            point = null;
            sink = null;
            pending.Clear();
        }

        // ------------------------------------------------------------------
        // Evento NewWorkbook do Excel
        // ------------------------------------------------------------------

        // O Excel.Application vem da pasta, pelo componente do documento (EstaPastaDeTrabalho)
        static void Hook(dynamic vbe)
        {
            object app = null;
            foreach (dynamic p in vbe.VBProjects)
            {
                dynamic doc = WorkbookComponent(p);
                if (doc == null) continue;
                try { app = doc.Properties.Item("Application").Object; }
                catch (Exception) { }
                if (app != null) break;
            }
            if (app == null) return; // nenhuma pasta aberta ainda
            IConnectionPointContainer container = (IConnectionPointContainer)app;
            Guid id = AppEventsId;
            container.FindConnectionPoint(ref id, out point);
            sink = new AutoReferenceEvents();
            point.Advise(sink, out cookie);
        }

        // Chamado pelo evento (thread de interface do Excel)
        internal static void OnNewWorkbook(object workbook)
        {
            try
            {
                if (!Settings.AutoReference) return;
                string name = ((dynamic)workbook).Name;
                pending[name] = DateTime.Now; // o projeto VBA aparece no VBE logo depois
            }
            catch (Exception ex) { Log.Error(ex); }
            finally { Marshal.ReleaseComObject(workbook); }
        }

        static void ProcessPending(dynamic vbe)
        {
            if (!Installed) { pending.Clear(); return; }
            foreach (dynamic p in vbe.VBProjects)
            {
                if (pending.Count == 0) return;
                string name = WorkbookName(p);
                if (name == null || !pending.ContainsKey(name)) continue;
                if ((int)p.Mode != vbext_vm_Design) continue; // tenta de novo quando o código parar
                pending.Remove(name);
                if (IsNew(p)) AddReference(p);
            }
            // Pastas que nunca apareceram (fechadas logo depois de criadas)
            foreach (string name in new List<string>(pending.Keys))
                if ((DateTime.Now - pending[name]).TotalSeconds > 30) pending.Remove(name);
        }

        // ------------------------------------------------------------------
        // Projeto e referência
        // ------------------------------------------------------------------

        // Pasta nova: ainda não salva (sem arquivo: o FileName dá erro ou vem vazio) ou salva pela
        // primeira vez nesta sessão do Excel
        static bool IsNew(dynamic p)
        {
            string file;
            try { file = p.FileName; }
            catch (Exception) { return true; }
            if (string.IsNullOrEmpty(file)) return true;
            try { return File.Exists(file) && File.GetCreationTime(file) >= SessionStart; }
            catch (Exception) { return false; }
        }

        static void AddReference(dynamic p)
        {
            try
            {
                if ((int)p.Mode != vbext_vm_Design || (int)p.Protection == vbext_pp_locked || (string)p.Name == LibraryName) return;
                foreach (dynamic r in p.References)
                {
                    string guid = "", name = "";
                    try { guid = r.Guid; name = r.Name; }
                    catch (Exception) { } // referência quebrada ("FALTANDO")
                    if (string.Equals(guid, LibId, StringComparison.OrdinalIgnoreCase)) return; // já tem
                    if (name == LibraryName) return;                                            // Sage.xlam
                }

                // Mexer nas referências marca a pasta como alterada, e o Excel pediria para salvar uma
                // pasta em branco ao fechá-la: volta o "Saved" ao que era
                dynamic doc = WorkbookComponent(p);
                bool wasSaved = false;
                try { wasSaved = doc != null && (bool)doc.Properties.Item("Saved").Value; }
                catch (Exception) { }

                p.References.AddFromGuid(LibId, 1, 0);
                Log.Info("Referência ao Sage Framework adicionada à pasta nova " + (WorkbookName(p) ?? (string)p.Name));

                if (wasSaved)
                {
                    try { doc.Properties.Item("Saved").Value = true; }
                    catch (Exception) { }
                }
            }
            catch (Exception ex) { Log.Error(new Exception("Referência ao Sage Framework", ex)); }
        }

        // O componente do documento da pasta (EstaPastaDeTrabalho / ThisWorkbook): o único com "Saved"
        static dynamic WorkbookComponent(dynamic p)
        {
            try
            {
                foreach (dynamic c in p.VBComponents)
                {
                    if ((int)c.Type != vbext_ct_Document) continue;
                    try
                    {
                        object saved = c.Properties.Item("Saved");
                        if (saved != null) return c;
                    }
                    catch (Exception) { } // planilha, não a pasta
                }
            }
            catch (Exception) { }
            return null;
        }

        static string WorkbookName(dynamic p)
        {
            dynamic doc = WorkbookComponent(p);
            if (doc == null) return null;
            try { return (string)doc.Properties.Item("Name").Value; }
            catch (Exception) { return null; }
        }
    }

    // Eventos do Excel.Application (dispinterface AppEvents); públicos, porque o Excel os chama pelo COM
    [ComImport, Guid("00024413-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ExcelAppEvents
    {
        [DispId(0x61D)] void NewWorkbook([In, MarshalAs(UnmanagedType.IDispatch)] object workbook);
    }

    // Só o NewWorkbook; os outros eventos do Excel caem em "membro não encontrado", que o Excel ignora
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public sealed class AutoReferenceEvents : ExcelAppEvents
    {
        public void NewWorkbook(object workbook) { AutoReference.OnNewWorkbook(workbook); }
    }
}
