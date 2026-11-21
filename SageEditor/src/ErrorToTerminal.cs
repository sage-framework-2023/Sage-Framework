using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace SageEditor
{
    // Erros em tempo de execução do VBA na aba Terminal, no lugar da caixa "Erro em tempo de
    // execução" (Editor: Erros no Terminal, sage.errorsInTerminal):
    //
    //   Pasta1.xlsm\Módulos\Modulo1\Calcular, linha 12
    //       valores(5) = 1
    //   Erro em tempo de execução '9': Subscrito fora do intervalo
    //
    // Como no Python: primeiro onde (o arquivo, a pasta do componente na janela Projeto, o
    // componente e o procedimento), depois o código da linha e, por último, a mensagem.
    //
    // Também o "Erro de compilação" (variável não definida, Sub que não existe...), com o trecho do
    // erro no lugar da linha.
    //
    // A caixa é reconhecida quando vai aparecer (o hook CBT do ThemeEngine chama Activating):
    // uma caixa do VBE com os botões Fim e Depurar, pelos identificadores (o texto muda com o
    // idioma). O Sage aciona Depurar, para o VBE marcar a linha do erro; quando a caixa fecha
    // (Poll), lê o módulo, o procedimento e a linha, encerra o código (Executar > Redefinir) no
    // modo Mostrar e Parar e escreve no Terminal. O cursor fica na linha do erro.
    //
    // Fica a caixa de sempre: com a configuração desligada, com o editor do VBA fechado ou
    // minimizado (o erro de uma macro rodada pela planilha não sumiria sem ninguém ver), sem o
    // Painel, e nas caixas que não são de erro (o MsgBox do próprio código nunca é tocado).
    static class ErrorToTerminal
    {
        public static dynamic Vbe;
        static IntPtr vbeWindow;

        // Botões da caixa "Erro em tempo de execução" do VBE, conferidos no Excel (o texto fica num
        // Static só, 4803: "Erro em tempo de execução '9':", linha em branco e a descrição)
        const int IdEnd = 4800, IdDebug = 4801, IdContinue = 4802;
        // Caixa "Erro de compilação" do VBE: uma caixa de mensagem com OK e Ajuda, com o título
        // abaixo (o MsgBox do código tem o título do Excel, "Microsoft Excel", se não for trocado)
        // (quando a caixa vai aparecer, o OK é o 1; depois, o Windows o troca por 2. Sim, Não,
        // Repetir... seriam outros identificadores)
        const int IdOk = 1, IdOkLater = 2, IdHelp = 9;
        static readonly int[] OtherButtons = { 3, 4, 5, 6, 7 };
        const string CompileTitle = "Microsoft Visual Basic for Applications";

        const int BM_CLICK = 0x00F5;
        const int ResetControlId = 228; // Executar > Redefinir

        static IntPtr pending;       // a caixa que o Sage está fechando
        static string pendingMessage;
        static int pendingSince, closedSince;
        static bool closed;

        [DllImport("user32.dll")] static extern IntPtr GetDlgItem(IntPtr dialog, int id);
        [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);

        public static void Start(IntPtr vbe) { vbeWindow = vbe; }

        // Pelo hook CBT (HCBT_ACTIVATE), quando uma caixa (#32770) da thread vai aparecer
        public static void Activating(IntPtr dialog)
        {
            try
            {
                if (pending != IntPtr.Zero || Settings.Errors == Settings.ErrorMode.Off || !TerminalWindow.Ready) return;
                if (vbeWindow == IntPtr.Zero || !Native.IsWindowVisible(vbeWindow) || IsIconic(vbeWindow)) return;
                IntPtr button = RuntimeError(dialog);
                if (button == IntPtr.Zero) button = CompileError(dialog);
                if (button == IntPtr.Zero) return;
                pendingMessage = Message(dialog);
                pending = dialog;
                pendingSince = Environment.TickCount;
                closed = false;
                PostMessage(button, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
            }
            catch (Exception ex) { Log.Error(ex); }
        }

        // Erro em tempo de execução: o botão Depurar (o VBE marca a linha), ou zero se não for a caixa
        static IntPtr RuntimeError(IntPtr dialog)
        {
            IntPtr end = GetDlgItem(dialog, IdEnd), debug = GetDlgItem(dialog, IdDebug), resume = GetDlgItem(dialog, IdContinue);
            if (!IsButton(end) || !IsButton(debug) || !IsButton(resume)) return IntPtr.Zero;
            return IsWindowEnabled(debug) ? debug : IntPtr.Zero; // sem Depurar (projeto protegido): fica a caixa
        }

        // Erro de compilação: o botão OK (o VBE marca o trecho do erro), ou zero se não for a caixa
        static IntPtr CompileError(IntPtr dialog)
        {
            StringBuilder title = new StringBuilder(256);
            GetWindowText(dialog, title, title.Capacity);
            if (title.ToString() != CompileTitle) return IntPtr.Zero;
            IntPtr ok = GetDlgItem(dialog, IdOk), help = GetDlgItem(dialog, IdHelp);
            if (ok == IntPtr.Zero) ok = GetDlgItem(dialog, IdOkLater);
            if (!IsButton(ok) || !IsButton(help)) return IntPtr.Zero;
            foreach (int id in OtherButtons) if (GetDlgItem(dialog, id) != IntPtr.Zero) return IntPtr.Zero;
            return ok;
        }

        static bool IsButton(IntPtr hwnd) { return hwnd != IntPtr.Zero && Native.ClassName(hwnd) == "Button"; }

        // O texto da caixa: as linhas de texto (Static) que não são de botões, numa linha só
        static string Message(IntPtr dialog)
        {
            List<string> parts = new List<string>();
            Native.EnumChildWindows(dialog, delegate(IntPtr child, IntPtr l)
            {
                if (Native.ClassName(child) != "Static") return true;
                StringBuilder sb = new StringBuilder(1024);
                GetWindowText(child, sb, sb.Capacity);
                // "Erro em tempo de execução '9':", linha em branco e a descrição: numa linha só
                foreach (string line in sb.ToString().Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                    if (line.Trim().Length > 0) parts.Add(line.Trim());
                return true;
            }, IntPtr.Zero);
            return string.Join(" ", parts.ToArray());
        }

        // Pelo timer do Connect: depois que a caixa fecha, o VBE está parado na linha do erro
        public static void Poll()
        {
            if (pending == IntPtr.Zero) return;
            if (Native.IsWindow(pending))
            {
                if (unchecked(Environment.TickCount - pendingSince) > 5000) pending = IntPtr.Zero; // a caixa não fechou: desiste
                return;
            }
            // A caixa fechou; o VBE ainda termina de parar o código na linha (Redefinir habilitado)
            if (!closed) { closed = true; closedSince = Environment.TickCount; }
            if (!Paused() && unchecked(Environment.TickCount - closedSince) < 2000) return;
            pending = IntPtr.Zero;
            try { Report(pendingMessage); }
            catch (Exception ex) { Log.Error(ex); }
        }

        static void Report(string message)
        {
            StringBuilder text = new StringBuilder();
            dynamic pane = null;
            int sl = 0, sc = 0, el = 0, ec = 0;
            try
            {
                pane = Vbe.ActiveCodePane;
                pane.GetSelection(ref sl, ref sc, ref el, ref ec);
                dynamic module = pane.CodeModule;
                dynamic component = module.Parent;
                // Arquivo\Pasta\Componente\Procedimento, como o caminho na janela Projeto
                string where = FileOf(component.Collection.Parent) + "\\" + Folder((int)component.Type) + "\\" + (string)component.Name;
                try
                {
                    int kind = 0;
                    string procedure = (string)module.ProcOfLine(sl, ref kind);
                    if (!string.IsNullOrEmpty(procedure)) where += "\\" + procedure;
                }
                catch (Exception) { } // fora de um procedimento (declarações)
                text.Append(string.Format(Strings.ErrorAt, where, sl));
                string code = ((string)module.Lines(sl, 1) ?? "").Trim();
                if (code.Length > 0) text.Append("\n    ").Append(code);
                text.Append("\n");
            }
            catch (Exception ex) { Log.Error(ex); pane = null; }

            text.Append(message);
            if (Settings.Errors == Settings.ErrorMode.Stop) Reset();
            else text.Append("\n").Append(Strings.ErrorPaused);

            lastReport = Environment.TickCount;
            TerminalWindow.ShowError(text.ToString());
            // O cursor na linha do erro, como o clique numa linha de erro de outros terminais
            if (pane != null && sl > 0)
            {
                try { pane.SetSelection(sl, sc, el, ec); pane.Show(); }
                catch (Exception) { }
            }
        }

        // Quando o último erro foi para o Terminal: o comando VBA do Terminal recebe o mesmo erro de
        // volta (o Application.Run) e não precisa mostrá-lo de novo (ExcelBridge.Run)
        static int lastReport = Environment.TickCount - int.MaxValue / 2;
        public static bool ReportedSince(int tick) { return unchecked(lastReport - tick) >= 0; }

        // O nome do arquivo do projeto (Pasta1.xlsm). Sem arquivo (pasta de trabalho ainda não
        // salva), o nome da pasta de trabalho do Excel que tem o projeto (Pasta1)
        internal static string FileOf(dynamic project)
        {
            try
            {
                string file = (string)project.FileName;
                if (!string.IsNullOrEmpty(file)) return System.IO.Path.GetFileName(file);
            }
            catch (Exception) { } // o VBE dá erro no FileName de um projeto sem arquivo
            try
            {
                dynamic app = ExcelBridge.Application();
                if (app != null)
                    foreach (dynamic book in app.Workbooks)
                        if (SameObject(book.VBProject, project)) return (string)book.Name;
            }
            catch (Exception) { }
            try { return (string)project.Name; }
            catch (Exception) { return "."; }
        }

        static bool SameObject(object a, object b)
        {
            if (a == null || b == null) return false;
            IntPtr x = Marshal.GetIUnknownForObject(a), y = Marshal.GetIUnknownForObject(b);
            try { return x == y; }
            finally { Marshal.Release(x); Marshal.Release(y); }
        }

        // A pasta do componente na janela Projeto (vbext_ComponentType)
        internal static string Folder(int type)
        {
            switch (type)
            {
                case 1: return Strings.FolderModules;    // vbext_ct_StdModule
                case 2: return Strings.FolderClasses;    // vbext_ct_ClassModule
                case 3: return Strings.FolderForms;      // vbext_ct_MSForm
                case 100: return Strings.FolderDocuments; // vbext_ct_Document: planilhas e EstaPastaDeTrabalho
                default: return Strings.FolderModules;
            }
        }

        // Executar > Redefinir: encerra o código parado na linha do erro (como o botão Fim)
        static void Reset()
        {
            try
            {
                // O comando roda quando o VBE volta ao laço de mensagens (Paused() ainda é true logo depois)
                dynamic control = ResetControl();
                if (control != null) control.Execute();
            }
            catch (Exception ex) { Log.Error(ex); }
        }


        // O código está parado (o Redefinir só fica habilitado assim)
        // Algum projeto com o código parado (VBProject.Mode = vbext_vm_Break). O Redefinir do menu
        // não serve para isso: fica habilitado mesmo sem código parado.
        static bool Paused()
        {
            try
            {
                foreach (dynamic project in Vbe.VBProjects)
                    if ((int)project.Mode == 1) return true;
            }
            catch (Exception) { }
            return false;
        }

        static dynamic ResetControl()
        {
            return Vbe.CommandBars.FindControl(Type.Missing, ResetControlId);
        }
    }
}
