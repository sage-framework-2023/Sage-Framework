using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace SageTypes
{
    // Membros globais da biblioteca Sage: usados sem Dim e sem New, como os módulos do Python
    // (json.loads, requests.get) ou o Application do Excel:
    //
    //   Set d = Json.Loads(texto)            ' ou Sage.Json.Loads(texto)
    //   Set r = Requests.Get(url)
    //   Debug.Print GetUser()                ' getpass.getuser(): o usuário logado no Windows
    //   Sleep 1.5                            ' time.sleep(1.5), sem congelar a tela do Excel
    //
    // O install.ps1 marca a classe como "app object" no .tlb ([AppObject]): o VBA cria uma
    // instância por projeto na primeira vez que um membro é usado.
    //
    // Interface dual: acrescente membros só no fim, com o próximo DispId (ver StringS).
    [ComVisible(true), Guid("B35FF712-35A6-4A4B-A8BE-D58FC7A3D471"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _Globals
    {
        [DispId(1)] Json Json { get; }
        [DispId(2)] Requests Requests { get; }
        [DispId(3)] StringS GetUser([Optional] object Domain);
        [DispId(4)] void Sleep(double Seconds);
    }

    [ComVisible(true), Guid("3129D3A3-5535-4294-8923-083DAA4A6F75"), ProgId("Sage.Globals"), AppObject]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_Globals))]
    public sealed class Globals : _Globals
    {
        static readonly Json json = new Json();
        static readonly Requests requests = new Requests();

        public Globals() { }

        public Json Json { get { return json; } }
        public Requests Requests { get { return requests; } }

        // O usuário logado no Windows (pela API, como o GetUserName; não lê a variável
        // USERNAME, que pode ser alterada). Domain:=True devolve também o domínio:
        // "EMPRESA\rodrigo" (num computador fora de domínio, o nome do computador).
        public StringS GetUser(object Domain)
        {
            bool withDomain = !Interop.IsMissing(Domain) && Domain != null &&
                Convert.ToBoolean(Interop.Unwrap(Domain), CultureInfo.InvariantCulture);
            string user = Environment.UserName;
            if (withDomain) user = Environment.UserDomainName + "\\" + user;
            return (StringS)Interop.Wrap(user);
        }

        // Espera, em segundos (aceita frações: 0.5). Enquanto espera, o Excel continua
        // redesenhando a tela e não aparece como "Não respondendo"; cliques e teclas ficam na
        // fila até a macro terminar, como em qualquer macro ocupada (nada roda no meio dela).
        public void Sleep(double Seconds)
        {
            if (double.IsNaN(Seconds) || Seconds < 0) throw Interop.Error(5, "ValueError: sleep length must be non-negative");
            Stopwatch clock = Stopwatch.StartNew();
            long total = (long)Math.Round(Seconds * 1000);
            MSG msg;
            while (true)
            {
                long left = total - clock.ElapsedMilliseconds;
                if (left <= 0) break;
                // Acorda quando chega pintura (ou a cada 50 ms, para a fila não parecer parada)
                MsgWaitForMultipleObjects(0, null, false, (uint)Math.Min(left, 50), QS_PAINT);
                while (PeekMessage(out msg, IntPtr.Zero, WM_PAINT, WM_PAINT, PM_REMOVE)) DispatchMessage(ref msg);
                PeekMessage(out msg, IntPtr.Zero, 0, 0, PM_NOREMOVE); // o Windows vê a fila sendo lida
            }
        }

        const uint QS_PAINT = 0x0020, PM_NOREMOVE = 0x0000, PM_REMOVE = 0x0001, WM_PAINT = 0x000F;

        [StructLayout(LayoutKind.Sequential)]
        struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam, lParam; public uint time; public int x, y; }

        [DllImport("user32.dll")] static extern uint MsgWaitForMultipleObjects(uint count, IntPtr[] handles, bool waitAll, uint milliseconds, uint wakeMask);
        [DllImport("user32.dll")] static extern bool PeekMessage(out MSG msg, IntPtr hwnd, uint min, uint max, uint remove);
        [DllImport("user32.dll")] static extern IntPtr DispatchMessage(ref MSG msg);
    }

    // Classe cujos membros ficam globais no VBA (TYPEFLAG_FAPPOBJECT, aplicado pelo install.ps1)
    [AttributeUsage(AttributeTargets.Class)]
    sealed class AppObjectAttribute : Attribute { }
}
