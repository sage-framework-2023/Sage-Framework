using System;
using System.Runtime.InteropServices;

namespace SageTypes
{
    // Membros globais da biblioteca Sage: usados sem Dim e sem New, como os módulos do Python
    // (json.loads, requests.get) ou o Application do Excel:
    //
    //   Set d = Json.Loads(texto)            ' ou Sage.Json.Loads(texto)
    //   Set r = Requests.Get(url)
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
    }

    // Classe cujos membros ficam globais no VBA (TYPEFLAG_FAPPOBJECT, aplicado pelo install.ps1)
    [AttributeUsage(AttributeTargets.Class)]
    sealed class AppObjectAttribute : Attribute { }
}
