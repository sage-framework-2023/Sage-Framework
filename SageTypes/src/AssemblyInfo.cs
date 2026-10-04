using System.Reflection;
using System.Runtime.InteropServices;

// Biblioteca de tipos "Sage" (Ferramentas > Referências no VBA; o nome é definido
// pelo install.ps1 ao gerar o .tlb). O GUID é o LIBID: não mude, ou as pastas de
// trabalho que já a referenciam ficam com "FALTANDO". A versão também fica fixa
// (vai no registro do COM).
[assembly: AssemblyTitle("SageTypes")]
[assembly: AssemblyDescription("Sage Framework")] // o nome em Ferramentas > Referências, com a versão: "Sage Framework 1.0"
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: Guid("3F8E2A61-7C4B-4E9D-A215-6B0C9D8E7F14")]
[assembly: ComVisible(false)] // só o que for marcado com [ComVisible(true)]
