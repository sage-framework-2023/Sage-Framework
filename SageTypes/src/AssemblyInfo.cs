using System.Reflection;
using System.Runtime.InteropServices;

// Biblioteca de tipos "Sage" (Ferramentas > Referências no VBA; o nome é definido
// pelo install.ps1 ao gerar o .tlb). O GUID é o LIBID: não mude, ou as pastas de
// trabalho que já a referenciam ficam com "FALTANDO".
//
// A versão é a da release (o build-release.ps1 confere): a biblioteca de tipos fica com
// principal.secundária (1.1) e o nome em Referências com os três números ("Sage Framework
// 1.1.0"). Mudar o número secundário a cada release faz o VBA reler a biblioteca (com a
// mesma versão ele mantinha a cópia antiga em cache), e pastas salvas com uma versão
// anterior continuam funcionando: o VBA usa a mais nova com o mesmo número principal. Só
// mude o principal (2.0) junto com mudanças que quebram o código de quem usa.
[assembly: AssemblyTitle("SageTypes")]
[assembly: AssemblyDescription("Sage Framework")] // o nome em Ferramentas > Referências, com a versão: "Sage Framework 1.1.0"
[assembly: AssemblyVersion("1.3.0.0")]
[assembly: Guid("3F8E2A61-7C4B-4E9D-A215-6B0C9D8E7F14")]
[assembly: ComVisible(false)] // só o que for marcado com [ComVisible(true)]
