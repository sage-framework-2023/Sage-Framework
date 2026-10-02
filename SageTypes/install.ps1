# Instala os tipos do Sage para o VBA (StringS...) só para o usuário atual, sem admin.
#   powershell -ExecutionPolicy Bypass -File install.ps1              compila, copia e registra
#   powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall   remove
# Depois, em cada pasta de trabalho: Ferramentas > Referências > "Sage".
# O Excel carrega o DLL quando o VBA usa um tipo; feche-o antes de instalar ou atualizar.

param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$Target = Join-Path $env:LOCALAPPDATA 'Sage\Types'
$Dll = Join-Path $Target 'SageTypes.dll'
$Tlb = Join-Path $Target 'SageTypes.tlb'

Add-Type -TypeDefinition @'
using System;
using System.Reflection;
using System.Runtime.InteropServices;

public static class SageTypeLib
{
    [ComImport, Guid("00020406-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICreateTypeLib
    {
        void CreateTypeInfo();
        void SetName([MarshalAs(UnmanagedType.LPWStr)] string name);
        void SetVersion(); void SetGuid(); void SetDocString();
        void SetHelpFileName(); void SetHelpContext(); void SetLcid(); void SetLibFlags();
        void SaveAllChanges();
    }

    // ICreateTypeInfo2: só os métodos usados têm assinatura; os demais guardam a posição na vtable
    [ComImport, Guid("0002040E-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICreateTypeInfo2
    {
        void SetGuid(); void SetTypeFlags(); void SetDocString(); void SetHelpContext(); void SetVersion();
        void AddRefTypeInfo();
        void AddFuncDesc(int index, IntPtr funcDesc);
        void AddImplType(); void SetImplTypeFlags(); void SetAlignment(); void SetSchema(); void AddVarDesc();
        void SetFuncAndParamNames(int index, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] names, int count);
        void SetVarName(int index, [MarshalAs(UnmanagedType.LPWStr)] string name);
        void SetTypeDescAlias(); void DefineFuncAsDllEntry(); void SetFuncDocString(); void SetVarDocString();
        void SetFuncHelpContext(); void SetVarHelpContext(); void SetMops(); void SetTypeIdldesc(); void LayOut();
        void DeleteFuncDesc(int index);
    }

    sealed class Sink : ITypeLibExporterNotifySink
    {
        public void ReportEvent(ExporterEventKind kind, int code, string message)
        {
            if (kind != ExporterEventKind.NOTIF_TYPECONVERTED) Console.WriteLine(message);
        }
        public object ResolveRef(Assembly assembly) { return null; }
    }

    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    static extern void LoadTypeLibEx(string file, int regKind, out System.Runtime.InteropServices.ComTypes.ITypeLib lib);
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    static extern void RegisterTypeLibForUser(System.Runtime.InteropServices.ComTypes.ITypeLib lib, string file, string helpDir);
    [DllImport("oleaut32.dll", PreserveSig = true)]
    static extern int UnRegisterTypeLibForUser(ref Guid libId, short major, short minor, int lcid, int syskind);

    // .tlb de 64 bits (o Excel instalado é de 64 bits). O nome da biblioteca, que o VBA
    // usa para qualificar (Sage.StringS), é "Sage", e não o nome do assembly.
    public static void Export(string dll, string tlb, string libraryName)
    {
        Assembly assembly = Assembly.LoadFrom(dll);
        ICreateTypeLib lib = (ICreateTypeLib)new TypeLibConverter().ConvertAssemblyToTypeLib(
            assembly, tlb, TypeLibExporterFlags.ExportAs64Bit, new Sink());
        lib.SetName(libraryName);
        ApplyLetProperties((System.Runtime.InteropServices.ComTypes.ITypeLib)lib);
        lib.SaveAllChanges();
    }

    // O .NET exporta o setter de uma propriedade Variant só como Property Set
    // (propputref), e "d(chave) = valor" não compila no VBA. Cada método let_X das
    // interfaces vira o Property Let (propput) de X: mesmo DispId de X, mesma posição
    // na vtable (onde o .NET implementou let_X), parâmetros iguais.
    static void ApplyLetProperties(System.Runtime.InteropServices.ComTypes.ITypeLib lib)
    {
        for (int i = 0; i < lib.GetTypeInfoCount(); i++)
        {
            System.Runtime.InteropServices.ComTypes.ITypeInfo info;
            lib.GetTypeInfo(i, out info);
            IntPtr pa;
            info.GetTypeAttr(out pa);
            var attr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)Marshal.PtrToStructure(pa, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));
            info.ReleaseTypeAttr(pa);
            if (attr.typekind == System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_ENUM)
            {
                RemoveEnumPrefix(info, attr.cVars);
                continue;
            }
            if (attr.typekind != System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_DISPATCH ||
                (attr.wTypeFlags & System.Runtime.InteropServices.ComTypes.TYPEFLAGS.TYPEFLAG_FDUAL) == 0) continue;
            // A interface da vtable por trás da visão IDispatch
            int href;
            info.GetRefTypeOfImplType(-1, out href);
            System.Runtime.InteropServices.ComTypes.ITypeInfo vtable;
            info.GetRefTypeInfo(href, out vtable);
            ApplyLetProperties(vtable);
        }
    }

    static void ApplyLetProperties(System.Runtime.InteropServices.ComTypes.ITypeInfo info)
    {
        IntPtr pa;
        info.GetTypeAttr(out pa);
        var attr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)Marshal.PtrToStructure(pa, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));
        info.ReleaseTypeAttr(pa);

        for (int f = 0; f < attr.cFuncs; f++)
        {
            IntPtr pf;
            info.GetFuncDesc(f, out pf);
            var desc = (System.Runtime.InteropServices.ComTypes.FUNCDESC)Marshal.PtrToStructure(pf, typeof(System.Runtime.InteropServices.ComTypes.FUNCDESC));
            string[] names = new string[desc.cParams + 1];
            int count;
            info.GetNames(desc.memid, names, names.Length, out count);
            if (!names[0].StartsWith("let_") || desc.invkind != System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_FUNC)
            {
                info.ReleaseFuncDesc(pf);
                continue;
            }

            string property = names[0].Substring(4);
            desc.memid = MemberId(info, attr.cFuncs, property);
            desc.invkind = System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYPUT;
            IntPtr copy = Marshal.AllocHGlobal(Marshal.SizeOf(desc));
            try
            {
                Marshal.StructureToPtr(desc, copy, false);
                ICreateTypeInfo2 create = (ICreateTypeInfo2)info;
                create.DeleteFuncDesc(f);
                create.AddFuncDesc(f, copy); // copia a descrição (inclusive os parâmetros de pf)
                // Property Let: o nome do último parâmetro (o valor atribuído) não entra
                string[] newNames = new string[desc.cParams];
                newNames[0] = property;
                for (int p = 1; p < desc.cParams; p++) newNames[p] = names[p];
                create.SetFuncAndParamNames(f, newNames, newNames.Length);
            }
            finally
            {
                Marshal.FreeHGlobal(copy);
                info.ReleaseFuncDesc(pf);
            }
        }
    }

    // O .NET exporta os membros de um enum como "Enum_Membro" (sgArrayTypes_sgTuple);
    // no VBA eles ficam só com o nome do membro (sgTuple), como no Sage.xlam
    static void RemoveEnumPrefix(System.Runtime.InteropServices.ComTypes.ITypeInfo info, int vars)
    {
        string enumName, doc, helpFile;
        int helpContext;
        info.GetDocumentation(-1, out enumName, out doc, out helpContext, out helpFile);
        string prefix = enumName + "_";
        for (int v = 0; v < vars; v++)
        {
            IntPtr pv;
            info.GetVarDesc(v, out pv);
            var desc = (System.Runtime.InteropServices.ComTypes.VARDESC)Marshal.PtrToStructure(pv, typeof(System.Runtime.InteropServices.ComTypes.VARDESC));
            info.ReleaseVarDesc(pv);
            string name;
            info.GetDocumentation(desc.memid, out name, out doc, out helpContext, out helpFile);
            if (name.StartsWith(prefix)) ((ICreateTypeInfo2)info).SetVarName(v, name.Substring(prefix.Length));
        }
    }

    static int MemberId(System.Runtime.InteropServices.ComTypes.ITypeInfo info, int funcs, string name)
    {
        for (int f = 0; f < funcs; f++)
        {
            IntPtr pf;
            info.GetFuncDesc(f, out pf);
            var desc = (System.Runtime.InteropServices.ComTypes.FUNCDESC)Marshal.PtrToStructure(pf, typeof(System.Runtime.InteropServices.ComTypes.FUNCDESC));
            info.ReleaseFuncDesc(pf);
            string[] names = new string[1];
            int count;
            info.GetNames(desc.memid, names, 1, out count);
            if (names[0] == name) return desc.memid;
        }
        throw new InvalidOperationException("let_" + name + " sem a propriedade " + name);
    }

    public static void Register(string tlb)
    {
        System.Runtime.InteropServices.ComTypes.ITypeLib lib;
        LoadTypeLibEx(tlb, 2 /* REGKIND_NONE */, out lib);
        RegisterTypeLibForUser(lib, tlb, null);
    }

    public static void Unregister(Guid libId)
    {
        UnRegisterTypeLibForUser(ref libId, 1, 0, 0, 3 /* SYS_WIN64 */);
    }
}
'@

$LibId = [Guid]'3F8E2A61-7C4B-4E9D-A215-6B0C9D8E7F14'   # = [assembly: Guid] em AssemblyInfo.cs
$LibraryName = 'Sage'   # Dim s As Sage.StringS
$Classes = @(
    # Classe .NET, CLSID (= [Guid] da classe), ProgId (= [ProgId] da classe)
    @{ Class = 'SageTypes.StringS'; Clsid = '{9A4C1E7B-2D58-4F3A-B6C9-0E1F2A3B4C5D}'; ProgId = 'Sage.StringS' }
    @{ Class = 'SageTypes.DictionaryS'; Clsid = '{4B8E6D21-9F3C-4A57-B1D0-E5C27A8F6B39}'; ProgId = 'Sage.DictionaryS' }
    @{ Class = 'SageTypes.ListS'; Clsid = '{E7A42C19-6D3B-4E85-8F1A-0C9B5D2E7A64}'; ProgId = 'Sage.ListS' }
)

function Remove-Registration {
    [SageTypeLib]::Unregister($LibId)
    foreach ($c in $Classes) {
        foreach ($key in "HKCU:\Software\Classes\CLSID\$($c.Clsid)", "HKCU:\Software\Classes\$($c.ProgId)") {
            if (Test-Path $key) { Remove-Item $key -Recurse }
        }
    }
    # Primeira versão, com ProgId SageTypes.StringS
    if (Test-Path 'HKCU:\Software\Classes\SageTypes.StringS') { Remove-Item 'HKCU:\Software\Classes\SageTypes.StringS' -Recurse }
}

if ($Uninstall) {
    Remove-Registration
    if (Test-Path $Target) { Remove-Item $Target -Recurse -ErrorAction SilentlyContinue }
    'Removido.'
    return
}

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    Write-Warning 'O Excel está aberto: feche-o se a cópia do DLL falhar.'
}

& cmd /c "`"$PSScriptRoot\build.cmd`""
if ($LASTEXITCODE -ne 0) { throw 'A compilação falhou.' }

New-Item $Target -ItemType Directory -Force | Out-Null
try { Copy-Item (Join-Path $PSScriptRoot 'bin\SageTypes.dll') $Dll -Force }
catch { throw "Não foi possível copiar o DLL (o Excel está usando?). Feche o Excel e rode de novo." }

# Biblioteca de tipos: o que o VBA lista em Referências e usa no IntelliSense
Remove-Registration
[SageTypeLib]::Export($Dll, $Tlb, $LibraryName)
[SageTypeLib]::Register($Tlb)

# Classes COM (.NET via mscoree), como o SageEditor
$assemblyName = [Reflection.AssemblyName]::GetAssemblyName($Dll).FullName
foreach ($c in $Classes) {
    $key = "HKCU:\Software\Classes\CLSID\$($c.Clsid)"
    New-Item "$key\ProgId" -Force | Out-Null
    Set-Item "$key\ProgId" $c.ProgId
    Set-Item $key $c.Class
    $inproc = "$key\InprocServer32"
    New-Item $inproc -Force | Out-Null
    Set-Item $inproc 'mscoree.dll'
    Set-ItemProperty $inproc ThreadingModel 'Both'
    Set-ItemProperty $inproc Class $c.Class
    Set-ItemProperty $inproc Assembly $assemblyName
    Set-ItemProperty $inproc RuntimeVersion 'v4.0.30319'
    Set-ItemProperty $inproc CodeBase ('file:///' + ($Dll -replace '\\', '/'))
    New-Item "HKCU:\Software\Classes\$($c.ProgId)\CLSID" -Force | Out-Null
    Set-Item "HKCU:\Software\Classes\$($c.ProgId)\CLSID" $c.Clsid
}

"Instalado em $Dll"
"No VBA: Ferramentas > Referências > Sage. Depois: Dim s As Sage.StringS: Set s = New Sage.StringS"
