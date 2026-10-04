# Instala os tipos do Sage para o VBA (StringS...) só para o usuário atual, sem admin.
#   powershell -ExecutionPolicy Bypass -File install.ps1              compila, copia e registra
#   powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall   remove
#   powershell -ExecutionPolicy Bypass -File install.ps1 -ExportTo <pasta>
#       só compila e gera em <pasta> o DLL, a duckdb.dll, o .tlb e o com.txt (classes COM
#       que o instalador registra), sem registrar nada (pacote da release; build-release.ps1)
# Depois, em cada pasta de trabalho: Ferramentas > Referências > "Sage".
# O Excel carrega o DLL quando o VBA usa um tipo; feche-o antes de instalar ou atualizar.

param([switch]$Uninstall, [string]$ExportTo)

$ErrorActionPreference = 'Stop'
$Target = Join-Path $env:LOCALAPPDATA 'Sage\Types'
$Dll = Join-Path $Target 'SageTypes.dll'
$Tlb = Join-Path $Target 'SageTypes.tlb'

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

public static class SageTypeLib
{
    [ComImport, Guid("00020406-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICreateTypeLib
    {
        void CreateTypeInfo();
        void SetName([MarshalAs(UnmanagedType.LPWStr)] string name);
        void SetVersion(); void SetGuid(); void SetDocString([MarshalAs(UnmanagedType.LPWStr)] string doc);
        void SetHelpFileName(); void SetHelpContext(); void SetLcid(); void SetLibFlags();
        void SaveAllChanges();
    }

    // ICreateTypeInfo2: só os métodos usados têm assinatura; os demais guardam a posição na vtable
    [ComImport, Guid("0002040E-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICreateTypeInfo2
    {
        void SetGuid(); void SetTypeFlags(int flags); void SetDocString(); void SetHelpContext(); void SetVersion();
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
        // Nome em Ferramentas > Referências: a descrição do assembly com a versão ("Sage Framework 1.0")
        Version version = assembly.GetName().Version;
        AssemblyDescriptionAttribute description = (AssemblyDescriptionAttribute)Attribute.GetCustomAttribute(assembly, typeof(AssemblyDescriptionAttribute));
        lib.SetDocString((description != null ? description.Description : libraryName) + " " + version.Major + "." + version.Minor);
        ApplyProperties((System.Runtime.InteropServices.ComTypes.ITypeLib)lib, PropertyMethods(assembly), AppObjects(assembly));
        lib.SaveAllChanges();
    }

    // Classes marcadas com [AppObject]: os membros ficam globais no VBA (Json.Loads sem Dim)
    static Dictionary<string, bool> AppObjects(Assembly assembly)
    {
        var result = new Dictionary<string, bool>();
        foreach (Type type in assembly.GetTypes())
            foreach (CustomAttributeData data in CustomAttributeData.GetCustomAttributes(type))
                if (data.Constructor.DeclaringType.Name == "AppObjectAttribute") result[type.Name] = true;
        return result;
    }

    // Métodos das interfaces marcados com [PropertyGet("X")] ou [PropertyLet("X")] no
    // código: "Interface.Método" -> (é Let, nome da propriedade)
    static Dictionary<string, KeyValuePair<bool, string>> PropertyMethods(Assembly assembly)
    {
        var result = new Dictionary<string, KeyValuePair<bool, string>>();
        foreach (Type type in assembly.GetTypes())
        {
            if (!type.IsInterface) continue;
            foreach (MethodInfo method in type.GetMethods())
                foreach (CustomAttributeData data in CustomAttributeData.GetCustomAttributes(method))
                {
                    string kind = data.Constructor.DeclaringType.Name;
                    if (kind != "PropertyGetAttribute" && kind != "PropertyLetAttribute") continue;
                    result[type.Name + "." + method.Name] = new KeyValuePair<bool, string>(
                        kind == "PropertyLetAttribute", (string)data.ConstructorArguments[0].Value);
                }
        }
        return result;
    }

    // O .NET não exporta propriedades com parâmetros além do indexador, e exporta o
    // setter de uma propriedade Variant só como Property Set (propputref), com o que
    // "d(chave) = valor" não compila no VBA. Por isso:
    // - [PropertyGet("X")] (GetAt): o método vira a leitura da propriedade X, com
    //   parâmetros, como At(linha, coluna);
    // - [PropertyLet("X")] (LetValue): o método vira o Property Let (propput) de X:
    //   mesmo DispId de X, mesma posição na vtable (onde o .NET o implementou),
    //   parâmetros iguais.
    static void ApplyProperties(System.Runtime.InteropServices.ComTypes.ITypeLib lib, Dictionary<string, KeyValuePair<bool, string>> methods,
        Dictionary<string, bool> appObjects)
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
            if (attr.typekind == System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_COCLASS)
            {
                string name, doc, helpFile;
                int helpContext;
                info.GetDocumentation(-1, out name, out doc, out helpContext, out helpFile);
                if (appObjects.ContainsKey(name))
                    ((ICreateTypeInfo2)info).SetTypeFlags((int)attr.wTypeFlags | 0x1 /* TYPEFLAG_FAPPOBJECT */);
                continue;
            }
            if (attr.typekind != System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_DISPATCH ||
                (attr.wTypeFlags & System.Runtime.InteropServices.ComTypes.TYPEFLAGS.TYPEFLAG_FDUAL) == 0) continue;
            // A interface da vtable por trás da visão IDispatch
            int href;
            info.GetRefTypeOfImplType(-1, out href);
            System.Runtime.InteropServices.ComTypes.ITypeInfo vtable;
            info.GetRefTypeInfo(href, out vtable);
            ApplyProperties(vtable, methods);
        }
    }

    static void ApplyProperties(System.Runtime.InteropServices.ComTypes.ITypeInfo info, Dictionary<string, KeyValuePair<bool, string>> methods)
    {
        string typeName, doc, helpFile;
        int helpContext;
        info.GetDocumentation(-1, out typeName, out doc, out helpContext, out helpFile);
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
            KeyValuePair<bool, string> mark;
            if (!methods.TryGetValue(typeName + "." + names[0], out mark) ||
                desc.invkind != System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_FUNC)
            {
                info.ReleaseFuncDesc(pf);
                continue;
            }

            string property = mark.Value;
            if (!mark.Key)
            {
                desc.invkind = System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYGET;
                IntPtr getCopy = Marshal.AllocHGlobal(Marshal.SizeOf(desc));
                try
                {
                    Marshal.StructureToPtr(desc, getCopy, false);
                    ICreateTypeInfo2 create = (ICreateTypeInfo2)info;
                    create.DeleteFuncDesc(f);
                    create.AddFuncDesc(f, getCopy);
                    names[0] = property;
                    create.SetFuncAndParamNames(f, names, count);
                }
                finally
                {
                    Marshal.FreeHGlobal(getCopy);
                    info.ReleaseFuncDesc(pf);
                }
                continue;
            }

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

    // O .NET exporta os membros de um enum como "Enum_Membro" (SgArrayTypes_SgTuple);
    // no VBA eles ficam só com o nome do membro (SgTuple)
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
        throw new InvalidOperationException("[PropertyLet(\"" + name + "\")] sem a propriedade " + name);
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
    @{ Class = 'SageTypes.DataFrame'; Clsid = '{2F9C6E14-8B3A-4D71-A5E2-7C0D9B4F3A68}'; ProgId = 'Sage.DataFrame' }
    @{ Class = 'SageTypes.DateTimeS'; Clsid = '{992489F3-7D1F-4D2F-B13A-05AB53F1A745}'; ProgId = 'Sage.DateTimeS' }
    @{ Class = 'SageTypes.Json'; Clsid = '{F255D235-A52E-4362-9484-F344705B9DE7}'; ProgId = 'Sage.Json' }
    @{ Class = 'SageTypes.Requests'; Clsid = '{18A710C0-DD90-407E-92A8-56E2E1FBB35C}'; ProgId = 'Sage.Requests' }
    @{ Class = 'SageTypes.Session'; Clsid = '{FCB13B55-C345-4703-89B6-F4635B7C5A9A}'; ProgId = 'Sage.Session' }
    @{ Class = 'SageTypes.Response'; Clsid = '{B028ADB0-40B2-461C-8AD1-B4236852C17E}'; ProgId = 'Sage.Response' }
    # Membros globais (Json, Requests): o VBA cria esta classe sozinho
    @{ Class = 'SageTypes.Globals'; Clsid = '{3129D3A3-5535-4294-8923-083DAA4A6F75}'; ProgId = 'Sage.Globals' }
)

# DuckDB (motor do DataFrame). No pacote da release vem em lib\; num clone do git é
# baixado do GitHub do DuckDB, nesta versão exata, conferindo o hash.
$DuckDB = @{
    Version = '1.5.6'
    Url     = 'https://github.com/duckdb/duckdb/releases/download/v1.5.6/libduckdb-windows-amd64.zip'
    ZipHash = '44CF59583F9951D2CB09B1BF115A63ECB2D8901E363903029D86C7D8683FE96A'
    DllHash = '7E90BDEF028D57B45490D6F53530E3C012A5D9ABDC4C6723449200663DC84ED0'
}

function Get-DuckDB {
    $lib = Join-Path $PSScriptRoot 'lib'
    $dll = Join-Path $lib 'duckdb.dll'
    if ((Test-Path $dll) -and (Get-FileHash $dll -Algorithm SHA256).Hash -eq $DuckDB.DllHash) { return $dll }
    Write-Host "Baixando o DuckDB $($DuckDB.Version)..."
    New-Item $lib -ItemType Directory -Force | Out-Null
    $zip = Join-Path ([IO.Path]::GetTempPath()) "libduckdb-$($DuckDB.Version).zip"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest $DuckDB.Url -OutFile $zip -UseBasicParsing
    if ((Get-FileHash $zip -Algorithm SHA256).Hash -ne $DuckDB.ZipHash) { Remove-Item $zip; throw 'O arquivo do DuckDB baixado não confere com o hash esperado.' }
    $unzip = Join-Path ([IO.Path]::GetTempPath()) "libduckdb-$($DuckDB.Version)"
    if (Test-Path $unzip) { Remove-Item $unzip -Recurse -Force }
    Expand-Archive $zip $unzip
    Copy-Item (Join-Path $unzip 'duckdb.dll') $dll -Force
    Remove-Item $zip, $unzip -Recurse -Force
    if ((Get-FileHash $dll -Algorithm SHA256).Hash -ne $DuckDB.DllHash) { throw 'A duckdb.dll não confere com o hash esperado.' }
    return $dll
}

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

if ($ExportTo) {
    & cmd /c "`"$PSScriptRoot\build.cmd`""
    if ($LASTEXITCODE -ne 0) { throw 'A compilação falhou.' }
    $duckdbDll = Get-DuckDB
    New-Item $ExportTo -ItemType Directory -Force | Out-Null
    $exportDll = Join-Path $ExportTo 'SageTypes.dll'
    Copy-Item (Join-Path $PSScriptRoot 'bin\SageTypes.dll') $exportDll -Force
    Copy-Item $duckdbDll (Join-Path $ExportTo 'duckdb.dll') -Force
    [SageTypeLib]::Export($exportDll, (Join-Path $ExportTo 'SageTypes.tlb'), $LibraryName)
    # dll|classe|CLSID|ProgId
    $lines = foreach ($c in $Classes) { "SageTypes.dll|$($c.Class)|$($c.Clsid)|$($c.ProgId)" }
    Set-Content (Join-Path $ExportTo 'com.txt') $lines -Encoding UTF8
    return
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

$duckdbDll = Get-DuckDB
New-Item $Target -ItemType Directory -Force | Out-Null
try
{
    Copy-Item (Join-Path $PSScriptRoot 'bin\SageTypes.dll') $Dll -Force
    $installedDuckDB = Join-Path $Target 'duckdb.dll'
    if (-not (Test-Path $installedDuckDB) -or (Get-FileHash $installedDuckDB -Algorithm SHA256).Hash -ne $DuckDB.DllHash) {
        Copy-Item $duckdbDll $installedDuckDB -Force
    }
}
catch { throw "Não foi possível copiar os DLLs (o Excel está usando?). Feche o Excel e rode de novo." }

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
