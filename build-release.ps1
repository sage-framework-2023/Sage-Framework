# Monta o instalador da release do GitHub, em um arquivo só:
#   dist\SageFramework-Setup-<versão>.exe   (e o .sha256 ao lado, para conferir o download)
#
# Compila o SageEditor e o SageTypes (com a biblioteca de tipos), junta o SageShortcuts, a
# duckdb.dll e a documentação num zip embutido no instalador (SageSetup\src\Setup.cs).
# A duckdb.dll vem de SageTypes\lib (baixada e conferida pelo SageTypes\install.ps1).
# Usa só o csc do .NET Framework 4, que já vem no Windows.
#
#   powershell -ExecutionPolicy Bypass -File build-release.ps1 -Version 1.0.0
param([string]$Version = (Get-Date -Format 'yyyy.M.d'))

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$build = Join-Path $root 'dist\build'
$stage = Join-Path $build 'payload'
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    Write-Warning 'O Excel está aberto: se ele estiver usando os DLLs em bin\, a compilação pode falhar.'
}
if (Test-Path $build) { Remove-Item $build -Recurse -Force }
New-Item $stage -ItemType Directory -Force | Out-Null

# Cada componente compila e exporta numa sessão própria do PowerShell (o SageTypes carrega o
# DLL para gerar o .tlb, e o arquivo ficaria preso nesta sessão)
function Export-Component([string]$Script, [string]$Target) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root $Script) -ExportTo $Target
    if ($LASTEXITCODE -ne 0) { throw "Falhou: $Script -ExportTo $Target" }
}
Write-Host 'Compilando o SageEditor...'
Export-Component 'SageEditor\install.ps1' (Join-Path $stage 'Editor')
Write-Host 'Compilando o SageTypes e gerando a biblioteca de tipos...'
Export-Component 'SageTypes\install.ps1' (Join-Path $stage 'Types')

# SageShortcuts roda como script (compila o .cs ao iniciar)
$shortcuts = Join-Path $stage 'Shortcuts'
New-Item (Join-Path $shortcuts 'src') -ItemType Directory -Force | Out-Null
Copy-Item (Join-Path $root 'SageShortcuts\SageShortcuts.ps1'), (Join-Path $root 'SageShortcuts\keybindings.txt') $shortcuts
Copy-Item (Join-Path $root 'SageShortcuts\src\SageShortcuts.cs') (Join-Path $shortcuts 'src')

# Documentação (atalho "Documentação" no menu Iniciar) e a licença do DuckDB
$docs = Join-Path $stage 'Docs'
New-Item $docs -ItemType Directory -Force | Out-Null
Copy-Item (Join-Path $root 'SageEditor\README.md') (Join-Path $docs 'SageEditor.md')
Copy-Item (Join-Path $root 'SageTypes\README.md') (Join-Path $docs 'SageTypes.md')
Copy-Item (Join-Path $root 'SageShortcuts\keybindings.txt') (Join-Path $docs 'SageShortcuts - atalhos.txt')
Copy-Item (Join-Path $root 'SageTypes\THIRD-PARTY-NOTICES.txt') $docs
Copy-Item (Join-Path $root 'SageTypes\THIRD-PARTY-NOTICES.txt') (Join-Path $stage 'Types')
# Ícones do vscode-icons no SageEditor (MIT)
Copy-Item (Join-Path $root 'SageEditor\THIRD-PARTY-NOTICES.txt') (Join-Path $docs 'SageEditor - THIRD-PARTY-NOTICES.txt')
Copy-Item (Join-Path $root 'SageEditor\THIRD-PARTY-NOTICES.txt') (Join-Path $stage 'Editor')

# Zip embutido
Add-Type -AssemblyName System.IO.Compression.FileSystem
$payload = Join-Path $build 'payload.zip'
[IO.Compression.ZipFile]::CreateFromDirectory($stage, $payload, 'Optimal', $false)

# Versão no instalador: o texto como veio; a do assembly só com números (1.0.0 -> 1.0.0.0)
$numbers = @(([regex]::Matches($Version, '\d+') | Select-Object -First 4 | ForEach-Object { [int]$_.Value }))
while ($numbers.Count -lt 4) { $numbers += 0 }
$assemblyVersion = $numbers -join '.'
$info = Join-Path $build 'BuildInfo.cs'
@"
using System.Reflection;
[assembly: AssemblyTitle("Sage Framework Setup")]
[assembly: AssemblyProduct("Sage Framework")]
[assembly: AssemblyDescription("Instalador do Sage Framework")]
[assembly: AssemblyVersion("$assemblyVersion")]
[assembly: AssemblyFileVersion("$assemblyVersion")]
[assembly: AssemblyInformationalVersion("$Version")]
namespace SageSetup { static class BuildInfo { public const string Version = "$Version"; } }
"@ | Set-Content $info -Encoding UTF8

$dist = Join-Path $root 'dist'
$exe = Join-Path $dist "SageFramework-Setup-$Version.exe"
if (Test-Path $exe) { Remove-Item $exe }
Write-Host 'Compilando o instalador...'
& $csc -nologo -target:winexe -platform:anycpu -optimize+ -codepage:65001 `
    "-win32icon:$(Join-Path $root 'SageInstaller\sage.ico')" `
    "-win32manifest:$(Join-Path $root 'SageSetup\app.manifest')" `
    "-resource:$payload,payload.zip" `
    -r:System.dll -r:System.Core.dll -r:System.Drawing.dll -r:System.Windows.Forms.dll `
    -r:System.IO.Compression.dll -r:Microsoft.CSharp.dll `
    "-out:$exe" (Join-Path $root 'SageSetup\src\Setup.cs') $info
if ($LASTEXITCODE -ne 0) { throw 'A compilação do instalador falhou.' }

$hash = (Get-FileHash $exe -Algorithm SHA256).Hash
"$hash  $(Split-Path $exe -Leaf)" | Set-Content "$exe.sha256" -Encoding ASCII
Remove-Item $build -Recurse -Force

"{0} ({1:N1} MB)" -f $exe, ((Get-Item $exe).Length / 1MB)
"SHA256: $hash"
