# Instala o add-in Sage no editor do VBA (só para o usuário atual, sem admin).
#   powershell -ExecutionPolicy Bypass -File install.ps1              compila, copia e registra
#   powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall   remove
#   powershell -ExecutionPolicy Bypass -File install.ps1 -ExportTo <pasta>
#       só compila e copia para <pasta> (pacote da release; build-release.ps1), com com.txt
#       (classes COM que o instalador registra)
# O Excel carrega o add-in ao abrir; feche-o antes de instalar ou atualizar.

param([switch]$Uninstall, [string]$ExportTo)

$ErrorActionPreference = 'Stop'
$ProgId = 'Sage.Editor'
$Clsid = '{3C1D5E7A-9B2F-4A6C-8E41-7F0A2B9D6C53}'
$Target = Join-Path $env:LOCALAPPDATA 'Sage\Editor'
$Dll = Join-Path $Target 'SageEditor.dll'
$AddinKey = "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\$ProgId"
$ClassKey = "HKCU:\Software\Classes\CLSID\$Clsid"
# Controle da janela Terminal (o VBE o hospeda na janela acoplável)
$HostProgId = 'Sage.TerminalHost'
$HostClsid = '{31E81384-D124-4B77-A222-BEE1ACED5A35}'
$HostClassKey = "HKCU:\Software\Classes\CLSID\$HostClsid"

if ($ExportTo) {
    & cmd /c "`"$PSScriptRoot\build.cmd`""
    if ($LASTEXITCODE -ne 0) { throw 'A compilação falhou.' }
    New-Item $ExportTo -ItemType Directory -Force | Out-Null
    Copy-Item (Join-Path $PSScriptRoot 'bin\SageEditor.dll') $ExportTo -Force
    # dll|classe|CLSID|ProgId
    Set-Content (Join-Path $ExportTo 'com.txt') @("SageEditor.dll|SageEditor.Connect|$Clsid|$ProgId",
        "SageEditor.dll|SageEditor.TerminalHost|$HostClsid|$HostProgId") -Encoding UTF8
    return
}

# Versão antiga, quando o add-in se chamava SageVBE (ProgId Sage.VBE)
foreach ($key in 'HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\Sage.VBE', 'HKCU:\Software\Classes\Sage.VBE') {
    if (Test-Path $key) { Remove-Item $key -Recurse }
}
$legacyTarget = Join-Path $env:LOCALAPPDATA 'Sage\VBE'
if (Test-Path $legacyTarget) { Remove-Item $legacyTarget -Recurse -ErrorAction SilentlyContinue }

if ($Uninstall) {
    foreach ($key in $AddinKey, $ClassKey, "HKCU:\Software\Classes\$ProgId", $HostClassKey, "HKCU:\Software\Classes\$HostProgId") {
        if (Test-Path $key) { Remove-Item $key -Recurse }
    }
    if (Test-Path $Target) { Remove-Item $Target -Recurse -ErrorAction SilentlyContinue }
    "Removido. As configurações ficam em $env:APPDATA\Sage."
    return
}

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    Write-Warning 'O Excel está aberto: a instalação vale a partir da próxima vez que ele abrir.'
}

& cmd /c "`"$PSScriptRoot\build.cmd`""
if ($LASTEXITCODE -ne 0) { throw 'A compilação falhou.' }

New-Item $Target -ItemType Directory -Force | Out-Null
try { Copy-Item (Join-Path $PSScriptRoot 'bin\SageEditor.dll') $Dll -Force }
catch { throw "Não foi possível copiar o DLL (o Excel está usando?). Feche o Excel e rode de novo." }

# Classes COM (.NET via mscoree): o add-in e o controle da janela Terminal
function Register-Class($progId, $clsid, $class) {
    $classKey = "HKCU:\Software\Classes\CLSID\$clsid"
    New-Item "HKCU:\Software\Classes\$progId\CLSID" -Force | Out-Null
    Set-Item "HKCU:\Software\Classes\$progId\CLSID" $clsid
    New-Item "$classKey\ProgId" -Force | Out-Null
    Set-Item "$classKey\ProgId" $progId
    $inproc = "$classKey\InprocServer32"
    New-Item $inproc -Force | Out-Null
    Set-Item $inproc 'mscoree.dll'
    Set-ItemProperty $inproc ThreadingModel 'Both'
    Set-ItemProperty $inproc Class $class
    Set-ItemProperty $inproc Assembly 'SageEditor, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null'
    Set-ItemProperty $inproc RuntimeVersion 'v4.0.30319'
    Set-ItemProperty $inproc CodeBase ('file:///' + ($Dll -replace '\\', '/'))
}
Register-Class $ProgId $Clsid 'SageEditor.Connect'
Register-Class $HostProgId $HostClsid 'SageEditor.TerminalHost'

# Registro como suplemento do VBE (3 = carregar ao iniciar)
New-Item $AddinKey -Force | Out-Null
New-ItemProperty $AddinKey LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
Set-ItemProperty $AddinKey FriendlyName 'Sage'
Set-ItemProperty $AddinKey Description 'Menu Sage e temas de cores para o editor do VBA'

"Instalado em $Dll"
"Abra o Excel, vá ao editor do VBA (Alt+F11) e use o menu Sage > Configurações."
