# Instala o add-in Sage no editor do VBA (só para o usuário atual, sem admin).
#   powershell -ExecutionPolicy Bypass -File install.ps1              compila, copia e registra
#   powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall   remove
# O Excel carrega o add-in ao abrir; feche-o antes de instalar ou atualizar.

param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$ProgId = 'Sage.Editor'
$Clsid = '{3C1D5E7A-9B2F-4A6C-8E41-7F0A2B9D6C53}'
$Target = Join-Path $env:LOCALAPPDATA 'Sage\Editor'
$Dll = Join-Path $Target 'SageEditor.dll'
$AddinKey = "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\$ProgId"
$ClassKey = "HKCU:\Software\Classes\CLSID\$Clsid"

# Versão antiga, quando o add-in se chamava SageVBE (ProgId Sage.VBE)
foreach ($key in 'HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\Sage.VBE', 'HKCU:\Software\Classes\Sage.VBE') {
    if (Test-Path $key) { Remove-Item $key -Recurse }
}
$legacyTarget = Join-Path $env:LOCALAPPDATA 'Sage\VBE'
if (Test-Path $legacyTarget) { Remove-Item $legacyTarget -Recurse -ErrorAction SilentlyContinue }

if ($Uninstall) {
    foreach ($key in $AddinKey, $ClassKey, "HKCU:\Software\Classes\$ProgId") {
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

# Classe COM (.NET via mscoree)
New-Item "HKCU:\Software\Classes\$ProgId\CLSID" -Force | Out-Null
Set-Item "HKCU:\Software\Classes\$ProgId\CLSID" $Clsid
New-Item "$ClassKey\ProgId" -Force | Out-Null
Set-Item "$ClassKey\ProgId" $ProgId
$inproc = "$ClassKey\InprocServer32"
New-Item $inproc -Force | Out-Null
Set-Item $inproc 'mscoree.dll'
Set-ItemProperty $inproc ThreadingModel 'Both'
Set-ItemProperty $inproc Class 'SageEditor.Connect'
Set-ItemProperty $inproc Assembly 'SageEditor, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null'
Set-ItemProperty $inproc RuntimeVersion 'v4.0.30319'
Set-ItemProperty $inproc CodeBase ('file:///' + ($Dll -replace '\\', '/'))

# Registro como suplemento do VBE (3 = carregar ao iniciar)
New-Item $AddinKey -Force | Out-Null
New-ItemProperty $AddinKey LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
Set-ItemProperty $AddinKey FriendlyName 'Sage'
Set-ItemProperty $AddinKey Description 'Menu Sage e temas de cores para o editor do VBA'

"Instalado em $Dll"
"Abra o Excel, vá ao editor do VBA (Alt+F11) e use o menu Sage > Configurações."
