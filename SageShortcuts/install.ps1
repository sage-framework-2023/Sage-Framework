# Instala o SageShortcuts para iniciar com o Windows e já o inicia agora.
#   powershell -ExecutionPolicy Bypass -File install.ps1              instala
#   powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall   remove e encerra

param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$script = Join-Path $PSScriptRoot 'SageShortcuts.ps1'
$link = Join-Path ([Environment]::GetFolderPath('Startup')) 'SageShortcuts.lnk'

function Stop-SageShortcuts([string]$Name = 'SageShortcuts') {
    try {
        $ev = [Threading.EventWaitHandle]::OpenExisting("Local\$Name.Stop")
        [void]$ev.Set()
        $ev.Dispose()
    } catch [Threading.WaitHandleCannotBeOpenedException] { }
}

# Versão antiga, quando o componente se chamava VBEShortcuts
$legacyLink = Join-Path ([Environment]::GetFolderPath('Startup')) 'VBEShortcuts.lnk'
if (Test-Path $legacyLink) { Remove-Item $legacyLink }
Stop-SageShortcuts 'VBEShortcuts'

if ($Uninstall) {
    if (Test-Path $link) { Remove-Item $link }
    Stop-SageShortcuts
    "Removido: $link"
    return
}

# Pelo conhost sem janela: com o Windows Terminal como terminal padrão (Windows 11), o
# powershell.exe abriria numa aba dele, e o -WindowStyle Hidden não a esconde
$powershell = Join-Path $PSHOME 'powershell.exe'
$arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($link)
$shortcut.TargetPath = Join-Path $env:WINDIR 'System32\conhost.exe'
$shortcut.Arguments = "--headless `"$powershell`" $arguments"
$shortcut.WorkingDirectory = $PSScriptRoot
$shortcut.WindowStyle = 7   # minimizado, para não piscar janela
$shortcut.Description = 'Atalhos de teclado do editor VBA'
$shortcut.Save()
"Instalado: $link"

Stop-SageShortcuts   # reinicia, caso uma versão anterior esteja rodando
Start-Sleep -Milliseconds 500
Start-Process $shortcut.TargetPath -ArgumentList $shortcut.Arguments -WindowStyle Hidden
"Iniciado (ícone na bandeja)."
