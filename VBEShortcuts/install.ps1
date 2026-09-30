# Instala o VBEShortcuts para iniciar com o Windows e já o inicia agora.
#   powershell -ExecutionPolicy Bypass -File install.ps1              instala
#   powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall   remove e encerra

param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$script = Join-Path $PSScriptRoot 'VBEShortcuts.ps1'
$link = Join-Path ([Environment]::GetFolderPath('Startup')) 'VBEShortcuts.lnk'

function Stop-VBEShortcuts {
    try {
        $ev = [Threading.EventWaitHandle]::OpenExisting('Local\VBEShortcuts.Stop')
        [void]$ev.Set()
        $ev.Dispose()
    } catch [Threading.WaitHandleCannotBeOpenedException] { }
}

if ($Uninstall) {
    if (Test-Path $link) { Remove-Item $link }
    Stop-VBEShortcuts
    "Removido: $link"
    return
}

$arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($link)
$shortcut.TargetPath = Join-Path $PSHOME 'powershell.exe'
$shortcut.Arguments = $arguments
$shortcut.WorkingDirectory = $PSScriptRoot
$shortcut.WindowStyle = 7   # minimizado, para não piscar janela
$shortcut.Description = 'Atalhos de teclado do editor VBA'
$shortcut.Save()
"Instalado: $link"

Stop-VBEShortcuts   # reinicia, caso uma versão anterior esteja rodando
Start-Sleep -Milliseconds 500
Start-Process $shortcut.TargetPath -ArgumentList $arguments -WindowStyle Hidden
"Iniciado (ícone na bandeja)."
