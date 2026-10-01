# Inicia o VBEShortcuts: compila src\VBEShortcuts.cs em memória e roda até
# receber --stop (ou o evento Local\VBEShortcuts.Stop), até o processo de
# --parent terminar ou até "Sair" no ícone da bandeja. Para iniciar com o
# Windows, rode install.ps1.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File VBEShortcuts.ps1 [--parent <pid>]
#
# Roda como script (e não como .exe) porque a política da máquina bloqueia
# executáveis sem assinatura.

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

$source = [IO.File]::ReadAllText((Join-Path $root 'src\VBEShortcuts.cs'), [Text.Encoding]::UTF8)
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Core', 'System.Drawing', 'System.Windows.Forms', 'Microsoft.CSharp')

$arguments = @($args) + @('--config', (Join-Path $root 'keybindings.txt'))
exit [VBEShortcuts.Program]::Main([string[]]$arguments)
