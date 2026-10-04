# Inicia o SageShortcuts: compila src\SageShortcuts.cs em memória e roda até
# receber --stop (ou o evento Local\SageShortcuts.Stop), até o processo de
# --parent terminar ou até "Sair" no ícone da bandeja. Para iniciar com o
# Windows, rode install.ps1.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File SageShortcuts.ps1 [--parent <pid>]
#
# Roda como script (e não como .exe) porque a política da máquina bloqueia
# executáveis sem assinatura.

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

$source = [IO.File]::ReadAllText((Join-Path $root 'src\SageShortcuts.cs'), [Text.Encoding]::UTF8)
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Core', 'System.Drawing', 'System.Windows.Forms', 'Microsoft.CSharp')

# Instalado em Arquivos de Programas (só leitura para o usuário): os atalhos ficam em
# %APPDATA%\Sage\keybindings.txt, copiado do padrão na primeira vez. Num clone do git,
# o próprio keybindings.txt da pasta.
$config = Join-Path $root 'keybindings.txt'
if ($root.StartsWith($env:ProgramFiles, [StringComparison]::OrdinalIgnoreCase)) {
    $userConfig = Join-Path $env:APPDATA 'Sage\keybindings.txt'
    if (-not (Test-Path $userConfig)) {
        New-Item (Split-Path $userConfig) -ItemType Directory -Force | Out-Null
        Copy-Item $config $userConfig
    }
    $config = $userConfig
}

$arguments = @($args) + @('--config', $config)
exit [SageShortcuts.Program]::Main([string[]]$arguments)
