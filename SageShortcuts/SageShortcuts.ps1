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

# ToolboxFile.cs (a Caixa de ferramentas) é usado também pelo instalador: vem num arquivo à parte
$source = [IO.File]::ReadAllText((Join-Path $root 'src\SageShortcuts.cs'), [Text.Encoding]::UTF8) + "`n" +
    [IO.File]::ReadAllText((Join-Path $root 'src\ToolboxFile.cs'), [Text.Encoding]::UTF8)
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @(
    'System.Core', 'System.Drawing', 'System.Windows.Forms', 'Microsoft.CSharp')

# O atalho de uma linha "Ctrl+K, Ctrl+C = comment  # ..." ("ctrl+k,ctrl+c"), ou nada
function Get-BindingKey([string]$line) {
    $text = ($line -replace '#.*$', '').Trim()
    if ($text -match '^(.+?)\s*=') { return ($Matches[1] -replace '\s+', '').ToLowerInvariant() }
    return $null
}

# Atalhos novos do padrão (de uma versão nova) entram no fim do arquivo do usuário, uma vez cada.
# Os já oferecidos ficam em keybindings.offered.txt, então um atalho que o usuário apagou não
# volta; um atalho que o usuário já usa para outra ação fica de fora.
function Add-NewBindings([string]$default, [string]$user) {
    $offeredPath = Join-Path (Split-Path $user) 'keybindings.offered.txt'
    $offered = New-Object 'System.Collections.Generic.HashSet[string]'
    if (Test-Path $offeredPath) { foreach ($k in [IO.File]::ReadAllLines($offeredPath)) { [void]$offered.Add($k) } }
    $taken = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($line in [IO.File]::ReadAllLines($user)) { $k = Get-BindingKey $line; if ($k) { [void]$taken.Add($k) } }
    $added = @()
    foreach ($line in [IO.File]::ReadAllLines($default)) {
        $k = Get-BindingKey $line
        if (-not $k -or -not $offered.Add($k)) { continue }
        if (-not $taken.Contains($k)) { $added += $line }
    }
    if ($added.Count -gt 0) {
        [IO.File]::AppendAllText($user, "`r`n" + ($added -join "`r`n") + "`r`n", (New-Object Text.UTF8Encoding $false))
    }
    [IO.File]::WriteAllLines($offeredPath, [string[]]@($offered))
}

# Instalado em Arquivos de Programas (só leitura para o usuário): os atalhos ficam em
# %APPDATA%\Sage\keybindings.txt, copiado do padrão na primeira vez e depois completado com os
# atalhos novos. Num clone do git, o próprio keybindings.txt da pasta.
$config = Join-Path $root 'keybindings.txt'
if ($root.StartsWith($env:ProgramFiles, [StringComparison]::OrdinalIgnoreCase)) {
    $userConfig = Join-Path $env:APPDATA 'Sage\keybindings.txt'
    if (-not (Test-Path $userConfig)) {
        New-Item (Split-Path $userConfig) -ItemType Directory -Force | Out-Null
        Copy-Item $config $userConfig
    }
    try { Add-NewBindings $config $userConfig }
    catch { } # sem os atalhos novos, mas os do usuário continuam valendo
    $config = $userConfig
}

$arguments = @($args) + @('--config', $config)
exit [SageShortcuts.Program]::Main([string[]]$arguments)
