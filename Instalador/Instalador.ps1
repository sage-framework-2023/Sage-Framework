# Abre o Sage Framework Installer (janela no estilo do Visual Studio Installer).
# Roda como script, e não como .exe, porque a política da máquina bloqueia
# executáveis sem assinatura. Use o atalho "Sage Framework Installer" para abrir com duplo clique.

$ErrorActionPreference = 'Stop'
$source = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'SageInstaller.cs'), [Text.Encoding]::UTF8)
Add-Type -TypeDefinition $source -Language CSharp -ReferencedAssemblies @('System.Drawing', 'System.Windows.Forms')
[SageInstaller.Program]::Run((Split-Path $PSScriptRoot -Parent))