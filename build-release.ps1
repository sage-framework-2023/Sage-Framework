# Monta o pacote da release do GitHub: dist\Sage-Framework-<versão>.zip
#
# Leva os arquivos do repositório (os do git e os novos ainda não commitados, menos os
# ignorados) e a duckdb.dll em SageTypes\lib, que não vai para o git. A dll precisa estar
# lá e conferir com o hash do SageTypes\install.ps1 (rode o install.ps1 uma vez: ele baixa).
#
#   powershell -ExecutionPolicy Bypass -File build-release.ps1 -Version 1.0.0
param([string]$Version = (Get-Date -Format 'yyyy.MM.dd'))

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

# Fora do pacote: o atalho .lnk (aponta para caminhos desta máquina), o suplemento legado
# e pastas de trabalho de teste, arquivos temporários do Office
$exclude = @('*.lnk', 'Sage.xlam', 'Pasta1.xlsm', '~$*', 'build-release.ps1')

# duckdb.dll conferida com o mesmo hash que o instalador exige
$install = Get-Content (Join-Path $root 'SageTypes\install.ps1') -Raw
if ($install -notmatch "DllHash\s*=\s*'([0-9A-F]{64})'") { throw 'Hash da duckdb.dll não encontrado no SageTypes\install.ps1.' }
$dllHash = $Matches[1]
$dll = Join-Path $root 'SageTypes\lib\duckdb.dll'
if (-not (Test-Path $dll)) { throw "Falta $dll. Rode SageTypes\install.ps1 uma vez (ele baixa a dll)." }
if ((Get-FileHash $dll -Algorithm SHA256).Hash -ne $dllHash) { throw "$dll não confere com o hash esperado." }

Push-Location $root
try {
    $files = @(git -c core.quotepath=off ls-files --cached --others --exclude-standard | Where-Object {
        $name = Split-Path $_ -Leaf
        -not ($exclude | Where-Object { $name -like $_ })
    } | Where-Object { Test-Path -LiteralPath $_ })
    if ($LASTEXITCODE -ne 0) { throw 'git ls-files falhou.' }
    $dirty = git status --porcelain
    if ($dirty) { Write-Warning 'Há alterações não commitadas; elas vão no pacote assim como estão.' }
}
finally { Pop-Location }
$files += 'SageTypes/lib/duckdb.dll'

$dist = Join-Path $root 'dist'
New-Item $dist -ItemType Directory -Force | Out-Null
$zip = Join-Path $dist "Sage-Framework-$Version.zip"
if (Test-Path $zip) { Remove-Item $zip }

Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$archive = [IO.Compression.ZipFile]::Open($zip, 'Create')
try {
    foreach ($file in $files) {
        $entry = 'Sage-Framework/' + ($file -replace '\\', '/')
        [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, (Join-Path $root $file), $entry, 'Optimal') | Out-Null
    }
}
finally { $archive.Dispose() }

"{0} arquivos em {1} ({2:N1} MB)" -f $files.Count, $zip, ((Get-Item $zip).Length / 1MB)
