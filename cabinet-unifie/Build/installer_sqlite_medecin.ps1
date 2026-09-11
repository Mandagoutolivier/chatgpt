[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$SqliteExe,
    [string]$Destination = (Join-Path $env:APPDATA 'CabinetCardio\Tools\sqlite3.exe'),
    [string]$Sha256 = ''
)
. (Join-Path $PSScriptRoot 'outils_construction.ps1')
if (-not (Test-Path -LiteralPath $SqliteExe -PathType Leaf)) { throw "SQLite absent : $SqliteExe" }
if ($Sha256 -and (Get-FileHash -LiteralPath $SqliteExe -Algorithm SHA256).Hash -ne $Sha256) { throw 'Empreinte SQLite incorrecte.' }
# Tester AVANT tout remplacement de l'executable deja installe.
$version = & $SqliteExe -version
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace([string]$version)) { throw 'SQLite ne fonctionne pas sur ce PC.' }
$check = & $SqliteExe ':memory:' 'SELECT 42;'
if ($LASTEXITCODE -ne 0 -or [string]$check -ne '42') { throw 'Le test SQL a echoue.' }
if ([IO.Path]::GetFullPath($SqliteExe) -ne [IO.Path]::GetFullPath($Destination)) { Publier-FichierConstruit $SqliteExe $Destination }
Write-Host "SQLite pret : $version"
