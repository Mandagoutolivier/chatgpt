param([Parameter(Mandatory=$true)][string]$SqliteExe)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $SqliteExe)) { throw "sqlite3.exe introuvable : $SqliteExe" }
$cible = Join-Path $env:APPDATA 'CabinetCardio\Tools'
New-Item -ItemType Directory -Force -Path $cible | Out-Null
Copy-Item $SqliteExe (Join-Path $cible 'sqlite3.exe') -Force
& (Join-Path $cible 'sqlite3.exe') -version
if ($LASTEXITCODE -ne 0) { throw 'sqlite3.exe ne demarre pas sur ce poste.' }
Write-Host "SQLite installe : $cible\sqlite3.exe" -ForegroundColor Green
