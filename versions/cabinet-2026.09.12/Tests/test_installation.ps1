$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'Build/outils_installation.ps1')
$tmp=Join-Path ([IO.Path]::GetTempPath()) ('cabinet-install-test-'+[guid]::NewGuid().ToString('N'))
$root=Join-Path $tmp 'sources';$stage=Join-Path $tmp 'preparation'
$count=0
function Verifier([bool]$Resultat,[string]$Nom) { if (-not $Resultat) { throw "ECHEC $Nom" };$script:count++;Write-Output "PASS : $Nom" }
try {
    [void][IO.Directory]::CreateDirectory((Join-Path $root 'Src'))
    [void][IO.Directory]::CreateDirectory((Join-Path $root 'Build'))
    [void][IO.Directory]::CreateDirectory($stage)
    $source=Join-Path $root 'Src/test.bas';[IO.File]::WriteAllText($source,'source fictive')
    foreach ($name in @('CabinetUnifie.dotm','Cabinet.xlsm')) { [IO.File]::WriteAllText((Join-Path $stage $name),'binaire fictif') }
    Ecrire-Preparation $stage 'Domicile' $root
    Verifier-Preparation $stage 'Domicile' $root
    Verifier $true 'empreintes et profil initiaux acceptes'
    $refuse=$false;try { Verifier-Preparation $stage 'CabinetMedecin' $root } catch { $refuse=$true };Verifier $refuse 'autre profil refuse'
    [IO.File]::WriteAllText($source,'modifiee')
    $refuse=$false;try { Verifier-Preparation $stage 'Domicile' $root } catch { $refuse=$true };Verifier $refuse 'modification source refusee'
    [IO.File]::WriteAllText($source,'source fictive')
    [IO.File]::WriteAllText((Join-Path $stage 'Cabinet.xlsm'),'autre binaire')
    $refuse=$false;try { Verifier-Preparation $stage 'Domicile' $root } catch { $refuse=$true };Verifier $refuse 'modification binaire refusee'
    [IO.File]::WriteAllText((Join-Path $stage 'Cabinet.xlsm'),'binaire fictif')
    Remove-Item -LiteralPath (Join-Path $stage 'CabinetUnifie.dotm')
    $refuse=$false;try { Verifier-Preparation $stage 'Domicile' $root } catch { $refuse=$true };Verifier $refuse 'modele manquant refusee'
    Write-Output "$count controles de preparation reussis."
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force }
