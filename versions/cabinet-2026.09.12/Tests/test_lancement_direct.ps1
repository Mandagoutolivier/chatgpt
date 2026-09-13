$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$executable = (Get-Process -Id $PID).Path
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('cabinet lancement direct ' + [guid]::NewGuid().ToString('N'))
$script:total = 0

# De vrais processus -File depuis un autre dossier : pas de simulation du binding.
# Chaque cas doit atteindre une validation connue AVANT tout acces NAS ou Office.
function Verifier-Lancement([string]$Nom, [string]$Script, [string[]]$Arguments, [string]$Attendu) {
    $stdout = Join-Path $tmp 'stdout.txt'
    $stderr = Join-Path $tmp 'stderr.txt'
    $fichier = Join-Path (Join-Path $root 'Build') $Script
    $liste = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $fichier) + $Arguments
    $ligne = ($liste | ForEach-Object { '"' + $_ + '"' }) -join ' '
    $processus = Start-Process -FilePath $executable -ArgumentList $ligne -WorkingDirectory $tmp `
        -RedirectStandardOutput $stdout -RedirectStandardError $stderr -Wait -PassThru
    try {
        $message = [IO.File]::ReadAllText($stdout) + [IO.File]::ReadAllText($stderr)
        if ($processus.ExitCode -eq 0 -or -not $message.Contains($Attendu)) {
            throw "ECHEC $Nom : sortie $($processus.ExitCode) ; $message"
        }
        if ($message.Contains('ParameterArgumentValidationErrorEmptyStringNotAllowed')) {
            throw "ECHEC $Nom : un chemin vide a ete evalue pendant le binding."
        }
        $script:total++
        Write-Output "PASS : $Nom"
    } finally {
        $processus.Dispose()
    }
}

try {
    [void][IO.Directory]::CreateDirectory($tmp)
    $absent = Join-Path $tmp 'modele-absent.dotm'
    $sortie = Join-Path $tmp 'sortie-interdite.dotm'
    $explicite = Join-Path $tmp 'sources explicites'
    [void][IO.Directory]::CreateDirectory($explicite)

    Verifier-Lancement 'NAS -File sans RacineSources' 'initialiser_nas.ps1' `
        @('-RacineNas', 'pas-un-chemin-unc') 'La racine doit etre un chemin UNC'
    Verifier-Lancement 'NAS -File avec RacineSources explicite' 'initialiser_nas.ps1' `
        @('-RacineNas', 'pas-un-chemin-unc', '-RacineSources', $root) 'La racine doit etre un chemin UNC'
    Verifier-Lancement 'Word -File trouve le manifeste adjacent' 'construire_modele_unifie.ps1' `
        @('-Prod6', $absent, '-Cabinet1', $absent, '-Sortie', $sortie) 'Modele absent :'
    Verifier-Lancement 'Excel -File trouve le manifeste adjacent' 'construire_cabinet_secretariat.ps1' `
        @('-CabinetXlsm', $absent, '-Sortie', $sortie) 'Modele absent :'
    Verifier-Lancement 'Excel -File conserve la racine explicite' 'construire_cabinet_secretariat.ps1' `
        @('-CabinetXlsm', $absent, '-Sortie', $sortie, '-RacineSources', $explicite) 'sources explicites'

    $attendu = if ($env:OS -eq 'Windows_NT') { 'Utilisez le chemin UNC du Synology' } else { 'Ce script necessite Windows' }
    Verifier-Lancement 'Installation -File sans DossierSources' 'installer_multi_postes.ps1' `
        @('-Profil', 'Domicile', '-RacineNas', 'pas-un-chemin-unc') $attendu
    if (Test-Path -LiteralPath $sortie) { throw 'Un test a produit un fichier Office.' }
    Write-Output "$script:total controles de lancement direct reussis."
} finally {
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force }
}
