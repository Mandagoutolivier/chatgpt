[CmdletBinding()]
param(
    [ValidateSet('Domicile','CabinetSecretariat','CabinetMedecin')][string]$Profil='',
    [string]$RacineNas='',
    [string]$DossierGdt='C:\CabinetCardioTestU2\GDT'
)
. (Join-Path $PSScriptRoot 'outils_assistant.ps1')
. (Join-Path $PSScriptRoot 'outils_installation.ps1')
if ($env:OS -ne 'Windows_NT') { throw 'Ce lanceur necessite Windows, Word et Excel.' }
$root=Split-Path $PSScriptRoot -Parent
$local=Join-Path $env:APPDATA 'CabinetCardio\Assistant'
[void][IO.Directory]::CreateDirectory($local)
$lock=$null
try {
    $lock=[IO.File]::Open((Join-Path $local 'assistant.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    if (-not $Profil) { $Profil=Choisir-ProfilAssistant }
    $statePath=Join-Path $local ($Profil+'.json')
    $journal=Join-Path $local 'acces-vba-a-restaurer.json'
    $state=Lire-EtatAssistant $statePath
    $sourceHash=Empreinte-SourcesAssistant $root
    # Les chemins de source sont immuables apres telechargement ; un nouveau paquet commence une nouvelle preparation.
    if ($null -eq $state -or $state.racineSources -ne $root -or $state.manifeste -ne $sourceHash) {
        $nasConserve='';if ($null -ne $state) { $nasConserve=$state.racineNas }
        $state=[pscustomobject]@{profil=$Profil;racineSources=$root;manifeste=$sourceHash;racineNas=$nasConserve;dossierPrepare='';phase='nouveau';wordCompile=$false;excelCompile=$false;wordHash='';excelHash='';recette=$false;dossierGdt=$DossierGdt}
        Ecrire-EtatAssistant $statePath $state
    }
    if (Test-Path -LiteralPath $journal) { Restaurer-AccesVbaAssistant $journal;Attendre-FermetureOffice }
    if ($state.phase -eq 'installe') { Write-Host "Le profil $Profil a deja ete installe par ce lanceur. Dossier : $($state.dossierPrepare)";return }
    $ancienNas=$state.racineNas
    $RacineNas=Choisir-RacineNasAssistant $RacineNas $ancienNas (Join-Path (Split-Path $local -Parent) 'chemin.txt')
    if (-not $RacineNas) { Write-Host 'Installation en pause.';return }
    if ($ancienNas -and $ancienNas -ne $RacineNas) {
        # Les essais d une autre racine ne valident pas le nouvel environnement.
        $state.recette=$false
        if ($state.phase -eq 'valide') { $state.phase='compile' }
    }
    $state.racineNas=$RacineNas;Ecrire-EtatAssistant $statePath $state
    Attendre-FermetureOffice
    if ($state.phase -ne 'valide') {
    try {
        Autoriser-AccesVbaAssistant $journal
        if ($state.phase -eq 'nouveau') {
            $output=@(& (Join-Path $PSScriptRoot 'installer_multi_postes.ps1') -Profil $Profil -Mode Preparation -RacineNas $RacineNas -DossierGdt $state.dossierGdt)
            $prepared=@($output | Where-Object { $null -ne $_ -and $null -ne $_.PSObject.Properties['DossierPrepare'] })
            if ($prepared.Count -ne 1) { throw 'Le constructeur ne fournit pas un unique dossier de preparation.' }
            $state.dossierPrepare=$prepared[0].DossierPrepare;$state.phase='prepare';Ecrire-EtatAssistant $statePath $state
        }
        if (-not (Test-Path -LiteralPath $state.dossierPrepare -PathType Container)) { throw 'Dossier prepare introuvable. Conserver le journal pour diagnostic.' }
        Actualiser-CompilationAssistant $state
        Ecrire-EtatAssistant $statePath $state
        $medecin=$Profil -in @('Domicile','CabinetMedecin')
        $secretariat=$Profil -in @('Domicile','CabinetSecretariat')
        if (-not $state.wordCompile -and $medecin) {
            $state.wordCompile=Compiler-ProjetAssistant (Join-Path $state.dossierPrepare 'CabinetUnifie.dotm') 'Word'
            $state.wordHash=(Get-FileHash -LiteralPath (Join-Path $state.dossierPrepare 'CabinetUnifie.dotm')).Hash
            Ecrire-EtatAssistant $statePath $state
        }
        if (-not $state.excelCompile -and $secretariat) {
            $state.excelCompile=Compiler-ProjetAssistant (Join-Path $state.dossierPrepare 'Cabinet.xlsm') 'Excel'
            $state.excelHash=(Get-FileHash -LiteralPath (Join-Path $state.dossierPrepare 'Cabinet.xlsm')).Hash
            Ecrire-EtatAssistant $statePath $state
        }
        if ($state.excelCompile) { $state.excelHash=(Get-FileHash -LiteralPath (Join-Path $state.dossierPrepare 'Cabinet.xlsm')).Hash }
        $state.phase='compile';Ecrire-EtatAssistant $statePath $state
        Restaurer-AccesVbaAssistant $journal
        if (-not $state.recette) {
            $state.recette=Confirmer-RecetteAssistant (Join-Path $root 'RECETTE_WINDOWS.md') $state.dossierPrepare
            Ecrire-EtatAssistant $statePath $state
            if (-not $state.recette) { Write-Host 'Preparation conservee. Relancez ce meme lanceur et choisissez le meme profil pour reprendre.';return }
        }
        Attendre-FermetureOffice
        Autoriser-AccesVbaAssistant $journal
        # La recette peut enregistrer le classeur : reconfirmer la compilation du fichier final.
        if ($medecin -and (Get-FileHash -LiteralPath (Join-Path $state.dossierPrepare 'CabinetUnifie.dotm')).Hash -ne $state.wordHash) {
            $state.wordCompile=Compiler-ProjetAssistant (Join-Path $state.dossierPrepare 'CabinetUnifie.dotm') 'Word'
            $state.wordHash=(Get-FileHash -LiteralPath (Join-Path $state.dossierPrepare 'CabinetUnifie.dotm')).Hash
        }
        if ($secretariat -and (Get-FileHash -LiteralPath (Join-Path $state.dossierPrepare 'Cabinet.xlsm')).Hash -ne $state.excelHash) {
            $state.excelCompile=Compiler-ProjetAssistant (Join-Path $state.dossierPrepare 'Cabinet.xlsm') 'Excel'
            $state.excelHash=(Get-FileHash -LiteralPath (Join-Path $state.dossierPrepare 'Cabinet.xlsm')).Hash
        }
        Ecrire-EtatAssistant $statePath $state
        & (Join-Path $PSScriptRoot 'valider_preparation.ps1') -DossierPrepare $state.dossierPrepare -CompilationWordValidee:$state.wordCompile -CompilationExcelValidee:$state.excelCompile -RecetteValidee:$state.recette
        $state.phase='valide';Ecrire-EtatAssistant $statePath $state
    } finally {
        if (Test-Path -LiteralPath $journal) { Restaurer-AccesVbaAssistant $journal;Attendre-FermetureOffice }
    }
    } else {
        Verifier-Preparation $state.dossierPrepare $Profil $root
    }
    Write-Host ''
    Write-Host 'Activation : le service NAS doit etre demarre et avoir ses donnees initialisees.'
    Write-Host 'Vous utiliserez son adresse HTTPS et le jeton du compte de ce poste.'
    Write-Host 'Ce jeton est distinct du mot de passe DSM et de la cle OpenAI.'
    & (Join-Path $PSScriptRoot 'installer_multi_postes.ps1') -Profil $Profil -Mode Installation -RacineNas $RacineNas -DossierPrepare $state.dossierPrepare -DossierGdt $state.dossierGdt -RedemanderConnexion
    $state.phase='installe';Ecrire-EtatAssistant $statePath $state
    Write-Host "Installation terminee pour $Profil. Vous pouvez fermer cette fenetre."
} finally { if ($null -ne $lock) { $lock.Dispose() } }
