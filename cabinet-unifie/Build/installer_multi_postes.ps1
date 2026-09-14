[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('Domicile','CabinetSecretariat','CabinetMedecin')][string]$Profil,
    [ValidateSet('Preparation','Installation')][string]$Mode = 'Preparation',
    [string]$RacineNas = '\\DS224\CabinetCardio',
    [string]$DossierGdt = 'C:\Mandagout',
    [string]$DossierSources = (Join-Path (Split-Path $PSScriptRoot -Parent) 'ModelesSource'),
    [string]$SqliteExe = '',
    [string]$SqliteSha256 = '',
    [switch]$ConserverModelesConstruits
)
. (Join-Path $PSScriptRoot 'outils_construction.ps1')
throw 'INSTALLATEUR ARCHIVE ET DESACTIVE : utilisez uniquement la version service NAS apres publication et recette du nouveau lanceur.'
if ($env:OS -ne 'Windows_NT') { throw 'Ce script necessite Windows avec Word et Excel installes.' }
$root = Split-Path $PSScriptRoot -Parent
$medecin = $Profil -in @('Domicile','CabinetMedecin')
$secretariat = $Profil -in @('Domicile','CabinetSecretariat')
if ($RacineNas -notmatch '^\\\\[^\\]+\\[^\\]+') { throw 'Utilisez le chemin UNC du Synology pour RacineNas.' }
if (Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue) { throw 'Fermez completement Word et Excel.' }
if (-not (Test-Path -LiteralPath $RacineNas -PathType Container)) { throw "NAS inaccessible : $RacineNas. A domicile, connectez le VPN Cabinet Freebox Pro." }
$local = Join-Path $env:APPDATA 'CabinetCardio'
$identifiant = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8)
$stage = Join-Path $local ('Versions\' + $identifiant)
$backupDir = Join-Path $local ('Sauvegardes\' + $identifiant)
[void][IO.Directory]::CreateDirectory($stage)
[void][IO.Directory]::CreateDirectory($backupDir)
$log = Join-Path $backupDir 'installation.log'
$changes = New-Object 'System.Collections.Generic.List[object]'
$lock = $null; $transcript = $false
$script:wordStartup = $null

function Memoriser-Fichier([string]$Destination) {
    $backup = $null
    if ([IO.File]::Exists($Destination)) {
        $backup = Join-Path $backupDir ([guid]::NewGuid().ToString('N') + '.bak')
        [IO.File]::Copy($Destination,$backup,$false)
    }
    $changes.Add([pscustomobject]@{destination=$Destination;backup=$backup})
    $changes | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $backupDir 'restauration.json') -Encoding UTF8
}
function Installer-Fichier([string]$Source,[string]$Destination) {
    Memoriser-Fichier $Destination
    [void][IO.Directory]::CreateDirectory((Split-Path $Destination -Parent))
    [IO.File]::Copy($Source,$Destination,$true)
    if ((Get-FileHash -LiteralPath $Source).Hash -ne (Get-FileHash -LiteralPath $Destination).Hash) { throw "Verification de copie echouee : $Destination" }
}
function Ecrire-Reglage([string]$Destination,[string]$Texte) {
    Memoriser-Fichier $Destination
    [IO.File]::WriteAllText($Destination,$Texte,(New-Object Text.UTF8Encoding($false)))
}
function Obtenir-Sqlite {
    if (-not [Environment]::Is64BitOperatingSystem) { throw 'Le paquet SQLite fourni est x64. Fournissez un sqlite3.exe compatible avec ce Windows.' }
    $dep = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'sqlite.lock.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $archive = Join-Path $stage 'sqlite-tools.zip'
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -UseBasicParsing -Uri $dep.url -OutFile $archive
    if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant() -ne $dep.sha256) { throw 'Archive SQLite differente de la version verifiee.' }
    $folder = Join-Path $stage 'sqlite-download'
    Expand-Archive -LiteralPath $archive -DestinationPath $folder
    $files = @(Get-ChildItem -LiteralPath $folder -Recurse -Filter sqlite3.exe)
    if ($files.Count -ne 1) { throw 'Archive SQLite inattendue.' }
    return $files[0].FullName
}
function Tester-Office {
    foreach ($name in @('Word.Application','Excel.Application')) {
        $app = $null
        try {
            $app=New-Object -ComObject $name
            $app.AutomationSecurity=3
            if ($name -eq 'Word.Application') { $script:wordStartup=[string]$app.Options.DefaultFilePath(8) }
            $app.Quit()
        } finally {
            if ($null -ne $app) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
        }
    }
}
try {
    $lock = [IO.File]::Open((Join-Path $local 'installation.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    Start-Transcript -LiteralPath $log | Out-Null; $transcript=$true
    # Les deux applications sont requises dans tous les profils : Word lit les
    # bases avec Excel ; Excel utilise Word pour la feuille de soins.
    Tester-Office
    if ($medecin) {
        if ([string]::IsNullOrWhiteSpace($SqliteExe)) { $SqliteExe=Obtenir-Sqlite }
        & (Join-Path $PSScriptRoot 'installer_sqlite_medecin.ps1') -SqliteExe $SqliteExe -Destination (Join-Path $stage 'sqlite3.exe') -Sha256 $SqliteSha256
        & (Join-Path $PSScriptRoot 'construire_modele_unifie.ps1') `
            -Prod6 (Join-Path $DossierSources 'ModeleCourrierChatGPT_PROD(6).dotm') `
            -Cabinet1 (Join-Path $DossierSources 'Cabinet(1).dotm') `
            -Sortie (Join-Path $stage 'CabinetUnifie.dotm') -RacineSources $root
    }
    if ($secretariat) {
        & (Join-Path $PSScriptRoot 'construire_cabinet_secretariat.ps1') `
            -CabinetXlsm (Join-Path $DossierSources 'Cabinet.xlsm') -Sortie (Join-Path $stage 'Cabinet.xlsm') -RacineSources $root
    }
    if ($Mode -eq 'Preparation') {
        Write-Host "Preparation terminee : $stage"
        Write-Host 'Aucun complement actif remplace. Effectuez la recette Word/Excel, puis utilisez -Mode Installation.'
        return
    }
    # Tous les binaires sont prets avant de modifier une installation active.
    if ($secretariat) { & (Join-Path $PSScriptRoot 'initialiser_nas.ps1') -RacineNas $RacineNas -RacineSources $root }
    foreach ($required in @('Base\Patients.xlsx','Config\config.ini','Config\Gras_Medicaments.xlsx','Config\Gras_Expressions.xlsx','Echange\Arrives','Echange\AEnvoyer','Modeles\LETTRE TYPE.dot','Base\base_travail_correspondants_v1.xlsx','Config\DDE\declencheurs_demandes.txt','Config\DDE\examens_complementaires.txt','Config\DDE\exclusions_demandes.txt')) {
        if (-not (Test-Path -LiteralPath (Join-Path $RacineNas $required))) { throw "Ressource NAS absente : $required. Installez le secretariat en premier." }
    }
    Ecrire-Reglage (Join-Path $local 'chemin.txt') ($RacineNas.TrimEnd('\') + "`r`n")
    $posteIni = "[POSTE]`r`nProfil=$Profil`r`n"
    if ($medecin) {
        if ($DossierGdt -notmatch '^[A-Za-z]:\\' -or $DossierGdt -match '[\r\n]') { throw 'DossierGdt doit etre un chemin local absolu, par exemple C:\Mandagout.' }
        [void][IO.Directory]::CreateDirectory($DossierGdt)
        $posteIni += "[ECG]`r`nDossierGdt=$DossierGdt`r`n"
    }
    # Conserver les autres reglages du poste (imprimante, etc.) lors d'une mise a jour.
    $postePath=Join-Path $local 'poste.ini'
    if (Test-Path -LiteralPath $postePath) {
        $ancien=[IO.File]::ReadAllText($postePath,[Text.Encoding]::UTF8)
        $posteIni = $ancien + "`r`n" + $posteIni
    }
    Ecrire-Reglage $postePath $posteIni
    if ($medecin) {
        Installer-Fichier (Join-Path $stage 'sqlite3.exe') (Join-Path $local 'Tools\sqlite3.exe')
        $startup=$script:wordStartup
        if ([string]::IsNullOrWhiteSpace($startup)) { throw 'Word ne fournit pas de dossier de demarrage.' }
        [void][IO.Directory]::CreateDirectory($startup)
        foreach ($old in Get-ChildItem -LiteralPath $startup -File -Filter '*.dotm') {
            if ($old.Name -eq 'Cabinet.dotm' -or $old.Name -like 'ModeleCourrierChatGPT*.dotm') {
                Memoriser-Fichier $old.FullName
                [IO.File]::Delete($old.FullName)
            }
        }
        Installer-Fichier (Join-Path $stage 'CabinetUnifie.dotm') (Join-Path $startup 'CabinetUnifie.dotm')
    }
    if ($secretariat) {
        $destination=Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CabinetCardio\Cabinet.xlsm'
        Installer-Fichier (Join-Path $stage 'Cabinet.xlsm') $destination
        $shortcut=Join-Path ([Environment]::GetFolderPath('Desktop')) 'Cabinet Cardio.lnk'
        Memoriser-Fichier $shortcut
        $shell=New-Object -ComObject WScript.Shell
        $link=$shell.CreateShortcut($shortcut)
        $link.TargetPath=$destination; $link.WorkingDirectory=Split-Path $destination -Parent; $link.Save()
    }
    Write-Host "Installation terminee : $Profil. Sauvegardes et journal : $backupDir"
    Write-Host 'Dragon : affectez A/B/C/D aux quatre macros Unifie_*. Normal.dotm conserve ses macros historiques.'
} catch {
    $cause=$_
    $rollbackErrors=New-Object 'System.Collections.Generic.List[string]'
    for ($i=$changes.Count-1; $i -ge 0; $i--) {
        $item=$changes[$i]
        try {
            if ($item.backup) { [IO.File]::Copy($item.backup,$item.destination,$true) }
            elseif ([IO.File]::Exists($item.destination)) { [IO.File]::Delete($item.destination) }
        } catch { $rollbackErrors.Add($item.destination) }
    }
    if ($rollbackErrors.Count) { Write-Warning ("Restauration locale incomplete : " + ($rollbackErrors -join ', ')) }
    Write-Warning "Installation interrompue. Les ressources NAS deja creees et les sauvegardes sont conservees. Journal : $log"
    throw $cause
} finally {
    if ($transcript) { Stop-Transcript | Out-Null }
    if ($null -ne $lock) { $lock.Dispose() }
    # Les modeles construits sont toujours conserves pour la recette et le diagnostic.
}
