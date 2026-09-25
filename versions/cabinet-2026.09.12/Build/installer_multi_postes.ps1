[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('Domicile','CabinetSecretariat','CabinetMedecin')][string]$Profil,
    [ValidateSet('Preparation','Installation')][string]$Mode = 'Preparation',
    [Parameter(Mandatory=$true)][string]$RacineNas,
    [string]$DossierGdt = 'C:\CabinetCardioTestU2\GDT',
    [string]$DossierSources = '',
    [string]$DossierPrepare = '',
    [string]$UrlService = '',
    [string]$FichierJeton = '',
    [switch]$RedemanderConnexion
)
. (Join-Path $PSScriptRoot 'outils_construction.ps1')
. (Join-Path $PSScriptRoot 'outils_installation.ps1')
if ($env:OS -ne 'Windows_NT') { throw 'Ce script necessite Windows avec Word et Excel installes.' }
$root = Split-Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($DossierSources)) {
    $DossierSources = Join-Path $root 'ModelesSource'
}
$medecin = $Profil -in @('Domicile','CabinetMedecin')
$secretariat = $Profil -in @('Domicile','CabinetSecretariat')
if ($RacineNas -notmatch '^\\\\[^\\]+\\[^\\]+') { throw 'Utilisez le chemin UNC du Synology pour RacineNas.' }
. (Join-Path $PSScriptRoot 'outils_assistant.ps1')
Attendre-FermetureOffice
if (-not (Test-Path -LiteralPath $RacineNas -PathType Container)) { throw "NAS inaccessible : $RacineNas. À domicile, connectez la liaison privée sécurisée du cabinet." }
$local = Join-Path $env:APPDATA 'CabinetCardio'
$identifiant = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8)
$stage = Join-Path $local ('Versions\' + $identifiant)
if ($Mode -eq 'Installation') {
    if ([string]::IsNullOrWhiteSpace($DossierPrepare)) { throw 'Indiquez -DossierPrepare : la version testee sera activee sans reconstruction.' }
    $stage=(Resolve-Path -LiteralPath $DossierPrepare).Path
    Verifier-Preparation $stage $Profil $root
    $receipt=Get-Content -LiteralPath (Join-Path $stage 'preparation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $receipt.PSObject.Properties['recetteValidee'] -or -not $receipt.recetteValidee) { throw 'Validez cette preparation avec Build/valider_preparation.ps1 apres la recette Windows.' }
}
$backupDir = Join-Path $local ('Sauvegardes\' + $identifiant)
[void][IO.Directory]::CreateDirectory($stage)
[void][IO.Directory]::CreateDirectory($backupDir)
$log = Join-Path $backupDir 'installation.log'
$changes = New-Object 'System.Collections.Generic.List[object]'
$lock = $null; $transcript = $false
$script:wordStartup = $null
$accesVbaJournal = Join-Path $backupDir 'acces-vba-a-restaurer.json'

function Memoriser-Fichier([string]$Destination) {
    $backup = $null; $sddl=$null
    if ([IO.File]::Exists($Destination)) {
        $backup = Join-Path $backupDir ([guid]::NewGuid().ToString('N') + '.bak')
        $sddl=(Get-Acl -LiteralPath $Destination).Sddl
        [IO.File]::Copy($Destination,$backup,$false)
        if ([IO.Path]::GetFileName($Destination) -eq 'service.token') { Proteger-FichierLocal $backup }
    }
    $changes.Add([pscustomobject]@{destination=$Destination;backup=$backup;sddl=$sddl})
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
function Installer-ConnexionService {
    $urlPath=Join-Path $local 'service.url';$tokenPath=Join-Path $local 'service.token'
    if ([string]::IsNullOrWhiteSpace($UrlService) -and (Test-Path -LiteralPath $urlPath)) { $script:UrlService=[IO.File]::ReadAllText($urlPath).Trim() }
    if ($RedemanderConnexion -and $UrlService) {
        $nouvelle=Read-Host ('Adresse HTTPS du service Synology [Entree : '+$UrlService+']')
        if ($nouvelle) { $script:UrlService=$nouvelle }
    }
    if ([string]::IsNullOrWhiteSpace($UrlService)) { $script:UrlService=Read-Host 'Adresse HTTPS du service Synology' }
    $uri=$null
    if (-not [Uri]::TryCreate($UrlService,[UriKind]::Absolute,[ref]$uri) -or $uri.Scheme -ne 'https' -or $uri.UserInfo) { throw 'Une adresse HTTPS sans identifiant dans l URL est requise.' }
    $token=''
    $reutiliser=(Test-Path -LiteralPath $tokenPath)
    if ($RedemanderConnexion -and $reutiliser -and -not $FichierJeton) {
        do { $choixJeton=Read-Host 'Conserver le jeton NAS deja enregistre ? OUI / NON' } until ($choixJeton -in @('OUI','NON'))
        $reutiliser=($choixJeton -eq 'OUI')
    }
    if ($FichierJeton) { $token=[IO.File]::ReadAllText((Resolve-Path -LiteralPath $FichierJeton).Path).Trim() }
    elseif ($reutiliser) { $token=[IO.File]::ReadAllText($tokenPath).Trim() }
    else {
        $secure=Read-Host 'Jeton de ce poste fourni lors de la creation du compte NAS' -AsSecureString
        $ptr=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { $token=[Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
    }
    if ($token.Length -lt 32 -or $token -match '[\r\n]') { throw 'Jeton de connexion invalide.' }
    $body=@{operation='whoami';params=@{}} | ConvertTo-Json -Compress
    $response=Invoke-RestMethod -Method Post -Uri ($UrlService.TrimEnd('/')+'/v1/rpc') -Headers @{Authorization=('Bearer '+$token)} -ContentType 'application/json' -Body $body -MaximumRedirection 0 -TimeoutSec 20
    if ($response.result.protocole -ne 2) { throw 'Service NAS incompatible avec cette version.' }
    if ($response.result.revision -ne '2026.09.21-u2c' -or $response.result.schema -ne 2) { throw 'Revision du service ou schema NAS incompatible avec la livraison U2 (migration ACTES requise).' }
    if ($medecin -and 'medecin' -notin $response.result.roles) { throw 'Ce compte ne possede pas le role medecin.' }
    if ($secretariat -and 'secretariat' -notin $response.result.roles) { throw 'Ce compte ne possede pas le role secretariat.' }
    # VBA Trim$ ne retire pas CR/LF : conserver une URL sans fin de ligne pour WinHTTP.
    Ecrire-Reglage $urlPath ($UrlService.TrimEnd('/'))
    Ecrire-Reglage $tokenPath $token
    Proteger-FichierLocal $tokenPath
    $token=$null
}
function Finaliser-ObjetsOfficeInstallation {
    # Les constructeurs s executent dans un scope enfant ; apres leur retour,
    # forcer la liberation de leurs proxies/collections COM devenus inaccessibles.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
function Tester-Office {
    foreach ($name in @('Word.Application','Excel.Application')) {
        $app = $null
        try {
            $app=New-Object -ComObject $name
            $app.AutomationSecurity=3
            if ($name -eq 'Word.Application') { $script:wordStartup=[string]$app.Options.DefaultFilePath(8) }
        } finally {
            if ($null -ne $app) { try { $app.Quit() } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } }
        }
    }
}
try {
    $lock = [IO.File]::Open((Join-Path $local 'installation.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    Start-Transcript -LiteralPath $log | Out-Null; $transcript=$true
    # La construction et elle seule a besoin de l acces programmatique au projet VBA.
    # L ancien reglage est journalise avant toute modification puis restaure dans finally.
    if ($Mode -eq 'Preparation') { Autoriser-AccesVbaAssistant $accesVbaJournal }
    # Les constructeurs utilisent Office ; Excel utilise Word pour la feuille de soins.
    Tester-Office
    # Les applications creees pour verifier Office doivent etre vraiment fermees avant les constructeurs.
    Attendre-FermetureOffice
    if ($Mode -eq 'Preparation' -and $medecin) {
        & (Join-Path $PSScriptRoot 'construire_modele_unifie.ps1') `
            -Prod6 (Join-Path $DossierSources 'ModeleCourrierChatGPT_PROD(6).dotm') `
            -Cabinet1 (Join-Path $DossierSources 'Cabinet(1).dotm') `
            -Sortie (Join-Path $stage 'CabinetUnifie.dotm') -RacineSources $root
        Finaliser-ObjetsOfficeInstallation
        Attendre-FermetureOffice -Noms 'WINWORD'
    }
    if ($Mode -eq 'Preparation' -and $secretariat) {
        & (Join-Path $PSScriptRoot 'construire_cabinet_secretariat.ps1') `
            -CabinetXlsm (Join-Path $DossierSources 'Cabinet.xlsm') -Sortie (Join-Path $stage 'Cabinet.xlsm') -RacineSources $root
        Finaliser-ObjetsOfficeInstallation
        Attendre-FermetureOffice -Noms 'EXCEL'
    }
    if ($Mode -eq 'Preparation') {
        Ecrire-Preparation $stage $Profil $root
        Write-Host "Preparation terminee : $stage"
        Write-Host 'Aucun complement actif remplace. Effectuez la recette Word/Excel, puis utilisez -Mode Installation.'
        Write-Output ([pscustomobject]@{DossierPrepare=$stage;Profil=$Profil;Mode='Preparation'})
        return
    }
    # Tous les binaires sont prets avant de modifier une installation active.
    # Le service NAS doit etre installe et le compte du poste cree au prealable.
    Installer-ConnexionService
    & (Join-Path $PSScriptRoot 'initialiser_nas.ps1') -RacineNas $RacineNas -RacineSources $root
    foreach ($required in @('Config\config.ini','Modeles\LETTRE TYPE.dot','Config\DDE\declencheurs_demandes.txt','Config\DDE\examens_complementaires.txt','Config\DDE\exclusions_demandes.txt')) {
        if (-not (Test-Path -LiteralPath (Join-Path $RacineNas $required))) { throw "Ressource NAS absente : $required. Installez le secretariat en premier." }
    }
    Ecrire-Reglage (Join-Path $local 'chemin.txt') ($RacineNas.TrimEnd('\') + "`r`n")
    $postePath=Join-Path $local 'poste.ini'
    $ancien='';if (Test-Path -LiteralPath $postePath) { $ancien=[IO.File]::ReadAllText($postePath,[Text.Encoding]::UTF8) }
    $reglages=[ordered]@{'poste|profil'=$Profil}
    if ($medecin) {
        if ($DossierGdt -notmatch '^[A-Za-z]:\\' -or $DossierGdt -match '[\r\n]') { throw 'DossierGdt doit etre un chemin local absolu.' }
        [void][IO.Directory]::CreateDirectory($DossierGdt)
        $reglages['ecg|dossiergdt']=$DossierGdt
    }
    Ecrire-Reglage $postePath (Fusionner-IniPoste $ancien $reglages)
    if ($medecin) {
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
    Installer-Fichier (Join-Path $stage 'preparation.json') (Join-Path $local ('installation-'+$Profil+'.json'))
    Write-Host "Installation terminee : $Profil. Sauvegardes et journal : $backupDir"
    Write-Host 'Dragon : affectez A/B/C/D aux quatre macros Unifie_*. Normal.dotm conserve ses macros historiques.'
} catch {
    $cause=$_
    $rollbackErrors=New-Object 'System.Collections.Generic.List[string]'
    for ($i=$changes.Count-1; $i -ge 0; $i--) {
        $item=$changes[$i]
        try {
            Restaurer-FichierAvecDroits $item
        } catch { $rollbackErrors.Add($item.destination) }
    }
    if ($rollbackErrors.Count) { Write-Warning ("Restauration locale incomplete : " + ($rollbackErrors -join ', ')) }
    Write-Warning "Installation interrompue. Les ressources NAS deja creees et les sauvegardes sont conservees. Journal : $log"
    throw $cause
} finally {
    # Ne jamais laisser AccessVBOM active par le lanceur, y compris apres une erreur de construction.
    if ($Mode -eq 'Preparation') {
        Attendre-FermetureOffice
        Restaurer-AccesVbaAssistant $accesVbaJournal
    }
    if ($transcript) { Stop-Transcript | Out-Null }
    if ($null -ne $lock) { $lock.Dispose() }
    # Les modeles construits sont toujours conserves pour la recette et le diagnostic.
}
