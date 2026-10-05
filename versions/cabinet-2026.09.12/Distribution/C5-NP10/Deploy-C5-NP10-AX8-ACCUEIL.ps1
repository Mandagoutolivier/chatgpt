#requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$Execute,
    [string]$ConfirmerBascule = '',
    [string]$RevisionServiceAttendue = '',
    [string]$AncienComplementWord = '',
    [string]$DossierPaquet = '\\DS224\CabinetCardio-Dev\Deploy\C5-NP10',
    [string]$BureauAccueil = 'C:\Users\accueil\Desktop',
    [string]$Restaurer = ''
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$HashWord = 'FF183F42D591BE91CEE2146AF2260439157957C3A5D4FD13D9328B31F19C2C7C'
$HashExcel = '454670F1CB00DD2929EA09D9E6AA0243E68C43E6478B279979E836AFB1FECD24'
$AccueilAdmin = '\\ACCUEIL\C$'

function Assert-Hash([string]$Path, [string]$Expected) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Fichier absent : $Path" }
    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if ($actual -ine $Expected) { throw "Empreinte incorrecte : $Path" }
}
function Remote-Path([string]$LocalPath) {
    if ($LocalPath -notmatch '^C:\\' -or $LocalPath -match '\.\.') { throw 'Chemin ACCUEIL invalide.' }
    return Join-Path $AccueilAdmin $LocalPath.Substring(3)
}
function Assert-OfficeFerme {
    if (@(Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue).Count) {
        throw 'Fermer Word et Excel sur AX8_MAX, dans toutes les sessions.'
    }
    # Lecture DCOM bornee. Aucun reglage WinRM/pare-feu n'est modifie.
    $job = Start-Job -ScriptBlock {
        $ErrorActionPreference = 'Stop'
        $computer = Get-WmiObject Win32_ComputerSystem -ComputerName ACCUEIL
        if ($computer.Name -ine 'ACCUEIL') { throw 'Le serveur atteint n est pas ACCUEIL.' }
        @(Get-WmiObject Win32_Process -ComputerName ACCUEIL -Filter "Name='EXCEL.EXE' OR Name='WINWORD.EXE'").Count
    }
    try {
        if (-not (Wait-Job $job -Timeout 20)) { throw 'Controle Office ACCUEIL expire (20 s). Verifier droits WMI/DCOM.' }
        $count = @(Receive-Job $job -ErrorAction Stop)
        if ($job.State -ne 'Completed' -or $count.Count -ne 1) { throw 'Controle Office ACCUEIL incomplet.' }
        if ([int]$count[0] -gt 0) { throw 'Fermer Word et Excel sur ACCUEIL, dans toutes les sessions.' }
    } finally { Stop-Job $job -ErrorAction SilentlyContinue; Remove-Job $job -Force -ErrorAction SilentlyContinue }
}
function Write-Json([string]$Path, $Object) {
    $temp = $Path + '.tmp'
    [IO.File]::WriteAllText($temp, ($Object | ConvertTo-Json -Depth 10), (New-Object Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temp -Destination $Path -Force
}
function New-Shortcut([string]$Path, [string]$Target, [string]$Arguments, [string]$WorkingDirectory) {
    if (Test-Path -LiteralPath $Path) { throw "Raccourci deja present : $Path" }
    $shell = New-Object -ComObject WScript.Shell
    $link = $null
    try {
        $link = $shell.CreateShortcut($Path)
        $link.TargetPath = $Target; $link.Arguments = $Arguments; $link.WorkingDirectory = $WorkingDirectory
        $link.Save()
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'Creation du raccourci echouee.' }
    } finally {
        if ($link) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($link) }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell)
    }
}
function Restore-Deployment($Journal, [string]$JournalPath) {
    # Ne pas effacer de fichiers ; masquer les nouveaux raccourcis et restaurer
    # l'ancien complement. Les copies C5 et les sauvegardes sont conservees.
    $restoreErrors = New-Object 'System.Collections.Generic.List[string]'
    if (Test-Path -LiteralPath $Journal.oldDisabled -PathType Leaf) {
        Assert-Hash $Journal.oldDisabled $Journal.oldHash
        if (Test-Path -LiteralPath $Journal.oldOriginal) { throw 'Retour arriere : ancien chemin deja occupe, aucune copie ecrasee.' }
        Move-Item -LiteralPath $Journal.oldDisabled -Destination $Journal.oldOriginal
        Assert-Hash $Journal.oldOriginal $Journal.oldHash
    } elseif (-not (Test-Path -LiteralPath $Journal.oldOriginal -PathType Leaf)) {
        Assert-Hash $Journal.oldBackup $Journal.oldHash
        Copy-Item -LiteralPath $Journal.oldBackup -Destination $Journal.oldOriginal
        Assert-Hash $Journal.oldOriginal $Journal.oldHash
    } else { Assert-Hash $Journal.oldOriginal $Journal.oldHash }
    # Un partage ACCUEIL devenu inaccessible ne doit pas empecher de retablir Word.
    foreach ($path in @($Journal.shortcuts)) {
        try {
            if ($path -like '\\ACCUEIL\C$\*' -and -not (Test-Path -LiteralPath '\\ACCUEIL\C$' -PathType Container -ErrorAction Stop)) {
                throw 'Partage ACCUEIL indisponible pour neutraliser le raccourci.'
            }
            if (Test-Path -LiteralPath $path -PathType Leaf -ErrorAction Stop) {
                $disabled = $path + '.inactif-' + [guid]::NewGuid().ToString('N')
                Move-Item -LiteralPath $path -Destination $disabled
            }
        } catch { $restoreErrors.Add('Raccourci non neutralise : '+$path) }
    }
    if ($restoreErrors.Count) {
        $Journal['state']='RESTAURATION_INCOMPLETE';$Journal['rollbackError']=($restoreErrors -join '; ')
    } else { $Journal['state'] = 'RESTAURE' }
    Write-Json $JournalPath $Journal
    if ($restoreErrors.Count) { throw ($restoreErrors -join '; ') }
}

if ($env:OS -ne 'Windows_NT') { throw 'Windows PowerShell 5.1 est requis sur AX8_MAX.' }
if ([Security.Principal.WindowsIdentity]::GetCurrent().Name -ine 'AX8_MAX\olivi') {
    throw 'Executer ce script sous AX8_MAX\olivi, pas dans CabinetU2Medecin ni RDC.'
}
if ($Restaurer) {
    if (@(Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Fermer Word et Excel sur AX8_MAX avant le retour arriere.' }
    $obj = Get-Content -LiteralPath $Restaurer -Raw -Encoding UTF8 | ConvertFrom-Json
    $journal = [ordered]@{}; foreach ($prop in $obj.PSObject.Properties) { $journal[$prop.Name] = $prop.Value }
    if ($journal['kind'] -ne 'CabinetCardio-C5-AX8-ACCUEIL-v2') { throw 'Journal de restauration incompatible.' }
    if (-not $Execute) { Write-Host 'Controle retour arriere : fermer les nouvelles applications ; les donnees NAS ne seront pas restaurees.'; $obj; return }
    Restore-Deployment $journal $Restaurer
    Write-Host 'Ancien complement restaure ; nouveaux raccourcis rendus inactifs.'
    return
}

if (-not (Test-Path -LiteralPath $AccueilAdmin -PathType Container)) { throw 'Partage administratif \\ACCUEIL\C$ inaccessible.' }

$sourceWord = Join-Path $PSScriptRoot 'CabinetDragonC5NP10.dotm'
$sourceExcel = Join-Path $PSScriptRoot 'Cabinet.xlsm'
$sourceLaunch = Join-Path $PSScriptRoot 'Lancer-Cabinet-C5NP10.ps1'
$sourceCommon = Join-Path $PSScriptRoot 'Controle-Service-C5.ps1'
Assert-Hash $sourceWord $HashWord; Assert-Hash $sourceExcel $HashExcel
foreach ($path in @($sourceLaunch,$sourceCommon)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Paquet incomplet : $path" }
    $tokens = $null; $errors = $null
    [void][Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw "Syntaxe invalide : $path" }
}
. $sourceCommon
$hashLaunch = (Get-FileHash -LiteralPath $sourceLaunch -Algorithm SHA256).Hash
$hashCommon = (Get-FileHash -LiteralPath $sourceCommon -Algorithm SHA256).Hash
$configAX8 = Join-Path $env:APPDATA 'CabinetCardio'
$configAccueil = '\\ACCUEIL\C$\Users\accueil\AppData\Roaming\CabinetCardio'
$doctor = Test-C5Service $configAX8 'medecin' $RevisionServiceAttendue
$secretary = Test-C5Service $configAccueil 'secretariat' $RevisionServiceAttendue
if ($doctor.Url -ine $secretary.Url) { throw 'Les deux postes ne ciblent pas le meme service.' }
if ($doctor.Revision -cne $secretary.Revision) { throw 'Les deux postes ne voient pas la meme revision de service.' }
if ($doctor.ID -eq $secretary.ID) { throw 'Les deux postes doivent avoir des comptes applicatifs distincts.' }
Assert-OfficeFerme
foreach ($name in @('Config\config.ini','Modeles\LETTRE TYPE.dot')) {
    if (-not (Test-Path -LiteralPath (Join-Path '\\DS224\CabinetCardio' $name) -PathType Leaf)) {
        throw "Ressource partagee de production absente : $name"
    }
}
# Le controle des clients ne peut pas prouver le montage clinique des conteneurs NAS.
if (-not $AncienComplementWord) {
    $startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
    $candidates = @(Get-ChildItem -LiteralPath $startup -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^(Cabinet|ModeleCourrierChatGPT).*\.dot[m]?$' })
    if ($candidates.Count -ne 1) { throw 'Preciser -AncienComplementWord : complement actuel absent ou plusieurs candidats.' }
    $AncienComplementWord = $candidates[0].FullName
}
$old = Get-Item -LiteralPath $AncienComplementWord -ErrorAction Stop
if ($old.PSIsContainer -or $old.FullName -notlike (Join-Path $env:APPDATA '*') -or $old.Name -notmatch '^(Cabinet|ModeleCourrierChatGPT).*\.dot[m]?$') {
    throw 'Ancien complement : fichier Cabinet .dot/.dotm dans le profil APPDATA attendu.'
}
$oldHash = (Get-FileHash -LiteralPath $old.FullName -Algorithm SHA256).Hash
$desktop = [Environment]::GetFolderPath('Desktop')
$remoteDesktop = Remote-Path $BureauAccueil
if ($DossierPaquet -notlike '\\DS224\CabinetCardio-Dev\*' -or $DossierPaquet -match '\.\.') {
    throw 'DossierPaquet doit rester sous \\DS224\CabinetCardio-Dev.'
}
foreach ($path in @($desktop,$remoteDesktop,'\\DS224\CabinetCardio-Dev')) {
    if (-not (Test-Path -LiteralPath $path -PathType Container)) { throw "Dossier requis inaccessible : $path" }
}

Write-Host 'Controle reussi : binaires, deux roles, protocole/schema 2, configuration clinique et Office ferme.' -ForegroundColor Green
Write-Host ('Service : '+$doctor.Url+' / revision '+$doctor.Revision)
Write-Host ('Ancien complement : '+$old.FullName)
if (-not $Execute) { Write-Host 'Aucune ecriture. Relancer avec -Execute -ConfirmerBascule DEPLOYER-C5-NP10 pour installer.'; return }
if (-not $RevisionServiceAttendue) { throw 'Pour installer, preciser -RevisionServiceAttendue avec la revision affichee par le controle.' }
if ($ConfirmerBascule -cne 'DEPLOYER-C5-NP10') { throw 'Option requise : -ConfirmerBascule DEPLOYER-C5-NP10' }

# Chemins uniques : aucune version deja installee n'est ecrasee.
$id = (Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid().ToString('N').Substring(0,6)
$axRoot = Join-Path $env:LOCALAPPDATA ('CabinetCardio\Versions\C5-NP10-'+$id)
$accLocalRoot = 'C:\Users\accueil\AppData\Local\CabinetCardio\Versions\C5-NP10-'+$id
$accRoot = Remote-Path $accLocalRoot
$package = Join-Path $DossierPaquet $id
[void][IO.Directory]::CreateDirectory($package)
$journalPath = Join-Path $package 'deploiement.json'
$localJournalRoot = Join-Path $env:LOCALAPPDATA 'CabinetCardio\Deploiements'
[void][IO.Directory]::CreateDirectory($localJournalRoot)
$localJournalPath = Join-Path $localJournalRoot ($id+'.json')
$journal = [ordered]@{
    kind='CabinetCardio-C5-AX8-ACCUEIL-v2'; state='PREPARATION'; id=$id; date=(Get-Date).ToString('o')
    targets=@('AX8_MAX\olivi','ACCUEIL\accueil'); sourceCommit='f2e74f52e145f3e07480ea9bbeac3bda3fafe346'
    wordSHA256=$HashWord; excelSHA256=$HashExcel; serviceRevision=$doctor.Revision
    oldOriginal=$old.FullName; oldDisabled=($old.FullName+'.desactive-'+$id); oldHash=$oldHash
    oldBackup=(Join-Path $package 'AncienComplement.dotm'); axRoot=$axRoot; accueilRoot=$accLocalRoot
    shortcuts=@(); rollbackError=''; localJournal=$localJournalPath
}
function Save-DeploymentJournal {
    # Ecrire localement en premier : permet la restauration si le NAS se deconnecte.
    Write-Json $localJournalPath $journal
    Write-Json $journalPath $journal
}
Save-DeploymentJournal
try {
    Copy-Item -LiteralPath $old.FullName -Destination $journal.oldBackup
    Assert-Hash $journal.oldBackup $oldHash
    foreach ($target in @($package,$axRoot,$accRoot)) {
        [void][IO.Directory]::CreateDirectory($target)
        # Nom interne necessaire aux macros TrouverModeleCabinet, contenu identique.
        Copy-Item -LiteralPath $sourceWord -Destination (Join-Path $target 'CabinetUnifie.dotm')
        Copy-Item -LiteralPath $sourceExcel -Destination (Join-Path $target 'Cabinet.xlsm')
        Copy-Item -LiteralPath $sourceLaunch -Destination (Join-Path $target 'Lancer-Cabinet-C5NP10.ps1')
        Copy-Item -LiteralPath $sourceCommon -Destination (Join-Path $target 'Controle-Service-C5.ps1')
        Assert-Hash (Join-Path $target 'CabinetUnifie.dotm') $HashWord
        Assert-Hash (Join-Path $target 'Cabinet.xlsm') $HashExcel
        Assert-Hash (Join-Path $target 'Lancer-Cabinet-C5NP10.ps1') $hashLaunch
        Assert-Hash (Join-Path $target 'Controle-Service-C5.ps1') $hashCommon
        $receipt = [ordered]@{ revision=$doctor.Revision; word=$HashWord; excel=$HashExcel; deployment=$id }
        Write-Json (Join-Path $target 'installation.json') $receipt
    }
    # Recontrole immediatement avant d'activer les raccourcis et de neutraliser l'ancien modele.
    Assert-OfficeFerme
    Assert-Hash $old.FullName $oldHash
    $wordLink = Join-Path $desktop ('Cabinet C5 - Medecin - '+$id+'.lnk')
    $excelLink = Join-Path $remoteDesktop ('Cabinet C5 - Secretariat - '+$id+'.lnk')
    $journal['shortcuts'] = @($wordLink,$excelLink)
    $journal['state'] = 'ACTIVATION'; Save-DeploymentJournal
    $ps = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    New-Shortcut $wordLink $ps ('-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $axRoot 'Lancer-Cabinet-C5NP10.ps1')+'" -Profil Medecin') $axRoot
    New-Shortcut $excelLink 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' ('-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $accLocalRoot 'Lancer-Cabinet-C5NP10.ps1')+'" -Profil Secretariat') $accLocalRoot
    Move-Item -LiteralPath $old.FullName -Destination $journal.oldDisabled
    Assert-Hash $journal.oldDisabled $oldHash
    $journal['state'] = 'INSTALLE'; Save-DeploymentJournal
    Write-Host 'Clients installes. Ouvrir les deux NOUVEAUX raccourcis dans les profils prevus.' -ForegroundColor Green
    Write-Host ('Journal/retour arriere : '+$localJournalPath)
    Write-Host 'RDC reste intact. Aucun Office lance, serveur migre, impression ou transaction clinique executee.'
} catch {
    $failure = $_.Exception.Message
    try { Restore-Deployment $journal $localJournalPath }
    catch { $journal['state']='RESTAURATION_INCOMPLETE'; $journal['rollbackError']=$_.Exception.Message; try { Write-Json $localJournalPath $journal } catch {} }
    try { Write-Json $journalPath $journal } catch {}
    throw ('Deploiement interrompu : '+$failure+'. Etat de restauration : '+$journal['state']+'. Journal : '+$localJournalPath)
}
