# Qualification du script complet, uniquement dans un repertoire temporaire fictif.
# Aucun Office, profil utilisateur, secret reel, reseau ou service NAS n'est utilise.
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Ce test exige Windows pour verifier les ACL reelles.' }
if (Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue) { throw 'Fermez Word et Excel avant cette qualification locale.' }
$root=Split-Path $PSScriptRoot -Parent
$restore=Join-Path $root 'Build/restaurer_poste.ps1'
. (Join-Path $root 'Build/outils_installation.ps1')
$script:u2Checks=0
function Exiger($Condition,[string]$Nom) {
    if (-not $Condition) { throw ('ECHEC RESTAURATION : '+$Nom) }
    $script:u2Checks++
    Write-Output ('PASS RESTAURATION : '+$Nom)
}
function Texte([string]$Path,[string]$Value) {
    [IO.File]::WriteAllText($Path,$Value,(New-Object Text.UTF8Encoding($false)))
}
function Snapshot([string]$Path) {
    $result=@(foreach ($node in Get-ChildItem -LiteralPath $Path -Force -Recurse | Sort-Object FullName) {
        $directory=($node -is [IO.DirectoryInfo])
        [pscustomobject]@{
            path=$node.FullName.Substring($Path.Length);directory=$directory
            sha=$(if ($directory) { '' } else { (Get-FileHash -LiteralPath $node.FullName -Algorithm SHA256).Hash })
            sddl=(Get-Acl -LiteralPath $node.FullName).Sddl
        }
    })
    return ConvertTo-Json -InputObject $result -Depth 4 -Compress
}
function Recu([string]$Directory,[object[]]$Entries) {
    ConvertTo-Json -InputObject $Entries -Depth 5 | Set-Content -LiteralPath (Join-Path $Directory 'restauration.json') -Encoding UTF8
}
function Exiger-RefusSansMutation([string]$Directory,[string]$Message,[string]$Nom) {
    $avant=Snapshot $script:u2Fixture
    $refuse=$false
    try { & $restore -DossierSauvegarde $Directory -Appliquer }
    catch { $refuse=$_.Exception.Message -like $Message }
    Exiger $refuse ($Nom+' : refus explicite')
    Exiger ((Snapshot $script:u2Fixture) -ceq $avant) ($Nom+' : aucun fichier ni droit modifie')
}
$script:u2Fixture=Join-Path ([IO.Path]::GetTempPath()) ('cabinet-restauration-complete-'+[guid]::NewGuid().ToString('N'))
$junction=$null
[void][IO.Directory]::CreateDirectory($script:u2Fixture)
try {
    $targets=Join-Path $script:u2Fixture 'cibles-fictives'
    $backupDir=Join-Path $script:u2Fixture 'sauvegarde-fictive'
    [void][IO.Directory]::CreateDirectory($targets)
    [void][IO.Directory]::CreateDirectory($backupDir)
    $existing=Join-Path $targets 'ancien.txt'
    $created=Join-Path $targets 'ajoute-par-installation.txt'
    $token=Join-Path $targets 'service.token'
    $initial=Join-Path $backupDir 'original.bak'
    $intermediate=Join-Path $backupDir 'intermediaire.bak'
    $tokenBackup=Join-Path $backupDir 'jeton-fictif.bak'
    Texte $existing 'FICTIF ORIGINAL : octets a restaurer'
    Proteger-FichierLocal $existing
    $originalSddl=(Get-Acl -LiteralPath $existing).Sddl
    [IO.File]::Copy($existing,$initial,$false)
    Texte $intermediate 'FICTIF INTERMEDIAIRE : premiere modification'
    Texte $existing 'FICTIF APRES : deuxieme modification'
    $looser=New-Object Security.AccessControl.FileSecurity
    $looser.SetSecurityDescriptorSddlForm($originalSddl,[Security.AccessControl.AccessControlSections]::Access)
    $looser.SetAccessRuleProtection($false,$true)
    [IO.File]::SetAccessControl($existing,$looser)
    Exiger (-not (Get-Acl -LiteralPath $existing).AreAccessRulesProtected) 'DACL fictive differente avant restauration'
    Texte $created 'FICTIF FICHIER AJOUTE'
    Texte $token 'FICTIF-ANCIEN-JETON-SANS-ACCES-000000000000'
    Proteger-FichierLocal $token
    $tokenSddl=(Get-Acl -LiteralPath $token).Sddl
    [IO.File]::Copy($token,$tokenBackup,$false)
    Proteger-FichierLocal $tokenBackup
    Texte $token 'FICTIF-NOUVEAU-JETON-SANS-ACCES-00000000000'
    $entries=@(
        [pscustomobject]@{destination=$existing;backup=$initial;sddl=$originalSddl},
        [pscustomobject]@{destination=$existing;backup=$intermediate;sddl=$originalSddl},
        [pscustomobject]@{destination=$created;backup=$null;sddl=$null},
        [pscustomobject]@{destination=$token;backup=$tokenBackup;sddl=$tokenSddl}
    )
    Recu $backupDir $entries
    $avant=Snapshot $script:u2Fixture
    & $restore -DossierSauvegarde $backupDir
    Exiger ((Snapshot $script:u2Fixture) -ceq $avant) 'simulation du script complet sans aucune mutation'
    & $restore -DossierSauvegarde $backupDir -Appliquer
    Exiger ([IO.File]::ReadAllText($existing) -ceq 'FICTIF ORIGINAL : octets a restaurer') 'restauration en ordre inverse des modifications successives'
    Exiger ((Get-FileHash -LiteralPath $existing).Hash -ceq (Get-FileHash -LiteralPath $initial).Hash) 'octets originaux restaures a l identique'
    Exiger ((Get-Acl -LiteralPath $existing).Sddl -ceq $originalSddl) 'SDDL originale restauree'
    Exiger (-not (Test-Path -LiteralPath $created)) 'suppression du seul fichier fictif ajoute sans sauvegarde'
    Exiger ((Get-FileHash -LiteralPath $token).Hash -ceq (Get-FileHash -LiteralPath $tokenBackup).Hash) 'jeton fictif precedent restaure'
    Exiger ((Get-Acl -LiteralPath $token).Sddl -ceq $tokenSddl) 'droits prives du jeton fictif restaures'
    $beforeDirs=@(Get-ChildItem -LiteralPath $backupDir -Directory -Filter 'avant-restauration-*')
    Exiger ($beforeDirs.Count -eq 1) 'une copie avant restauration est conservee'
    $before=$beforeDirs[0].FullName
    Exiger ([IO.File]::ReadAllText((Join-Path $before '1.bak')) -ceq 'FICTIF APRES : deuxieme modification') 'copie de l etat initial avant restauration'
    Exiger ([IO.File]::ReadAllText((Join-Path $before '0.bak')) -ceq 'FICTIF INTERMEDIAIRE : premiere modification') 'copie intermediaire coherente avec ordre inverse'
    Exiger ([IO.File]::ReadAllText((Join-Path $before '2.bak')) -ceq 'FICTIF FICHIER AJOUTE') 'fichier ajoute conserve avant suppression'
    $savedToken=Join-Path $before '3.bak'
    Exiger ([IO.File]::ReadAllText($savedToken) -ceq 'FICTIF-NOUVEAU-JETON-SANS-ACCES-00000000000') 'copie avant restauration du jeton fictif'
    $savedAcl=Get-Acl -LiteralPath $savedToken
    $savedRules=@($savedAcl.GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier]))
    $allowed=@([Security.Principal.WindowsIdentity]::GetCurrent().User.Value,'S-1-5-18','S-1-5-32-544')
    Exiger ($savedAcl.AreAccessRulesProtected -and $savedRules.Count -eq 3 -and
        @($savedRules | Where-Object { $_.IsInherited -or $_.IdentityReference.Value -notin $allowed -or
            $_.AccessControlType -ne 'Allow' -or $_.FileSystemRights -ne 'FullControl' }).Count -eq 0) 'copie du jeton privee : compte courant, SYSTEM et administrateurs seulement'

    # L'entree invalide est placee AVANT une entree valide, donc serait restauree
    # apres elle. Le script doit refuser l'ensemble avant toute premiere mutation.
    $invalidDir=Join-Path $script:u2Fixture 'sauvegarde-invalide'
    [void][IO.Directory]::CreateDirectory($invalidDir)
    $goodBackup=Join-Path $invalidDir 'valide.bak'
    Texte $goodBackup 'FICTIF VALEUR DE RETOUR ARRIERE'
    Texte $existing 'FICTIF VALEUR QUI DOIT RESTER INTACTE'
    $valid=[pscustomobject]@{destination=$existing;backup=$goodBackup;sddl=$originalSddl}
    $bad=[pscustomobject]@{destination=$token;backup=(Join-Path $invalidDir 'absent.bak');sddl=$tokenSddl}
    Recu $invalidDir @($bad,$valid)
    Exiger-RefusSansMutation $invalidDir '*Sauvegarde manquante*' 'backup manquant'
    $bad.backup=$initial
    Recu $invalidDir @($bad,$valid)
    Exiger-RefusSansMutation $invalidDir '*hors du dossier choisi*' 'backup hors dossier'
    $bad.backup=$goodBackup;$bad.sddl=$null
    Recu $invalidDir @($bad,$valid)
    Exiger-RefusSansMutation $invalidDir '*Sauvegarde des droits absente*' 'SDDL manquante'
    $bad.sddl='SDDL FICTIVE INVALIDE'
    Recu $invalidDir @($bad,$valid)
    Exiger-RefusSansMutation $invalidDir '*Sauvegarde des droits illisible*' 'SDDL malformee'
    $bad.sddl='O:'+([Security.Principal.WindowsIdentity]::GetCurrent().User.Value)
    Recu $invalidDir @($bad,$valid)
    Exiger-RefusSansMutation $invalidDir '*sans DACL explicite*' 'SDDL sans DACL'
    $bad.sddl=$tokenSddl;$bad.destination='fichier-relatif-fictif.txt'
    Recu $invalidDir @($bad,$valid)
    Exiger-RefusSansMutation $invalidDir '*Chemin local absolu requis*' 'destination relative'
    $bad.destination='\\serveur-fictif.invalid\partage\fichier.txt'
    Recu $invalidDir @($bad,$valid)
    Exiger-RefusSansMutation $invalidDir '*Chemin local absolu requis*' 'destination UNC refusee avant tout acces'
    $bad.destination=$targets
    Recu $invalidDir @($bad,$valid)
    Exiger-RefusSansMutation $invalidDir '*Destination de restauration invalide*' 'destination repertoire'
    $bad.destination=$token+':flux'
    Recu $invalidDir @($bad,$valid)
    Exiger-RefusSansMutation $invalidDir '*Flux de fichier non autorise*' 'flux NTFS secondaire'
    $bad.destination=Join-Path $invalidDir 'restauration.json'
    Recu $invalidDir @($bad,$valid)
    Exiger-RefusSansMutation $invalidDir '*propre dossier de sauvegarde*' 'destination dans sauvegarde'
    Recu $invalidDir @([pscustomobject]@{backup=$goodBackup;sddl=$tokenSddl},$valid)
    Exiger-RefusSansMutation $invalidDir '*Entree du recu de restauration incomplete*' 'destination manquante'
    # Une jonction vers une autre zone DU MEME TEST exerce le refus de redirection.
    # Elle est retiree explicitement avant tout nettoyage recursif.
    $junction=Join-Path $invalidDir 'jonction-fictive'
    New-Item -ItemType Junction -Path $junction -Value $backupDir | Out-Null
    $bad.destination=$token;$bad.backup=Join-Path $junction 'original.bak';$bad.sddl=$tokenSddl
    Recu $invalidDir @($bad,$valid)
    $refuse=$false
    $existingBefore=(Get-FileHash -LiteralPath $existing).Hash
    $tokenBefore=(Get-FileHash -LiteralPath $token).Hash
    try { & $restore -DossierSauvegarde $invalidDir -Appliquer }
    catch { $refuse=$_.Exception.Message -like '*Chemin redirige non autorise*' }
    Exiger $refuse 'backup sous jonction refuse'
    Exiger ((Get-FileHash -LiteralPath $existing).Hash -ceq $existingBefore -and
        (Get-FileHash -LiteralPath $token).Hash -ceq $tokenBefore -and
        @(Get-ChildItem -LiteralPath $invalidDir -Directory -Filter 'avant-restauration-*').Count -eq 0) 'refus jonction avant toute mutation des cibles'
    [IO.Directory]::Delete($junction)
    $junction=$null
    Exiger ([IO.File]::Exists($initial)) 'cible de la jonction conservee apres nettoyage'
    [pscustomobject]@{Statut='SUCCESS';Controles=$script:u2Checks;ScriptComplet=$true;Donnees='FICTIVES';AccesNAS=$false;InstallationCliniqueModifiee=$false} | ConvertTo-Json
} finally {
    if ($junction -and (Test-Path -LiteralPath $junction)) { [IO.Directory]::Delete($junction) }
    if (Test-Path -LiteralPath $script:u2Fixture) { Remove-Item -LiteralPath $script:u2Fixture -Recurse -Force }
}
