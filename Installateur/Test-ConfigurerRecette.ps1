# Tests Windows 5.1 : fichiers et ACL reels ; NAS, saisie et HTTP simules.
# Aucun Office, serveur, jeton reel ou fichier du profil courant n'est utilise.
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Ce test des ACL exige Windows.' }
$helper = Join-Path $PSScriptRoot 'Configurer_Recette_Domicile.ps1'
$fixture = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'Fixtures\outils_installation_17c84b9.ps1')).Replace("`r`n","`n")
$temp = Join-Path ([IO.Path]::GetTempPath()) ('Recette test & espaces ' + [guid]::NewGuid().ToString('N'))
$ancienAppdata = $env:APPDATA; $ancienLocal = $env:LOCALAPPDATA
$tlsAvant = [Net.ServicePointManager]::SecurityProtocol
$enc = New-Object Text.UTF8Encoding($false)
$global:RecetteJetonFictif = 'T' * 64

function Assert-Recette([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Get-Process { param($Name, $ErrorAction) }
function Get-Content {
    param([string]$LiteralPath, [switch]$Raw, [string]$Encoding)
    if ($LiteralPath -eq '\\DS224\CabinetCardioTest\Config\config.ini') {
        '[ECG]', 'DossierGdt=', '[SORTIE]', 'Dossier=\\DS224\CabinetCardioTest\sortiedragon'
    } else { Microsoft.PowerShell.Management\Get-Content @PSBoundParameters }
}
function Read-Host {
    param([string]$Prompt, [switch]$AsSecureString)
    $global:RecetteLectures++
    ConvertTo-SecureString $global:RecetteJetonFictif -AsPlainText -Force
}
function Invoke-RestMethod {
    param($Method,$Uri,$Headers,$ContentType,$Body,$MaximumRedirection,$TimeoutSec)
    $global:RecetteAppels++
    Assert-Recette ($Uri -eq 'https://192.168.10.1:8443/v1/rpc') 'Destination HTTP inattendue.'
    Assert-Recette ($Headers.Authorization -eq ('Bearer ' + $global:RecetteJetonFictif)) 'Transport du jeton incorrect.'
    Assert-Recette ($MaximumRedirection -eq 0) 'Redirections autorisees.'
    Assert-Recette (($Body | ConvertFrom-Json).operation -eq 'whoami') 'Operation autre que lecture.'
    if ($global:RecetteCas -eq 'http-refuse') { throw [InvalidOperationException]::new('Refus HTTP simule sans secret.') }
    @{result=@{ID='domicile-test';roles=@('medecin','secretariat');protocole=2;version='2026.09.12'}}
}

try {
    foreach ($cas in @('succes','http-refuse','echec-ecriture','binaire-altere')) {
        $global:RecetteCas=$cas; $global:RecetteAppels=0; $global:RecetteLectures=0
        $dossierCas=Join-Path $temp $cas
        $env:APPDATA=Join-Path $dossierCas 'Roaming'
        $env:LOCALAPPDATA=Join-Path $dossierCas 'Local'
        $local=Join-Path $env:APPDATA 'CabinetCardio'
        $stage=Join-Path $local 'Versions\prepare'
        $sources=Join-Path $env:LOCALAPPDATA 'CabinetCardio\Installation\Sources\17c84b98374f54c076930583796e50602458d294'
        foreach ($dossier in @($stage,(Join-Path $local 'Assistant'),(Join-Path $sources 'Build'),(Join-Path $sources 'Src'))) {
            [void][IO.Directory]::CreateDirectory($dossier)
        }
        $outil=Join-Path $sources 'Build\outils_installation.ps1'
        [IO.File]::WriteAllText($outil,$fixture,$enc)
        . $outil
        foreach ($nom in @('CabinetUnifie.dotm','Cabinet.xlsm','sqlite3.exe')) {
            [IO.File]::WriteAllText((Join-Path $stage $nom),('Fichier fictif non executable : ' + $nom),$enc)
        }
        Ecrire-Preparation $stage 'Domicile' $sources
        $recuPath=Join-Path $stage 'preparation.json'
        $recu=Get-Content -LiteralPath $recuPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $recu | Add-Member -NotePropertyName recetteValidee -NotePropertyValue $true
        [IO.File]::WriteAllText($recuPath,($recu | ConvertTo-Json -Depth 8),$enc)
        $etatPath=Join-Path $local 'Assistant\Domicile.json'
        $etat=[pscustomobject]@{
            profil='Domicile';phase='valide';racineSources=$sources;dossierPrepare=$stage
            racineNas='\\DS224\CabinetCardio';dossierGdt='C:\Mandagout';recette=$true
            wordCompile=$true;excelCompile=$true;wordHash=$recu.binaires.'CabinetUnifie.dotm';excelHash=$recu.binaires.'Cabinet.xlsm'
        }
        [IO.File]::WriteAllText($etatPath,($etat | ConvertTo-Json -Depth 8),$enc)
        $cheminPath=Join-Path $local 'chemin.txt'
        if ($cas -eq 'echec-ecriture') { [void][IO.Directory]::CreateDirectory($cheminPath) }
        else { [IO.File]::WriteAllText($cheminPath,'C:\CabinetCardio',$enc) }
        $jetonPath=Join-Path $local 'service.token'
        [IO.File]::WriteAllText($jetonPath,'ancien-jeton-fictif',$enc)
        Proteger-FichierLocal $jetonPath
        $aclAvant=(Get-Acl -LiteralPath $jetonPath).Sddl
        $hashEtat=(Get-FileHash -LiteralPath $etatPath).Hash
        $hashRecu=(Get-FileHash -LiteralPath $recuPath).Hash
        if ($cas -eq 'binaire-altere') { [IO.File]::AppendAllText((Join-Path $stage 'Cabinet.xlsm'),'altere') }
        $erreur=$null
        try { & $helper *> $null } catch { $erreur=$_ }

        if ($cas -eq 'succes') {
            if ($null -ne $erreur) { throw $erreur }
            $apres=Get-Content -LiteralPath $etatPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $recuApres=Get-Content -LiteralPath $recuPath -Raw -Encoding UTF8 | ConvertFrom-Json
            Assert-Recette ($apres.phase -eq 'compile' -and -not $apres.recette -and -not $recuApres.recetteValidee) 'Une recette a ete conservee ou inventee.'
            Assert-Recette ($apres.wordHash -eq $etat.wordHash -and $apres.excelHash -eq $etat.excelHash) 'Compilation modifiee.'
            Assert-Recette ([IO.File]::ReadAllText($cheminPath).Trim() -eq '\\DS224\CabinetCardioTest') 'Racine incorrecte.'
            Assert-Recette ($apres.dossierGdt -eq (Join-Path $env:LOCALAPPDATA 'CabinetCardioTest\GDT')) 'GDT mal isole.'
            Assert-Recette ([IO.File]::ReadAllText($jetonPath) -eq $global:RecetteJetonFictif) 'Jeton non enregistre.'
            $acl=Get-Acl -LiteralPath $jetonPath
            Assert-Recette $acl.AreAccessRulesProtected 'ACL du jeton heritee.'
            $sids=@([Security.Principal.WindowsIdentity]::GetCurrent().User.Value,'S-1-5-18','S-1-5-32-544')
            foreach ($regle in $acl.Access) {
                $sid=$regle.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value
                Assert-Recette ($sid -in $sids) 'Acces au jeton accorde a un tiers.'
            }
            & $helper *> $null
            Assert-Recette (-not (Get-Content -LiteralPath $etatPath -Raw -Encoding UTF8 | ConvertFrom-Json).recette) 'La relance valide implicitement la recette.'
        } else {
            Assert-Recette ($null -ne $erreur) ('Echec attendu non detecte : ' + $cas)
            Assert-Recette ((Get-FileHash -LiteralPath $etatPath).Hash -eq $hashEtat) 'Etat non restaure.'
            Assert-Recette ((Get-FileHash -LiteralPath $recuPath).Hash -eq $hashRecu) 'Recu non restaure.'
            Assert-Recette ([IO.File]::ReadAllText($jetonPath) -eq 'ancien-jeton-fictif') 'Ancien jeton perdu.'
            Assert-Recette ((Get-Acl -LiteralPath $jetonPath).Sddl -eq $aclAvant) 'Droits du jeton non restaures.'
            Assert-Recette (-not [IO.File]::Exists((Join-Path $local 'poste.ini'))) 'Configuration partielle conservee.'
            if ($cas -eq 'binaire-altere') { Assert-Recette ($global:RecetteLectures -eq 0) 'Jeton demande malgre binaire altere.' }
            if ($cas -eq 'echec-ecriture') {
                Assert-Recette (-not [IO.Directory]::Exists((Join-Path $env:LOCALAPPDATA 'CabinetCardioTest\GDT'))) 'Dossier GDT temporaire non retire.'
            }
        }
        Assert-Recette ([Net.ServicePointManager]::SecurityProtocol -eq $tlsAvant) 'Reglage TLS non restaure.'
        Write-Host ('PASS : recette Domicile - ' + $cas)
    }
} finally {
    $env:APPDATA=$ancienAppdata; $env:LOCALAPPDATA=$ancienLocal
    foreach ($nom in @('RecetteCas','RecetteAppels','RecetteLectures','RecetteJetonFictif')) {
        Remove-Variable -Name $nom -Scope Global -ErrorAction SilentlyContinue
    }
    if ([IO.Directory]::Exists($temp)) { Remove-Item -LiteralPath $temp -Recurse -Force }
}
