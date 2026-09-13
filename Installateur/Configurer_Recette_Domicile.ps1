#requires -Version 5.1
[CmdletBinding()]
param()

# Configuration de recette pour la preparation 17c84b9 deja compilee.
# Ne construit ni n'installe de complement Office. Ne valide aucun essai.
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Executer dans Windows PowerShell, avec votre session habituelle.' }

$racineNas = '\\DS224\CabinetCardioTest'
$urlService = 'https://192.168.10.1:8443'
$localCabinet = Join-Path $env:APPDATA 'CabinetCardio'
$dossierGdt = Join-Path $env:LOCALAPPDATA 'CabinetCardioTest\GDT'
$sourcesAttendues = Join-Path $env:LOCALAPPDATA 'CabinetCardio\Installation\Sources\17c84b98374f54c076930583796e50602458d294'
$etatPath = Join-Path $localCabinet 'Assistant\Domicile.json'
$verrous = New-Object 'System.Collections.Generic.List[object]'
$changements = New-Object 'System.Collections.Generic.List[object]'
$secret = $null; $jeton = $null; $entetes = $null
$ptr = [IntPtr]::Zero
$ancienTls = [Net.ServicePointManager]::SecurityProtocol
$ecrituresCommencees = $false
$gdtCree = $false
$sauvegarde = ''
$encodage = New-Object Text.UTF8Encoding($false)

function Verifier-OfficeFerme {
    if (@(Get-Process -Name WINWORD,EXCEL -ErrorAction SilentlyContinue).Count) {
        throw 'Fermez completement Word et Excel, puis relancez ce script.'
    }
}

function Memoriser-Reglage([string]$Chemin) {
    $existe = [IO.File]::Exists($Chemin)
    $copie = $null; $acl = $null
    if ($existe) {
        $acl = Get-Acl -LiteralPath $Chemin
        $copie = Join-Path $sauvegarde ($changements.Count.ToString('00') + '.bak')
        [IO.File]::WriteAllBytes($copie, [byte[]]@())
        Proteger-FichierLocal $copie
        [IO.File]::WriteAllBytes($copie, [IO.File]::ReadAllBytes($Chemin))
    }
    $changements.Add([pscustomobject]@{destination=$Chemin;existe=$existe;copie=$copie;acl=$acl})
}

try {
    Verifier-OfficeFerme
    if (-not (Test-Path -LiteralPath $etatPath -PathType Leaf)) {
        throw 'Preparation Domicile absente. Ce correctif exige la preparation 17c84b9 deja compilee.'
    }
    foreach ($nom in @('Assistant\assistant.lock', 'installation.lock')) {
        $verrous.Add([IO.File]::Open((Join-Path $localCabinet $nom),
            [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None))
    }
    $etat = Get-Content -LiteralPath $etatPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($etat.profil -ne 'Domicile' -or $etat.phase -notin @('valide','compile') -or
        $etat.racineSources -ne $sourcesAttendues) {
        throw 'Etat inattendu : ce script concerne uniquement la preparation Domicile 17c84b9 non activee.'
    }
    $stage = [IO.Path]::GetFullPath([string]$etat.dossierPrepare)
    $prefixeVersions = [IO.Path]::GetFullPath((Join-Path $localCabinet 'Versions')) + [IO.Path]::DirectorySeparatorChar
    if (-not $stage.StartsWith($prefixeVersions, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Le dossier prepare doit se trouver dans CabinetCardio\Versions.'
    }
    $outil = Join-Path $sourcesAttendues 'Build\outils_installation.ps1'
    if ((Get-FileHash -LiteralPath $outil -Algorithm SHA256).Hash -ne 'DAA32259ACA3C5266F2F6114EB6BE008EFCF02A31BCD5EE95B862FEE0178E793') {
        throw 'Outils de verification differents de la version 17c84b9 attendue.'
    }
    . $outil
    Verifier-Preparation $stage 'Domicile' $sourcesAttendues
    if (-not $etat.wordCompile -or -not $etat.excelCompile -or
        $etat.wordHash -ne (Get-FileHash -LiteralPath (Join-Path $stage 'CabinetUnifie.dotm')).Hash -or
        $etat.excelHash -ne (Get-FileHash -LiteralPath (Join-Path $stage 'Cabinet.xlsm')).Hash) {
        throw 'Les compilations memorisees ne correspondent plus aux fichiers prepares.'
    }
    $recuPath = Join-Path $stage 'preparation.json'
    $recu = Get-Content -LiteralPath $recuPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $configNas = Join-Path $racineNas 'Config\config.ini'
    $section = ''; $sortie = ''
    foreach ($ligne in Get-Content -LiteralPath $configNas -Encoding UTF8) {
        if ($ligne -match '^\s*\[([^]]+)\]') { $section = $Matches[1].Trim() }
        elseif ($section -eq 'SORTIE' -and $ligne -match '^\s*Dossier\s*=(.*)$') {
            $sortie = $Matches[1].Trim().TrimEnd([char]'\')
        }
    }
    if ($sortie -ne ($racineNas + '\sortiedragon')) {
        throw 'Le dossier SORTIE du NAS doit etre \\DS224\CabinetCardioTest\sortiedragon.'
    }

    Write-Host 'Fichiers prepares verifies. Configuration du poste pour les essais uniquement.'
    Write-Host ('NAS : ' + $racineNas)
    Write-Host ('GDT de test : ' + $dossierGdt)
    Write-Host 'Copiez maintenant le jeton depuis votre gestionnaire de mots de passe.'
    $secret = Read-Host 'Jeton domicile-test' -AsSecureString
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret)
    $jeton = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    if ($jeton -cnotmatch '\A[A-Za-z0-9_-]{64}\z') {
        throw 'Le jeton doit comporter les 64 caracteres generes a la creation du compte.'
    }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $entetes = @{Authorization=('Bearer ' + $jeton)}
    try {
        $reponse = Invoke-RestMethod -Method Post -Uri ($urlService + '/v1/rpc') `
            -Headers $entetes -ContentType 'application/json' `
            -Body '{"operation":"whoami","params":{}}' -MaximumRedirection 0 -TimeoutSec 20
    } catch {
        $code = 'indisponible'
        $http = $_.Exception.PSObject.Properties['Response']
        if ($null -ne $http -and $null -ne $http.Value) { $code = [int]$http.Value.StatusCode }
        throw "Connexion refusee avant toute modification des reglages. Code HTTP : $code"
    }
    $compte = $reponse.result
    if ($compte.ID -ne 'domicile-test' -or $compte.protocole -ne 2 -or
        'medecin' -notin $compte.roles -or 'secretariat' -notin $compte.roles) {
        throw 'Le service doit reconnaitre domicile-test avec les roles medecin et secretariat.'
    }
    Verifier-OfficeFerme

    $cheminPath = Join-Path $localCabinet 'chemin.txt'
    $postePath = Join-Path $localCabinet 'poste.ini'
    $urlPath = Join-Path $localCabinet 'service.url'
    $jetonPath = Join-Path $localCabinet 'service.token'
    $sqlitePath = Join-Path $localCabinet 'Tools\sqlite3.exe'
    $id = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8)
    $sauvegarde = Join-Path $localCabinet ('Sauvegardes\recette-test-' + $id)
    [void][IO.Directory]::CreateDirectory($sauvegarde)
    foreach ($chemin in @($cheminPath,$postePath,$urlPath,$jetonPath,$sqlitePath,$etatPath,$recuPath)) {
        Memoriser-Reglage $chemin
    }
    $journal = @($changements | ForEach-Object {
        $sddl = $null
        if ($null -ne $_.acl) { $sddl = $_.acl.Sddl }
        [pscustomobject]@{destination=$_.destination;existait=$_.existe;sauvegarde=$_.copie;aclSddl=$sddl}
    })
    [IO.File]::WriteAllText((Join-Path $sauvegarde 'restauration.json'),
        (ConvertTo-Json -InputObject $journal -Depth 5), $encodage)

    $poste = ''
    if ([IO.File]::Exists($postePath)) { $poste = [IO.File]::ReadAllText($postePath, [Text.Encoding]::UTF8) }
    $poste += "`r`n[POSTE]`r`nProfil=Domicile`r`n[ECG]`r`nDossierGdt=$dossierGdt`r`n"
    $etat.racineNas = $racineNas
    $etat.dossierGdt = $dossierGdt
    $etat.recette = $false
    $etat.phase = 'compile'
    $recu | Add-Member -NotePropertyName recetteValidee -NotePropertyValue $false -Force
    $etatTexte = $etat | ConvertTo-Json -Depth 8
    $recuTexte = $recu | ConvertTo-Json -Depth 8

    $ecrituresCommencees = $true
    # Invalider d'abord l'ancienne recette. Aucune modification des sources ou binaires Office.
    [IO.File]::WriteAllText($recuPath, $recuTexte, $encodage)
    [IO.File]::WriteAllText($etatPath, $etatTexte, $encodage)
    if (-not [IO.Directory]::Exists($dossierGdt)) {
        [void][IO.Directory]::CreateDirectory($dossierGdt)
        $gdtCree = $true
    }
    [void][IO.Directory]::CreateDirectory((Split-Path $sqlitePath -Parent))
    [IO.File]::WriteAllBytes($sqlitePath, [IO.File]::ReadAllBytes((Join-Path $stage 'sqlite3.exe')))
    if ((Get-FileHash -LiteralPath $sqlitePath).Hash -ne $recu.binaires.'sqlite3.exe') {
        throw 'La copie du programme SQLite ne correspond pas au fichier prepare.'
    }
    [IO.File]::WriteAllText($postePath, $poste, $encodage)
    # VBA Trim$ retire les espaces, mais pas CR/LF : conserver l'URL sans fin de ligne.
    [IO.File]::WriteAllText($urlPath, $urlService, $encodage)
    # Restreindre les droits du fichier AVANT d'y enregistrer le jeton.
    if (-not [IO.File]::Exists($jetonPath)) { [IO.File]::WriteAllBytes($jetonPath, [byte[]]@()) }
    Proteger-FichierLocal $jetonPath
    [IO.File]::WriteAllText($jetonPath, $jeton, $encodage)
    [IO.File]::WriteAllText($cheminPath, ($racineNas + "`r`n"), $encodage)

    Write-Host ''
    Write-Host 'OK : configuration locale prete pour la recette Domicile.'
    Write-Host ('NAS : ' + $racineNas)
    Write-Host ('Service : ' + $urlService)
    Write-Host ('GDT de test : ' + $dossierGdt)
    Write-Host ('Fichiers Office prepares : ' + $stage)
    Write-Host ('Sauvegarde des anciens reglages : ' + $sauvegarde)
    Write-Host 'Jeton enregistre avec des droits restreints. Aucun jeton affiche.'
    Write-Host 'Compilation conservee ; recette a effectuer. Aucun complement Office actif installe.'
} catch {
    $cause = $_
    if ($ecrituresCommencees) {
        $echecs = New-Object 'System.Collections.Generic.List[string]'
        for ($i=$changements.Count-1; $i -ge 0; $i--) {
            $item = $changements[$i]
            try {
                if ($item.existe) {
                    [IO.File]::WriteAllBytes($item.destination, [IO.File]::ReadAllBytes($item.copie))
                    Set-Acl -LiteralPath $item.destination -AclObject $item.acl
                } elseif ([IO.File]::Exists($item.destination)) {
                    [IO.File]::Delete($item.destination)
                }
            } catch { $echecs.Add($item.destination) }
        }
        if ($gdtCree) {
            try { [IO.Directory]::Delete($dossierGdt, $false) }
            catch { Write-Warning ('Dossier de test conserve : ' + $dossierGdt) }
        }
        if ($echecs.Count) { Write-Warning ('Restauration incomplete : ' + ($echecs -join ', ')) }
        else { Write-Host 'Les anciens reglages ont ete restaures.' }
        Write-Host ('Sauvegarde et journal de restauration : ' + $sauvegarde)
    }
    throw $cause
} finally {
    if ($ptr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
    if ($null -ne $secret) { $secret.Dispose() }
    $jeton = $null; $entetes = $null
    [Net.ServicePointManager]::SecurityProtocol = $ancienTls
    foreach ($verrou in $verrous) { $verrou.Dispose() }
}
