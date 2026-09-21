$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$source=Join-Path $root 'Build/installer_multi_postes.ps1'
$tokens=$null; $parseErrors=$null
$tree=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count) { throw 'Installateur PowerShell invalide.' }

# Charger uniquement ces fonctions : aucun lancement de l installateur, de Word,
# du reseau ou de la configuration du poste dans cette recette hors Office.
foreach ($name in @('Ecrire-Reglage','Installer-ConnexionService')) {
    $definitions=@($tree.FindAll({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    },$true))
    if ($definitions.Count -ne 1) { throw ('Fonction absente ou ambigue : '+$name) }
    . ([ScriptBlock]::Create($definitions[0].Extent.Text))
}

function Memoriser-Fichier([string]$Destination) {
    if ((Split-Path $Destination -Parent) -ne $local) { throw 'Ecriture hors de la recette.' }
}
function Proteger-FichierLocal([string]$Path) {
    if ($Path -ne (Join-Path $local 'service.token')) { throw 'Fichier inattendu.' }
}
function Invoke-RestMethod {
    param($Method,$Uri,$Headers,$ContentType,$Body,$MaximumRedirection,$TimeoutSec)
    if ($Method -ne 'Post' -or $Uri -cne 'https://nas-recette.invalid:8444/v1/rpc' -or
        $Headers.Authorization -cne ('Bearer '+$fixtureToken) -or
        ($Body | ConvertFrom-Json).operation -cne 'whoami' -or $MaximumRedirection -ne 0) {
        throw 'Requete authentifiee simulee incorrecte.'
    }
    $script:appelsSimules++
    return [pscustomobject]@{result=[pscustomobject]@{
        protocole=2; revision='2026.09.21-u2c'; schema=2; roles=@('medecin','secretariat')
    }}
}
function Read-Host { throw 'Aucune saisie interactive autorisee dans ce test.' }

$local=Join-Path ([IO.Path]::GetTempPath()) ('cabinet-service-url-'+[Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($local)
$fixtureToken='FICTIF-LOCAL-SANS-ACCES-RESEAU-000000000000'
$medecin=$true; $secretariat=$true; $RedemanderConnexion=$false; $FichierJeton=''
$script:appelsSimules=0
try {
    [IO.File]::WriteAllText((Join-Path $local 'service.token'),$fixtureToken,[Text.UTF8Encoding]::new($false))
    foreach ($cas in @('adresse explicite','barre finale','ancien fichier CRLF')) {
        $script:UrlService='https://nas-recette.invalid:8444'
        if ($cas -eq 'barre finale') { $script:UrlService+='/' }
        if ($cas -eq 'ancien fichier CRLF') {
            [IO.File]::WriteAllText((Join-Path $local 'service.url'),"https://nas-recette.invalid:8444`r`n")
            $script:UrlService=''
        }
        Installer-ConnexionService
        $bytes=[IO.File]::ReadAllBytes((Join-Path $local 'service.url'))
        $expected=[Text.Encoding]::ASCII.GetBytes('https://nas-recette.invalid:8444')
        if ([Convert]::ToBase64String($bytes) -cne [Convert]::ToBase64String($expected)) {
            throw ('URL non utilisable par WinHTTP (BOM, CR/LF ou contenu) : '+$cas)
        }
        if ([IO.File]::ReadAllText((Join-Path $local 'service.token')) -cne $fixtureToken) {
            throw 'Le jeton fictif a ete altere.'
        }
        Write-Output ('PASS URL SERVICE : '+$cas)
    }
    if ($script:appelsSimules -ne 3) { throw 'Nombre de verifications simulees incorrect.' }
} finally {
    if (Test-Path -LiteralPath $local) { Remove-Item -LiteralPath $local -Recurse -Force }
}

[executed on device: RDC (851a212c-2a55-4d12-8c1a-7ceb6ea36f05)]