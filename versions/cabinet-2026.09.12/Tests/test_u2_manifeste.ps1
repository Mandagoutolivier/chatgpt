$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'Build/outils_construction.ps1')
. (Join-Path $root 'Build/outils_recette_u1.ps1')
. (Join-Path $root 'Build/outils_configuration.ps1')
function Exiger-U2($valeur,$nom){if(-not $valeur){throw ('ECHEC U2 : '+$nom)};Write-Output ('PASS U2 : '+$nom)}
$prod=Lire-Manifeste $root
$recette=Lire-Manifeste $root -InclureRecette
foreach($hote in @('word','excel')){
    $tests=@($prod.($hote+'_recette')|ForEach-Object{$_.name})
    Exiger-U2 (@($prod.$hote|Where-Object{$_.name -in $tests}).Count -eq 0) ('production sans modules de recette '+$hote)
    Exiger-U2 (@($recette.$hote|Where-Object{$_.name -in $tests}).Count -eq $tests.Count) ('recette explicite complete '+$hote)
    Exiger-U2 (@($recette.$hote|Group-Object name|Where-Object{$_.Count -ne 1}).Count -eq 0) ('composants uniques '+$hote)
}
$encore=Lire-Manifeste $root
Exiger-U2 ($encore.word.Count -eq $prod.word.Count -and $encore.excel.Count -eq $prod.excel.Count) 'lecture de recette sans contamination de la production'

function Exiger-RefusRecette([string]$json,[string]$nom) {
    $refuse=$false
    try { Verifier-ResultatRecetteOffice $json $nom | Out-Null } catch { $refuse=$true }
    Exiger-U2 $refuse ('contrat refuse '+$nom)
}
Verifier-ResultatRecetteOffice '{"reussis":19,"attendus":19,"echec":false}' 'compte futur' | Out-Null
Exiger-U2 $true 'contrat accepte un compte futur coherent'
Exiger-RefusRecette '{"reussis":17,"attendus":18,"echec":false}' 'compteurs divergents'
Exiger-RefusRecette '{"reussis":17,"echec":false}' 'attendus absent'
Exiger-RefusRecette '{"reussis":18,"attendus":18,"echec":true,"description":"fictif"}' 'echec explicite'
Exiger-RefusRecette 'json-invalide' 'json invalide'


function Tester-ConfigurationSortie([string]$contenu,[bool]$doitReussir,[string]$nom) {
    $chemin=Join-Path ([IO.Path]::GetTempPath()) ('cabinet-config-'+[guid]::NewGuid().ToString('N')+'.ini')
    try {
        [IO.File]::WriteAllText($chemin,$contenu,[Text.UTF8Encoding]::new($false))
        $reussit=$true
        try { Verifier-ConfigurationSortie $chemin } catch { $reussit=$false }
        Exiger-U2 ($reussit -eq $doitReussir) ('configuration sortie '+$nom)
    } finally {
        Remove-Item -LiteralPath $chemin -Force -ErrorAction SilentlyContinue
    }
}
Tester-ConfigurationSortie "[SORTIE]`nExportActif=1`nDossier=\\NAS-RECETTE\CabinetCardioTestU2\Echange\AEnvoyer`nNomFichier=PublicationID" $true 'active explicite'
Tester-ConfigurationSortie "[SORTIE]`nExportActif=0`nDossier=\\NAS-RECETTE\CabinetCardioTestU2\Echange\AEnvoyer`nNomFichier=IdentitePublication" $true 'inactive explicite'
Tester-ConfigurationSortie "[SORTIE]`nDossier=\\NAS-RECETTE\CabinetCardioTestU2\Echange\AEnvoyer`nNomFichier=PublicationID" $false 'activation absente'
Tester-ConfigurationSortie "[SORTIE]`nExportActif=oui`nDossier=\\NAS-RECETTE\CabinetCardioTestU2\Echange\AEnvoyer`nNomFichier=PublicationID" $false 'activation invalide'
Tester-ConfigurationSortie "[SORTIE]`nExportActif=1`nDossier=`nNomFichier=PublicationID" $false 'dossier absent'
Tester-ConfigurationSortie "[SORTIE]`nExportActif=1`nDossier=\\NAS-RECETTE\CabinetCardioTestU2\Echange\AEnvoyer`nNomFichier=Ancien" $false 'strategie de nom absente'
