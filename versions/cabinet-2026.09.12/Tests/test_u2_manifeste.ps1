$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'Build/outils_construction.ps1')
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
