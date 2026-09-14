[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$DossierSauvegarde,[switch]$Appliquer)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'outils_installation.ps1')
if ($env:OS -ne 'Windows_NT') { throw 'Restauration locale Windows uniquement.' }
if (Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue) { throw 'Fermez Word et Excel.' }
$root=(Resolve-Path -LiteralPath $DossierSauvegarde).Path
$items=@(Get-Content -LiteralPath (Join-Path $root 'restauration.json') -Raw -Encoding UTF8 | ConvertFrom-Json)
foreach ($item in $items) {
    if ($item.backup -and (-not (Test-Path -LiteralPath $item.backup -PathType Leaf) -or -not ([IO.Path]::GetFullPath($item.backup).StartsWith($root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)))) { throw 'Sauvegarde manquante ou situee hors du dossier choisi.' }
    Write-Host ('Restaurer : '+$item.destination)
}
if (-not $Appliquer) { Write-Host 'Simulation : aucune modification. Ajouter -Appliquer pour restaurer les fichiers indiques.';return }
$before=Join-Path $root ('avant-restauration-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($before)
for ($i=$items.Count-1;$i -ge 0;$i--) {
    $item=$items[$i]
    if (Test-Path -LiteralPath $item.destination -PathType Leaf) { $copy=Join-Path $before ($i.ToString()+'.bak');[IO.File]::Copy($item.destination,$copy,$false);if ([IO.Path]::GetFileName($item.destination) -eq 'service.token') { Proteger-FichierLocal $copy } }
    Restaurer-FichierAvecDroits $item
}
Write-Host 'Fichiers locaux restaures. Le serveur NAS et ses transactions ne sont pas modifies.'
