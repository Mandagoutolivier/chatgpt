[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$DossierSauvegarde,[switch]$Appliquer)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'outils_installation.ps1')
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Restauration locale Windows uniquement.' }
if (Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue) { throw 'Fermez Word et Excel.' }
# Tout le recu doit etre coherent avant la premiere ecriture. Une erreur dans
# une entree ancienne ne doit pas laisser les entrees suivantes deja restaurees.
function Chemin-LocalRestauration($Valeur) {
    if ($Valeur -isnot [string] -or [string]::IsNullOrWhiteSpace($Valeur) -or
        $Valeur -notmatch '^[A-Za-z]:[\\/]') { throw 'Chemin local absolu requis dans le recu de restauration.' }
    # .NET Framework peut refuser un flux NTFS avant la normalisation du chemin.
    if ($Valeur.Substring(2).Contains(':')) { throw 'Flux de fichier non autorise dans le recu de restauration.' }
    $full=[IO.Path]::GetFullPath($Valeur)
    foreach ($part in $full.Substring(3).Split([char]'\')) {
        if ($part -and ($part -match '[<>:"|?*\x00-\x1f]' -or $part.EndsWith('.') -or $part.EndsWith(' ') -or
            $part -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)')) { throw 'Composant de chemin invalide dans le recu de restauration.' }
    }
    $drive=New-Object IO.DriveInfo ([IO.Path]::GetPathRoot($full))
    if ($drive.DriveType -eq [IO.DriveType]::Network) { throw 'Lecteur reseau non autorise pour la restauration locale.' }
    $cursor=$full
    while ($cursor) {
        $node=$null
        try { $node=Get-Item -LiteralPath $cursor -Force -ErrorAction Stop }
        catch { if ($_.CategoryInfo.Category -ne [Management.Automation.ErrorCategory]::ObjectNotFound) { throw } }
        if ($null -ne $node -and ($node.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw 'Chemin redirige non autorise pour la restauration locale.'
        }
        if ($cursor -eq [IO.Path]::GetPathRoot($cursor)) { break }
        $parent=[IO.Path]::GetDirectoryName($cursor.TrimEnd([char]'\'))
        if ($parent -eq $cursor) { break }
        $cursor=$parent
    }
    return $full
}
$root=(Chemin-LocalRestauration $DossierSauvegarde).TrimEnd([char]'\')
if (-not [IO.Directory]::Exists($root)) { throw 'Dossier de sauvegarde absent.' }
$receipt=Chemin-LocalRestauration (Join-Path $root 'restauration.json')
if (-not [IO.File]::Exists($receipt)) { throw 'Recu de restauration absent.' }
# Windows PowerShell 5.1 peut emettre le tableau JSON comme un seul objet.
# Normaliser apres l'affectation pour accepter les recus a une ou plusieurs entrees.
$receiptData=Get-Content -LiteralPath $receipt -Raw -Encoding UTF8 | ConvertFrom-Json
$items=@($receiptData)
if ($items.Count -eq 0) { throw 'Recu de restauration vide.' }
$plan=New-Object 'System.Collections.Generic.List[object]'
foreach ($item in $items) {
    if ($null -eq $item -or $item -isnot [pscustomobject] -or
        -not $item.PSObject.Properties['destination'] -or -not $item.PSObject.Properties['backup']) {
        throw 'Entree du recu de restauration incomplete.'
    }
    $destination=Chemin-LocalRestauration $item.destination
    if ($destination.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase) -or $destination -eq $root) {
        throw 'Une destination ne peut pas remplacer son propre dossier de sauvegarde.'
    }
    if ([IO.Directory]::Exists($destination) -or -not [IO.Directory]::Exists([IO.Path]::GetDirectoryName($destination))) {
        throw 'Destination de restauration invalide : fichier attendu dans un dossier existant.'
    }
    $backup=$null;$sddl=$null
    if ($null -ne $item.backup) {
        $backup=Chemin-LocalRestauration $item.backup
        if (-not $backup.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase) -or -not [IO.File]::Exists($backup)) {
            throw 'Sauvegarde manquante ou situee hors du dossier choisi.'
        }
        if (-not $item.PSObject.Properties['sddl'] -or $item.sddl -isnot [string] -or [string]::IsNullOrWhiteSpace($item.sddl)) {
            throw 'Sauvegarde des droits absente dans le recu de restauration.'
        }
        $sddl=[string]$item.sddl
        try { $saved=New-Object Security.AccessControl.RawSecurityDescriptor($sddl) }
        catch { throw 'Sauvegarde des droits illisible dans le recu de restauration.' }
        if ($null -eq $saved.DiscretionaryAcl) { throw 'Sauvegarde des droits sans DACL explicite : restauration interrompue.' }
    }
    $plan.Add([pscustomobject]@{destination=$destination;backup=$backup;sddl=$sddl})
}
$items=@($plan.ToArray())
foreach ($item in $items) { Write-Host ('Restaurer : '+$item.destination) }
if (-not $Appliquer) { Write-Host 'Simulation : aucune modification. Ajouter -Appliquer pour restaurer les fichiers indiques.';return }
$before=Join-Path $root ('avant-restauration-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($before)
for ($i=$items.Count-1;$i -ge 0;$i--) {
    $item=$items[$i]
    if (Test-Path -LiteralPath $item.destination -PathType Leaf) { $copy=Join-Path $before ($i.ToString()+'.bak');[IO.File]::Copy($item.destination,$copy,$false);if ([IO.Path]::GetFileName($item.destination) -eq 'service.token') { Proteger-FichierLocal $copy } }
    Restaurer-FichierAvecDroits $item
}
Write-Host 'Fichiers locaux restaures. Le serveur NAS et ses transactions ne sont pas modifies.'
