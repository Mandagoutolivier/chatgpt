Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Lire-Manifeste([string]$Racine) {
    $path = Join-Path $Racine 'Build\manifest.json'
    Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Verifier-ModeleSource([string]$Chemin, [string]$Nom, $Manifeste) {
    if (-not (Test-Path -LiteralPath $Chemin -PathType Leaf)) { throw "Modele absent : $Chemin" }
    $attendu = $Manifeste.models.PSObject.Properties[$Nom].Value
    $actuel = (Get-FileHash -LiteralPath $Chemin -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actuel -ne $attendu) { throw "Empreinte du modele $Nom differente de la version revue. Construction interrompue." }
}

function Lire-CodeVba([string]$Chemin) {
    # CodeModule recoit une chaine Unicode par COM. Aucun import ANSI d'un .bas UTF-8.
    $text = [IO.File]::ReadAllText($Chemin, [Text.Encoding]::UTF8)
    $lines = $text -split '\r?\n' | Where-Object { $_ -notmatch '^Attribute\s' }
    return ($lines -join "`r`n")
}

function Installer-SourcesVba($Projet, $Entrees, [string]$Racine) {
    if ($Projet.Protection -ne 0) { throw 'Le projet VBA est verrouille : utilisez les modeles sources non proteges.' }
    foreach ($entry in $Entrees) {
        $file = Join-Path $Racine $entry.path
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Source VBA absente : $file" }
        $component = $null
        foreach ($candidate in $Projet.VBComponents) {
            if ($candidate.Name -eq $entry.name) { $component = $candidate; break }
        }
        if ($null -eq $component) {
            if ($entry.kind -ne 'module') { throw "Formulaire ou classe absent du modele source : $($entry.name)" }
            $component = $Projet.VBComponents.Add(1)
            $component.Name = $entry.name
        }
        $code = Lire-CodeVba $file
        if ($component.CodeModule.CountOfLines -gt 0) { $component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines) }
        if (-not [string]::IsNullOrWhiteSpace($code)) { $component.CodeModule.AddFromString($code) }
    }
    $attendus=@($Entrees | ForEach-Object { $_.name })
    foreach ($component in @($Projet.VBComponents)) {
        if ($component.Name -notin $attendus) {
            if ($component.Type -eq 100) { throw "Module document inattendu : $($component.Name)" }
            $Projet.VBComponents.Remove($component)
        }
    }
    foreach ($reference in $Projet.References) {
        if ($reference.IsBroken) { throw 'Une reference VBA est manquante sur ce PC. Reparez-la dans Outils > References.' }
    }
}

function Ecrire-EntreeZip($Zip, [string]$Nom, [string]$Texte) {
    $old = $Zip.GetEntry($Nom)
    if ($null -ne $old) { $old.Delete() }
    $entry = $Zip.CreateEntry($Nom)
    $writer = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding($false)))
    try { $writer.Write($Texte) } finally { $writer.Dispose() }
}

function Lire-EntreeZip($Zip, [string]$Nom) {
    $entry = $Zip.GetEntry($Nom)
    if ($null -eq $entry) { throw "Entree OOXML absente : $Nom" }
    $reader = New-Object IO.StreamReader($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Installer-Ruban([string]$Modele, [string]$Ruban) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::Open($Modele, [IO.Compression.ZipArchiveMode]::Update)
    try {
        [xml]$rels = Lire-EntreeZip $zip '_rels/.rels'
        $ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
        foreach ($rel in @($rels.DocumentElement.ChildNodes)) {
            if ($rel.LocalName -eq 'Relationship' -and $rel.GetAttribute('Type') -match '/ui/extensibility$') {
                [void]$rels.DocumentElement.RemoveChild($rel)
            }
        }
        $ids = @($rels.DocumentElement.ChildNodes | ForEach-Object { $_.GetAttribute('Id') })
        $id = 'rIdCabinetUnifie'
        while ($ids -contains $id) { $id += 'X' }
        $rel = $rels.CreateElement('Relationship', $ns)
        $rel.SetAttribute('Id', $id)
        $rel.SetAttribute('Type', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility')
        $rel.SetAttribute('Target', 'customUI/cabinet.xml')
        [void]$rels.DocumentElement.AppendChild($rel)
        Ecrire-EntreeZip $zip '_rels/.rels' $rels.OuterXml
        [xml]$types = Lire-EntreeZip $zip '[Content_Types].xml'
        $found = $false
        foreach ($node in $types.DocumentElement.ChildNodes) {
            if ($node.LocalName -eq 'Override' -and $node.GetAttribute('PartName') -eq '/customUI/cabinet.xml') { $found = $true }
        }
        if (-not $found) {
            $node = $types.CreateElement('Override', $types.DocumentElement.NamespaceURI)
            $node.SetAttribute('PartName', '/customUI/cabinet.xml')
            $node.SetAttribute('ContentType', 'application/xml')
            [void]$types.DocumentElement.AppendChild($node)
        }
        Ecrire-EntreeZip $zip '[Content_Types].xml' $types.OuterXml
        Ecrire-EntreeZip $zip 'customUI/cabinet.xml' ([IO.File]::ReadAllText($Ruban, [Text.Encoding]::UTF8))
    } finally { $zip.Dispose() }
}

function Publier-FichierConstruit([string]$Source, [string]$Destination) {
    # L'echec d'une construction ne modifie jamais le fichier de sortie precedent.
    $parent = Split-Path $Destination -Parent
    [void][IO.Directory]::CreateDirectory($parent)
    $temp = Join-Path $parent ([guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [IO.File]::Copy($Source, $temp, $false)
        if ([IO.File]::Exists($Destination)) {
            $backup = $Destination + '.avant-' + [guid]::NewGuid().ToString('N')
            [IO.File]::Replace($temp, $Destination, $backup)
        } else { [IO.File]::Move($temp, $Destination) }
    } finally { if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) } }
}
