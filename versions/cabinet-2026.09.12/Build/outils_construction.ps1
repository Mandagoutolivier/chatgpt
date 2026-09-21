Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Lire-Manifeste([string]$Racine, [switch]$InclureRecette) {
    $path = Join-Path $Racine 'Build\manifest.json'
    $manifest=Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($InclureRecette) {
        foreach ($hostName in @('word','excel')) { $manifest.$hostName=@($manifest.$hostName)+@($manifest.($hostName+'_recette')) }
    }
    return $manifest
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

function Lier-FeuillesVba($Classeur, $Entrees) {
    # Le CodeName automatique d une feuille sans code peut dependre de la langue d Excel.
    # Identifier la feuille par son nom dans le classeur source verifie, puis fixer le nom du composant.
    $bindings=New-Object 'System.Collections.Generic.List[object]'
    foreach ($entry in $Entrees) {
        if ($entry.kind -ne 'document' -or $null -eq $entry.PSObject.Properties['sheet']) { continue }
        try { $sheet=$Classeur.Worksheets.Item([string]$entry.sheet) }
        catch { throw "Feuille Excel absente du modele : $($entry.sheet)" }
        $current=[string]$sheet.CodeName
        $component=$null
        foreach ($candidate in $Classeur.VBProject.VBComponents) {
            if ($candidate.Name -eq $current) { $component=$candidate }
            if ($candidate.Name -eq $entry.name -and $candidate.Name -ne $current) {
                throw "Nom VBA deja utilise par un autre composant : $($entry.name)"
            }
        }
        if ($null -eq $component -or $component.Type -ne 100) { throw "Module de feuille Excel introuvable : $($entry.sheet)" }
        $bindings.Add([pscustomobject]@{component=$component;name=$entry.name})
    }
    foreach ($binding in $bindings) { $binding.component.Name=$binding.name }
}

function Lire-ModuleVba($Module) {
    $count=[int]$Module.CountOfLines
    if ($count -eq 0) { return '' }
    return [string]$Module.Lines(1,$count)
}

function Installer-SourcesVba($Projet, $Entrees, [string]$Racine, [string[]]$RetraitsAutorises=@()) {
    if ($Projet.Protection -ne 0) { throw 'Le projet VBA est verrouille : utilisez les modeles sources non proteges.' }
    $attendus=@($Entrees | ForEach-Object { $_.name })
    # Examiner tous les composants avant le premier DeleteLines/Remove.
    foreach ($component in @($Projet.VBComponents)) {
        if ($component.Name -notin $attendus -and ($component.Type -eq 100 -or $component.Name -notin $RetraitsAutorises)) {
            throw "Composant inconnu conserve : $($component.Name). Inventaire requis avant construction."
        }
    }
    foreach ($entry in $Entrees) {
        $file = Join-Path $Racine $entry.path
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Source VBA absente : $file" }
        $component = $null
        foreach ($candidate in $Projet.VBComponents) {
            if ($candidate.Name -eq $entry.name) { $component = $candidate; break }
        }
        if ($null -eq $component) {
            if ($entry.kind -eq 'form' -and $null -ne $entry.PSObject.Properties['designer']) {
                $component = $Projet.VBComponents.Add(3)
            } elseif ($entry.kind -eq 'module') {
                $component = $Projet.VBComponents.Add(1)
            } else { throw "Formulaire ou classe absent du modele source : $($entry.name)" }
            $component.Name = $entry.name
        }
        if ($null -ne $entry.PSObject.Properties['designer']) {
            Installer-ControlesFormulaire $component $entry.designer
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
    Add-Type -AssemblyName System.IO.Compression
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

# Formulaires nouveaux / boutons ajoutes : uniquement la description du manifeste.
function Installer-ControlesFormulaire($Component,$Spec) {
    if ($Component.Type -ne 3) { throw 'Designer reserve aux formulaires.' }
    $form=$Component.Designer
    $append=$null -ne $Spec.PSObject.Properties['append'] -and [bool]$Spec.append
    $base=0
    if ($append) {
        foreach ($control in $form.Controls) {
            if ($control.Name -notin @($Spec.controls | ForEach-Object {$_.name})) {
                $base=[Math]::Max($base,[double]$control.Top+[double]$control.Height+10)
            }
        }
    } else {
        foreach ($key in @('Caption','Width','Height')) { Definir-ProprieteFormulaire $Component $key $Spec.$key }
        Definir-ProprieteFormulaire $Component 'ShowModal' $false
    }
    foreach ($c in $Spec.controls) {
        if ($c.type -notin @('Forms.Label.1','Forms.TextBox.1','Forms.ListBox.1','Forms.CommandButton.1')) { throw 'Type de controle non autorise.' }
        if ($c.name -notmatch '^[A-Za-z][A-Za-z0-9_]*$') { throw 'Nom de controle invalide.' }
        $control=$null
        foreach ($existing in $form.Controls) { if ($existing.Name -eq $c.name) {$control=$existing;break} }
        if ($null -eq $control) { $control=$form.Controls.Add([string]$c.type,[string]$c.name,$true) }
        $control.Left=[double]$c.left; $control.Top=$base+[double]$c.top
        $control.Width=[double]$c.width; $control.Height=[double]$c.height
        if ($null -ne $c.PSObject.Properties['caption']) { $control.Caption=[string]$c.caption }
        # Police heritee du formulaire ; pas de propriete Font sur le proxy de controle du designer.
        if ($append) { Definir-ProprieteFormulaire $Component 'Height' ([Math]::Max([double]$Component.Properties.Item('Height').Value,[double]$control.Top+[double]$control.Height+35)) }
    }
}

# Les valeurs VBIDE.Property sont des VARIANT ; eviter le cache de conversion
# PowerShell 5.1 qui reutilise le type String de Caption pour Width/Height.
function Definir-ProprieteFormulaire($Component,[string]$Name,[object]$Value) {
    $property=$Component.Properties.Item($Name)
    [void]$property.GetType().InvokeMember('Value',[Reflection.BindingFlags]::SetProperty,$null,$property,[object[]]@($Value))
}