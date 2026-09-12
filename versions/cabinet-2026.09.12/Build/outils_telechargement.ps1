Set-StrictMode -Version 2.0
$ErrorActionPreference='Stop'

function Tester-PaquetCabinet([string]$Dossier,$Empreintes) {
    if (-not (Test-Path -LiteralPath $Dossier -PathType Container)) { throw 'Paquet absent.' }
    if ((Get-Item -LiteralPath $Dossier).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Lien de dossier refuse.' }
    $entries=@(Get-ChildItem -LiteralPath $Dossier -Recurse -Force)
    foreach ($file in $entries) {
        if ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Lien dans le paquet refuse.' }
    }
    $files=@($entries | Where-Object { -not $_.PSIsContainer })
    if ($files.Count -ne @($Empreintes.PSObject.Properties).Count) { throw 'Liste des fichiers du paquet differente de la version attendue.' }
    foreach ($file in $files) {
        $relative=$file.FullName.Substring($Dossier.TrimEnd([char[]]@('\','/')).Length+1).Replace('\','/')
        $p=$Empreintes.PSObject.Properties[$relative]
        if ($null -eq $p -or (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash -ne $p.Value) { throw "Fichier absent du manifeste ou modifie : $relative" }
    }
}

function Extraire-PaquetCabinet([string]$Archive,[string]$Destination,[string]$Commit,$Empreintes) {
    if ($Commit -notmatch '^[a-f0-9]{40}$') { throw 'Identifiant de version invalide.' }
    if (Test-Path -LiteralPath $Destination) { throw 'Le dossier de destination doit etre nouveau.' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip=$null;$created=$false
    try {
        $zip=[IO.Compression.ZipFile]::OpenRead($Archive)
        if ($zip.Entries.Count -gt 10000) { throw 'Archive trop volumineuse.' }
        $prefix="chatgpt-$Commit/versions/cabinet-2026.09.12/"
        $selected=New-Object 'System.Collections.Generic.List[object]'
        $seen=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        [long]$total=0
        foreach ($entry in $zip.Entries) {
            if (-not $entry.FullName.StartsWith($prefix,[StringComparison]::Ordinal)) { continue }
            $name=$entry.FullName.Substring($prefix.Length)
            if (-not $name -or $name.EndsWith('/')) { continue }
            $segments=$name.Split('/')
            if ($name -match '[\\:\x00-\x1f]' -or @($segments | Where-Object { $_ -in @('','..','.') -or $_ -match '[. ]$' }).Count) { throw 'Chemin invalide dans le ZIP.' }
            if (-not $seen.Add($name)) { throw 'Doublon dans le ZIP.' }
            $kind=($entry.ExternalAttributes -shr 16) -band 0xF000
            if ($kind -ne 0 -and $kind -ne 0x8000) { throw 'Entree non reguliere dans le ZIP.' }
            if ($null -eq $Empreintes.PSObject.Properties[$name]) { throw "Fichier non attendu dans le ZIP : $name" }
            $total+=$entry.Length
            if ($entry.Length -gt 100MB -or $total -gt 300MB) { throw 'Taille decompressee excessive.' }
            $selected.Add([pscustomobject]@{entry=$entry;name=$name})
        }
        if ($selected.Count -ne @($Empreintes.PSObject.Properties).Count) { throw 'ZIP incomplet ou version incorrecte. Telechargez le lien indique par le lanceur.' }
        [void][IO.Directory]::CreateDirectory($Destination);$created=$true
        foreach ($item in $selected) {
            $target=Join-Path $Destination $item.name
            [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
            $inputStream=$item.entry.Open();$outputStream=$null
            try {
                $outputStream=[IO.File]::Open($target,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
                $inputStream.CopyTo($outputStream)
            } finally { $inputStream.Dispose();if ($null -ne $outputStream) { $outputStream.Dispose() } }
        }
        Tester-PaquetCabinet $Destination $Empreintes
    } catch {
        if ($created -and (Test-Path -LiteralPath $Destination)) { Remove-Item -LiteralPath $Destination -Recurse -Force }
        throw
    } finally { if ($null -ne $zip) { $zip.Dispose() } }
}
