$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'Build/outils_construction.ps1')
$count=0
function Empreinte-EntreeZip($Archive,[string]$Nom) {
    $stream=$Archive.GetEntry($Nom).Open();$sha=[Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($stream)) }
    finally { $stream.Dispose();$sha.Dispose() }
}
function Verifier([bool]$Condition,[string]$Nom) {
    if (-not $Condition) { throw "ECHEC : $Nom" }
    Write-Output "PASS : $Nom"
    $script:count++
}
foreach ($file in Get-ChildItem -LiteralPath $root -Recurse -Filter '*.ps1') {
    $tokens=$null;$errors=$null
    [void][Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$errors)
    Verifier ($errors.Count -eq 0) ("syntaxe " + $file.Name + ' ' + ($errors -join ' '))
}
$tmp=Join-Path ([IO.Path]::GetTempPath()) ('cabinet-tests-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($tmp)
try {
    $source=Join-Path $root 'ModelesSource/ModeleCourrierChatGPT_PROD(6).dotm'
    $model=Join-Path $tmp 'modele.dotm'
    [IO.File]::Copy($source,$model)
    # Un processus neuf ne doit pas dependre d un premier OpenRead pour charger ZipArchiveMode.
    $bootstrapModel=Join-Path $tmp 'demarrage-frais.dotm'
    [IO.File]::Copy($source,$bootstrapModel)
    $bootstrapScript=Join-Path $tmp 'demarrage-frais.ps1'
    $helper=(Join-Path $root 'Build/outils_construction.ps1').Replace("'","''")
    $ruban=(Join-Path $root 'Build/ruban_unifie.xml').Replace("'","''")
    $bootstrapCode="`$ErrorActionPreference='Stop'`r`n. '$helper'`r`nInstaller-Ruban '$($bootstrapModel.Replace("'","''"))' '$ruban'"
    [IO.File]::WriteAllText($bootstrapScript,$bootstrapCode,(New-Object Text.UTF8Encoding($true)))
    & (Get-Process -Id $PID).Path -NoLogo -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File $bootstrapScript
    Verifier ($LASTEXITCODE -eq 0) 'injection du ruban dans un processus PowerShell neuf'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip=[IO.Compression.ZipFile]::OpenRead($model)
    try {
        [xml]$beforeTypes=Lire-EntreeZip $zip '[Content_Types].xml'
        [xml]$beforeRels=Lire-EntreeZip $zip '_rels/.rels'
        $vbaBefore=Empreinte-EntreeZip $zip 'word/vbaProject.bin'
    } finally { $zip.Dispose() }
    Installer-Ruban $model (Join-Path $root 'Build/ruban_unifie.xml')
    Installer-Ruban $model (Join-Path $root 'Build/ruban_unifie.xml')
    $zip=[IO.Compression.ZipFile]::OpenRead($model)
    try {
        [xml]$afterTypes=Lire-EntreeZip $zip '[Content_Types].xml'
        [xml]$afterRels=Lire-EntreeZip $zip '_rels/.rels'
        foreach ($node in $beforeTypes.DocumentElement.ChildNodes) {
            $matches=@($afterTypes.DocumentElement.ChildNodes | Where-Object {
                $_.LocalName -eq $node.LocalName -and $_.GetAttribute('PartName') -eq $node.GetAttribute('PartName') -and
                $_.GetAttribute('Extension') -eq $node.GetAttribute('Extension') -and $_.GetAttribute('ContentType') -eq $node.GetAttribute('ContentType')
            })
            Verifier ($matches.Count -eq 1) ('type OOXML conserve '+$node.GetAttribute('ContentType'))
        }
        foreach ($node in $beforeRels.DocumentElement.ChildNodes) {
            if ($node.GetAttribute('Type') -notmatch '/ui/extensibility$') {
                $matches=@($afterRels.DocumentElement.ChildNodes | Where-Object {
                    $_.GetAttribute('Id') -eq $node.GetAttribute('Id') -and $_.GetAttribute('Type') -eq $node.GetAttribute('Type') -and $_.GetAttribute('Target') -eq $node.GetAttribute('Target')
                })
                Verifier ($matches.Count -eq 1) ('relation OOXML conservee '+$node.GetAttribute('Id'))
            }
        }
        $ui=@($afterRels.DocumentElement.ChildNodes | Where-Object { $_.GetAttribute('Type') -match '/ui/extensibility$' })
        Verifier ($ui.Count -eq 1) 'injection du ruban idempotente'
        Verifier ((Empreinte-EntreeZip $zip 'word/vbaProject.bin') -eq $vbaBefore) 'SHA256 du projet binaire conserve par injection OOXML'
        Verifier ($null -ne $zip.GetEntry('customUI/cabinet.xml')) 'ruban present'
    } finally { $zip.Dispose() }
    $manifest=Lire-Manifeste $root
    Verifier-ModeleSource $source 'ModeleCourrierChatGPT_PROD(6).dotm' $manifest
    $rejected=$false
    try { Verifier-ModeleSource $model 'ModeleCourrierChatGPT_PROD(6).dotm' $manifest } catch { $rejected=$true }
    Verifier $rejected 'source modifiee refusee'
    $vba=Lire-CodeVba (Join-Path $root 'Src/Word/modCourrier.bas')
    Verifier (-not ($vba -match '(?m)^Attribute ')) 'metadonnees VBA exclues du CodeModule'
    $out=Join-Path $tmp 'destination.txt';$inputFile=Join-Path $tmp 'nouveau.txt'
    [IO.File]::WriteAllText($out,'ancien');[IO.File]::WriteAllText($inputFile,'nouveau')
    Publier-FichierConstruit $inputFile $out
    Verifier ([IO.File]::ReadAllText($out) -eq 'nouveau') 'publication atomique du fichier construit'
    $backups=@(Get-ChildItem -LiteralPath $tmp -Filter 'destination.txt.avant-*')
    Verifier ($backups.Count -eq 1 -and [IO.File]::ReadAllText($backups[0].FullName) -eq 'ancien') 'version precedente sauvegardee'
    Write-Output "$count controles de construction reussis. Aucun test COM Office effectue."
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force }
