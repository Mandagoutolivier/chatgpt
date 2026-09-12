$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'Build/outils_installation.ps1')
. (Join-Path $root 'Build/outils_assistant.ps1')
. (Join-Path $root 'Build/outils_telechargement.ps1')
Add-Type -AssemblyName System.IO.Compression.FileSystem
$tmp=Join-Path ([IO.Path]::GetTempPath()) ('cabinet-assistant-test-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($tmp)
$count=0
function Verifier([bool]$Condition,[string]$Nom) { if (-not $Condition) { throw "ECHEC $Nom" };$script:count++;Write-Output "PASS : $Nom" }
function Refuser([scriptblock]$Action,[string]$Nom) { $refus=$false;try { & $Action } catch { $refus=$true };Verifier $refus $Nom }
function Ecrire-Zip([string]$Nom,[string[]]$Noms,[bool]$Lien=$false) {
    $path=Join-Path $tmp $Nom
    $zip=[IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $Noms) {
            $entry=$zip.CreateEntry("chatgpt-$commit/versions/cabinet-2026.09.12/$name")
            if ($Lien) { $entry.ExternalAttributes=([int]0xA000 -shl 16) }
            $writer=New-Object IO.StreamWriter($entry.Open());try { $writer.Write('fictif') } finally { $writer.Dispose() }
        }
    } finally { $zip.Dispose() }
    return $path
}
try {
    $statePath=Join-Path $tmp 'etat.json'
    Verifier ($null -eq (Lire-EtatAssistant $statePath)) 'premier lancement sans etat'
    Ecrire-EtatAssistant $statePath ([pscustomobject]@{phase='prepare';profil='Domicile'})
    Ecrire-EtatAssistant $statePath ([pscustomobject]@{phase='valide';profil='Domicile'})
    Verifier ((Lire-EtatAssistant $statePath).phase -eq 'valide') 'reprise depuis etat remplace atomiquement'
    Verifier (@(Get-ChildItem $tmp -Filter '*.tmp').Count -eq 0) 'pas de residu apres ecriture'
    Tester-RegleAccesVba @($null,1)
    Verifier $true 'strategie permissive acceptee'
    Refuser { Tester-RegleAccesVba @(1,0) } 'interdiction administrateur respectee'
    $sources=Join-Path $tmp 'sources'
    foreach ($d in @('Src','Build')) { [void][IO.Directory]::CreateDirectory((Join-Path $sources $d)) }
    $f=Join-Path $sources 'Src/test.bas';[IO.File]::WriteAllText($f,'avant')
    $before=Empreinte-SourcesAssistant $sources
    [IO.File]::WriteAllText($f,'apres')
    Verifier ((Empreinte-SourcesAssistant $sources) -ne $before) 'modification source detectee hors manifeste'
    $binary=Join-Path $tmp 'CabinetUnifie.dotm';[IO.File]::WriteAllText($binary,'avant')
    $state=[pscustomobject]@{dossierPrepare=$tmp;wordCompile=$true;excelCompile=$false;wordHash=(Get-FileHash $binary).Hash;excelHash='';recette=$true}
    Actualiser-CompilationAssistant $state
    Verifier ($state.wordCompile -and $state.recette) 'reprise fichier inchange'
    [IO.File]::WriteAllText($binary,'apres');Actualiser-CompilationAssistant $state
    Verifier (-not $state.wordCompile -and -not $state.recette) 'fichier change invalide compilation et recette'
    $commit='a'*40
    $sample=Join-Path $tmp 'sample';[IO.File]::WriteAllText($sample,'fictif')
    $hash=(Get-FileHash $sample).Hash
    $map=[pscustomobject]@{'Build/test.ps1'=$hash}
    $zip=Ecrire-Zip 'bon.zip' @('Build/test.ps1')
    $dest=Join-Path $tmp 'paquet'
    Extraire-PaquetCabinet $zip $dest $commit $map
    Tester-PaquetCabinet $dest $map
    Verifier $true 'ZIP attendu extrait et controle'
    Refuser { Extraire-PaquetCabinet $zip $dest $commit $map } 'destination existante preservee'
    [IO.File]::WriteAllText((Join-Path $dest 'intrus.ps1'),'x')
    Refuser { Tester-PaquetCabinet $dest $map } 'fichier supplementaire dans cache refuse'
    Remove-Item -LiteralPath (Join-Path $dest 'intrus.ps1')
    [IO.File]::WriteAllText((Join-Path $dest 'Build/test.ps1'),'altere')
    Refuser { Tester-PaquetCabinet $dest $map } 'cache altere refuse'
    $badMap=[pscustomobject]@{'Build/test.ps1'=('0'*64)}
    $badDest=Join-Path $tmp 'invalide'
    Refuser { Extraire-PaquetCabinet $zip $badDest $commit $badMap } 'empreinte ZIP alteree refusee'
    Verifier (-not (Test-Path $badDest)) 'extraction invalide nettoyee'
    foreach ($name in @('../fuite.ps1','Build/../../fuite.ps1','Build/test.ps1:evil','Build\evil.ps1')) {
        $badZip=Ecrire-Zip ([guid]::NewGuid().ToString('N')+'.zip') @($name)
        $badMap=[pscustomobject]@{};$badMap | Add-Member -NotePropertyName $name -NotePropertyValue $hash
        Refuser { Extraire-PaquetCabinet $badZip $badDest $commit $badMap } "chemin ZIP dangereux refuse : $name"
    }
    $duplicate=Ecrire-Zip 'doublon.zip' @('Build/test.ps1','Build/TEST.ps1')
    Refuser { Extraire-PaquetCabinet $duplicate $badDest $commit $map } 'doublon de casse refuse'
    $link=Ecrire-Zip 'lien.zip' @('Build/test.ps1') $true
    Refuser { Extraire-PaquetCabinet $link $badDest $commit $map } 'lien ZIP refuse'
    Refuser { Extraire-PaquetCabinet $zip $badDest ('b'*40) $map } 'mauvaise version du depot refusee'
    # Parse aussi le lanceur autonome lorsqu il existe (second commit de publication).
    $repo=Split-Path (Split-Path $root -Parent) -Parent
    $standalone=Join-Path $repo 'Installateur/Demarrer_Installation_Cabinet.ps1'
    if (Test-Path $standalone) {
        $tokens=$null;$errors=$null
        [void][Management.Automation.Language.Parser]::ParseFile($standalone,[ref]$tokens,[ref]$errors)
        Verifier ($errors.Count -eq 0) 'syntaxe du lanceur autonome'
        $cmd=[IO.File]::ReadAllText((Join-Path $repo 'Installateur/Demarrer_Installation_Cabinet.cmd'))
        $marker=[regex]::Match($cmd,'(?m)^# CABINET_POWERSHELL_PAYLOAD_V1\r?$')
        Verifier ($marker.Success -and $cmd.Substring($marker.Index+$marker.Length).Replace("`r`n","`n").Trim() -ceq [IO.File]::ReadAllText($standalone).Replace("`r`n","`n").Trim()) 'CMD contient exactement le PowerShell auditable'
    }
    Write-Output "$count controles de l assistant reussis."
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force }
