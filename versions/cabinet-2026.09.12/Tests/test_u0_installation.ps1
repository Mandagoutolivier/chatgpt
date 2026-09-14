$ErrorActionPreference='Stop'
$version=Split-Path $PSScriptRoot -Parent
$repo=Split-Path (Split-Path $version -Parent) -Parent
$temp=Join-Path ([IO.Path]::GetTempPath()) ('u0-legacy-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
try {
    # Une sentinelle preexistante ne doit etre ni modifiee ni accompagnee de nouvelles donnees.
    $sentinel=Join-Path $temp 'Patients.xlsx'
    [IO.File]::WriteAllText($sentinel,'BASE FICTIVE A CONSERVER')
    $hash=(Get-FileHash -LiteralPath $sentinel).Hash
    foreach ($script in Get-ChildItem -LiteralPath (Join-Path $repo 'cabinet-unifie/Build') -Filter '*.ps1') {
        $refused=$false
        try { & $script.FullName -RacineNas $temp -Profil Domicile -Mode Installation } catch { $refused=$true }
        if (-not $refused) { throw "Ancien script execute : $($script.Name)" }
        if ((Get-FileHash -LiteralPath $sentinel).Hash -ne $hash -or @(Get-ChildItem -LiteralPath $temp -Recurse -Force).Count -ne 1) { throw 'Ancien script a modifie la cible.' }
    }
    # Parser tous les nouveaux scripts : detecte aussi les erreurs des fichiers non appeles en CI.
    foreach ($script in Get-ChildItem -LiteralPath (Join-Path $version 'Build') -Filter '*.ps1') {
        $tokens=$null;$errors=$null
        [void][Management.Automation.Language.Parser]::ParseFile($script.FullName,[ref]$tokens,[ref]$errors)
        if (@($errors).Count) { throw "Erreur de syntaxe : $($script.Name) : $errors" }
    }
    Write-Host 'PASS U0 : anciens scripts refuses sans modification ; syntaxe PowerShell valide.'
} finally { Remove-Item -LiteralPath $temp -Recurse -Force }
