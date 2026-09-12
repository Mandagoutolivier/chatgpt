# Test du vrai en-tete CMD avec un payload inoffensif : ni reseau, ni Office, ni registre.
$ErrorActionPreference='Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Test CMD reserve a Windows.' }
$cmdPath=Join-Path $PSScriptRoot 'Demarrer_Installation_Cabinet.cmd'
if (-not (Test-Path -LiteralPath $cmdPath)) { Write-Host 'Lanceur pas encore assemble dans ce commit.';return }
$source=[IO.File]::ReadAllText($cmdPath)
$marker=[regex]::Match($source,'(?m)^# CABINET_POWERSHELL_PAYLOAD_V1\r?$')
if (-not $marker.Success) { throw 'Marqueur du payload absent.' }
$header=$source.Substring(0,$marker.Index+$marker.Length).TrimEnd([char[]]@("`r","`n"))
$payload=@'
[IO.File]::WriteAllText($env:CABINET_TEST_PROOF,'OK')
exit 17
'@
$temp=Join-Path ([IO.Path]::GetTempPath()) ('Cabinet test & espaces '+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
$old=$env:CABINET_TEST_PROOF
try {
    $env:CABINET_TEST_PROOF=Join-Path $temp 'preuve.txt'
    $runner=Join-Path $temp 'Demarrer test.cmd'
    [IO.File]::WriteAllText($runner,($header+"`r`n"+$payload),[Text.Encoding]::ASCII)
    $info=New-Object Diagnostics.ProcessStartInfo
    $info.FileName=$env:ComSpec
    $info.Arguments='/d /c ""'+$runner+'" < NUL"'
    $info.UseShellExecute=$false;$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true
    $process=[Diagnostics.Process]::Start($info)
    try {
        $stdout=$process.StandardOutput.ReadToEnd();$stderr=$process.StandardError.ReadToEnd()
        $process.WaitForExit()
        if ($process.ExitCode -ne 17 -or -not (Test-Path -LiteralPath $env:CABINET_TEST_PROOF)) { throw ("Lanceur CMD en echec : $stdout $stderr") }
        if ([IO.File]::ReadAllText($env:CABINET_TEST_PROOF) -ne 'OK') { throw 'Payload non execute correctement.' }
        Write-Host 'PASS : CMD extrait son payload dans un chemin avec espaces et esperluette, execute PowerShell et propage le code retour.'
    } finally { $process.Dispose() }
} finally {
    $env:CABINET_TEST_PROOF=$old
    Remove-Item -LiteralPath $temp -Recurse -Force
}
