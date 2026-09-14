[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Sortie,
    [string]$DossierDemarrageWord='',
    [switch]$TesterService
)
$ErrorActionPreference='Stop'
$local=Join-Path $env:APPDATA 'CabinetCardio'
if (-not $DossierDemarrageWord) { $DossierDemarrageWord=Join-Path $env:APPDATA 'Microsoft\Word\STARTUP' }
$report=[ordered]@{dateUtc=[DateTime]::UtcNow.ToString('o');poste=$env:COMPUTERNAME;racineNas='';
    startupWordControle=$DossierDemarrageWord;startupPersonnaliseAConfirmer=$true;modeles=@();installations=@();service=@{};
    officeEnCours=@(Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue | ForEach-Object {$_.ProcessName})}
$path=Join-Path $local 'chemin.txt'
if (Test-Path -LiteralPath $path) { $report.racineNas=[IO.File]::ReadAllText($path).Trim() }
$folders=@($DossierDemarrageWord,(Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CabinetCardio'))
foreach ($folder in $folders) {
    if (Test-Path -LiteralPath $folder -PathType Container) {
        foreach ($file in Get-ChildItem -LiteralPath $folder -File | Where-Object {$_.Extension -in @('.dotm','.xlsm')}) {
            $report.modeles+=@{chemin=$file.FullName;octets=$file.Length;sha256=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash}
        }
    }
}
foreach ($receipt in Get-ChildItem -LiteralPath $local -Filter 'installation-*.json' -File -ErrorAction SilentlyContinue) {
    $r=Get-Content -LiteralPath $receipt.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $report.installations+=@{profil=$r.profil;release=$r.release;commitSources=$r.commitSources;binaires=$r.binaires}
}
if ($TesterService) {
    $token=$null
    try {
        $url=[IO.File]::ReadAllText((Join-Path $local 'service.url')).Trim()
        $uri=[Uri]$url
        if ($uri.Scheme -ne 'https' -or $uri.UserInfo) { throw 'URL HTTPS requise.' }
        $token=[IO.File]::ReadAllText((Join-Path $local 'service.token')).Trim()
        $r=Invoke-RestMethod -Method Post -Uri ($url.TrimEnd('/')+'/v1/rpc') -Headers @{Authorization=('Bearer '+$token)} -ContentType 'application/json' -Body '{"operation":"whoami","params":{}}' -MaximumRedirection 0 -TimeoutSec 15
        $report.service=@{joignable=$true;version=$r.result.version;revision=$r.result.revision;schema=$r.result.schema;protocole=$r.result.protocole;roles=$r.result.roles}
    } catch { $report.service=@{joignable=$false;erreur='Connexion non validee ; aucun secret inclus dans le rapport.'} }
    finally { $token=$null }
}
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Sortie -Encoding UTF8
Write-Host "Inventaire cree : $Sortie. Aucun modele ouvert, aucune configuration modifiee."
