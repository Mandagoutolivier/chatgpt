#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$RacineSources,
    [Parameter(Mandatory=$true)][string]$Sortie
)
Set-StrictMode -Version 2.0
$ErrorActionPreference='Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Qualification Office reservee a Windows.' }
function Read-Host { throw 'Interaction requise : qualification autonome interrompue, conserver le suivi.' }
# Qualification de sources publiques sur copies, aucun NAS, HTTP ou secret reel.
# Executer dans un PowerShell neuf avec Word et Excel fermes.
function Assert-LocalQualification([string]$Path,[switch]$Directory) {
    if ($Path -notmatch '^[A-Za-z]:\\' -or [IO.Path]::GetFullPath($Path) -ine $Path) { throw 'Un chemin local absolu est requis.' }
    $item=Get-Item -LiteralPath $Path -Force
    if (($item -is [IO.DirectoryInfo]) -ne [bool]$Directory) { throw 'Type de chemin inattendu.' }
    $drive=[IO.DriveInfo]::new([IO.Path]::GetPathRoot($Path))
    if ($drive.DriveType -ne [IO.DriveType]::Fixed) { throw 'Disque local fixe requis.' }
    while ($null -ne $item) {
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Chemin redirige refuse.' }
        if ($item -is [IO.DirectoryInfo]) { $item=$item.Parent } else { $item=$item.Directory }
    }
}
function Assert-StartupVide([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return }
    Assert-LocalQualification $Path -Directory
    if (@(Get-ChildItem -LiteralPath $Path -Force -File).Count -gt 0) { throw 'Un dossier STARTUP/XLSTART contient des fichiers : qualification interrompue avant Office.' }
}
$RacineSources=[IO.Path]::GetFullPath($RacineSources)
$Sortie=[IO.Path]::GetFullPath($Sortie)
# La suite U2 ajoute 164 caracteres (GUIDs, SHA256, extension temporaire).
# FileSystemObject VBA utilise encore MAX_PATH ; refuser avant tout lancement COM.
if ($Sortie.Length -gt 90) { throw 'Sortie trop longue : choisir au maximum 90 caracteres pour rester sous MAX_PATH avec les fichiers temporaires VBA.' }
Assert-LocalQualification $RacineSources -Directory
Assert-LocalQualification (Split-Path $Sortie -Parent) -Directory
if (Test-Path -LiteralPath $Sortie) { throw 'La sortie doit etre neuve.' }
if (Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue) { throw 'Office deja ouvert : aucune instance existante ne sera utilisee.' }
# Le demarrage COM ne garantit pas l heritage APPDATA : l instrumentation VBA
# remplace egalement tous les chemins de poste par des litteraux locaux de fixture.
$originalAppData=[Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)
if ([string]::IsNullOrWhiteSpace($originalAppData)) { throw 'Dossier ApplicationData de la session introuvable.' }
$normal=Join-Path $originalAppData 'Microsoft\Templates\Normal.dotm'
$normalExists=Test-Path -LiteralPath $normal
$normalHash='';$normalBackup=$null
Add-Type -AssemblyName System.IO.Compression.FileSystem
if ($normalExists) {
    Assert-LocalQualification $normal
    $zip=[IO.Compression.ZipFile]::OpenRead($normal)
    try { if ($null -ne $zip.GetEntry('word/vbaProject.bin')) { throw 'Normal contient du VBA : demarrage automatique non autorise pour cette qualification.' } }
    finally { $zip.Dispose() }
    $normalHash=(Get-FileHash -LiteralPath $normal -Algorithm SHA256).Hash
}
$startup=@((Join-Path $originalAppData 'Microsoft\Word\STARTUP'),(Join-Path $originalAppData 'Microsoft\Excel\XLSTART'))
foreach ($hostName in @('Word','Excel')) {
    foreach ($baseKey in @('HKCU:\Software','HKLM:\Software','HKCU:\Software\WOW6432Node','HKLM:\Software\WOW6432Node')) {
        $path="$baseKey\Microsoft\Office\16.0\$hostName\Options"
        $key=Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
        if ($null -ne $key) {
            foreach ($name in @($key.GetValueNames())) {
                if ($name -match '^(STARTUP-PATH|AltStartup|OPEN\d*)$' -and [string]$key.GetValue($name)) {
                    throw 'Office possede un lancement ou un dossier de demarrage personnalise : revue requise.'
                }
            }
        }
    }
}
foreach ($baseKey in @('HKCU:\Software','HKLM:\Software','HKCU:\Software\WOW6432Node','HKLM:\Software\WOW6432Node')) {
    $key=Get-Item -LiteralPath "$baseKey\Microsoft\Office\16.0\Common\General" -ErrorAction SilentlyContinue
    if ($null -ne $key -and [string]$key.GetValue('UserTemplates')) { throw 'Dossier de modeles Word personnalise : revue requise.' }
}
$programFolders=@([Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles),[Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFilesX86))|Where-Object {$_}|Select-Object -Unique
if (@($programFolders).Count -eq 0) { throw 'Dossiers ProgramFiles introuvables : controle STARTUP impossible.' }
foreach ($programRoot in $programFolders) {
    if (-not $programRoot) { continue }
    foreach ($relative in @('Microsoft Office\Office16','Microsoft Office\root\Office16')) {
        $startup+=Join-Path (Join-Path $programRoot $relative) 'STARTUP'
        $startup+=Join-Path (Join-Path $programRoot $relative) 'XLSTART'
    }
}
foreach ($baseKey in @('HKLM:\Software','HKLM:\Software\WOW6432Node')) {
    $install=Get-ItemProperty -LiteralPath "$baseKey\Microsoft\Office\16.0\Common\InstallRoot" -Name Path -ErrorAction SilentlyContinue
    if ($null -ne $install -and [string]$install.Path) {
        $startup+=Join-Path ([string]$install.Path) 'STARTUP'
        $startup+=Join-Path ([string]$install.Path) 'XLSTART'
    }
}
foreach ($path in $startup) { Assert-StartupVide $path }
[void][IO.Directory]::CreateDirectory($Sortie)
$acl=[Security.AccessControl.DirectorySecurity]::new()
$acl.SetAccessRuleProtection($true,$false)
foreach ($sid in @([Security.Principal.WindowsIdentity]::GetCurrent().User.Value,'S-1-5-18','S-1-5-32-544')) {
    $rule=[Security.AccessControl.FileSystemAccessRule]::new([Security.Principal.SecurityIdentifier]::new($sid),[Security.AccessControl.FileSystemRights]::FullControl,[Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit',[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Allow)
    $acl.AddAccessRule($rule)
}
[IO.Directory]::SetAccessControl($Sortie,$acl)
$report=[ordered]@{Statut='EN_COURS';CopiesInstrumentees=$true;SuitesTerminees=$false;OfficeFerme=$false;NormalRestaure=$false;SecuriteRestauree=$false;ActivationEffectuee=$false;NASContacte=$false;ImpressionReelle=$false;Erreur=$null}
$before=@();$environment=@{};$cleanupError=$null
try {
    if ($normalExists) {
        $normalBackup=Join-Path $Sortie 'Normal-original.dotm'
        [IO.File]::Copy($normal,$normalBackup,$false)
        if ((Get-FileHash $normalBackup -Algorithm SHA256).Hash -ne $normalHash) { throw 'Sauvegarde Normal non conforme.' }
    }
    # Une modification restrictive avant le premier COM protege contre les macros
    # de demarrage. Aucun parametre de strategie administrateur n est modifie.
    foreach ($hostName in @('Word','Excel')) {
        $path="HKCU:\Software\Microsoft\Office\16.0\$hostName\Security"
        $key=Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
        $exists=$null -ne $key -and 'VBAWarnings' -in @($key.GetValueNames())
        $value=$null;$kind=$null
        if ($exists) { $value=$key.GetValue('VBAWarnings');$kind=[string]$key.GetValueKind('VBAWarnings') }
        $accessExists=$null -ne $key -and 'AccessVBOM' -in @($key.GetValueNames())
        $accessValue=$null;$accessKind=$null
        if ($accessExists) { $accessValue=$key.GetValue('AccessVBOM');$accessKind=[string]$key.GetValueKind('AccessVBOM') }
        $before+=@([pscustomobject]@{path=$path;existed=$exists;value=$value;kind=$kind;accessExisted=$accessExists;accessValue=$accessValue;accessKind=$accessKind})
        [IO.File]::WriteAllText((Join-Path $Sortie 'securite-macros-a-restaurer.json'),($before|ConvertTo-Json -Depth 4),[Text.UTF8Encoding]::new($false))
        if ($null -eq $key) { [void](New-Item -Path $path -Force) }
        New-ItemProperty -LiteralPath $path -Name VBAWarnings -Value 4 -PropertyType DWord -Force|Out-Null
    }
    foreach ($name in @('APPDATA','LOCALAPPDATA','TEMP','TMP','CABINET_QUALIFICATION_ISOLEE')) {
        $environment[$name]=[Environment]::GetEnvironmentVariable($name,'Process')
    }
    foreach ($pair in @(@('APPDATA','AppData'),@('LOCALAPPDATA','LocalAppData'),@('TEMP','Temp'),@('TMP','Temp'))) {
        $path=Join-Path (Join-Path $Sortie 'Environnement') $pair[1]
        [void][IO.Directory]::CreateDirectory($path)
        [Environment]::SetEnvironmentVariable($pair[0],$path,'Process')
    }
    $env:CABINET_QUALIFICATION_ISOLEE=$Sortie
    & (Join-Path $PSScriptRoot 'Tester_U1_Office.ps1') -Sortie $Sortie -RecetteU2 -RacineSources $RacineSources -HorsReseau
    $report.SuitesTerminees=$true
} catch { $report.Erreur=$_.Exception.Message }
finally {
    foreach ($name in $environment.Keys) { [Environment]::SetEnvironmentVariable($name,$environment[$name],'Process') }
    # Jamais de kill : si Office reste actif, conserver les journaux de reprise.
    for ($i=0;$i -lt 120 -and (Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue);$i++) { Start-Sleep -Milliseconds 250 }
    $report.OfficeFerme=-not [bool](Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue)
    if ($report.OfficeFerme) {
        try {
            if ($normalExists) {
                if ($null -eq $normalBackup) { throw 'Sauvegarde Normal absente.' }
                if (-not (Test-Path -LiteralPath $normal) -or (Get-FileHash -LiteralPath $normal -Algorithm SHA256).Hash -ne $normalHash) { [IO.File]::Copy($normalBackup,$normal,$true) }
                if ((Get-FileHash -LiteralPath $normal -Algorithm SHA256).Hash -ne $normalHash) { throw 'Normal non restaure.' }
            } elseif (Test-Path -LiteralPath $normal) {
                [IO.File]::Copy($normal,(Join-Path $Sortie 'Normal-cree-pendant-qualification.dotm'),$false)
                [IO.File]::Delete($normal)
            }
            $report.NormalRestaure=$true
            foreach ($item in $before) {
                $key=Get-Item -LiteralPath $item.path
                if ([int]$key.GetValue('VBAWarnings') -ne 4) { throw 'Parametre macros change pendant la qualification ; restauration automatique interrompue.' }
                if ($item.existed) { New-ItemProperty -LiteralPath $item.path -Name VBAWarnings -Value $item.value -PropertyType $item.kind -Force|Out-Null }
                else { Remove-ItemProperty -LiteralPath $item.path -Name VBAWarnings }
                $key=Get-Item -LiteralPath $item.path
                $accessPresent='AccessVBOM' -in @($key.GetValueNames())
                if ($accessPresent -ne $item.accessExisted -or ($accessPresent -and
                    ($key.GetValue('AccessVBOM') -ne $item.accessValue -or [string]$key.GetValueKind('AccessVBOM') -cne $item.accessKind))) { throw 'AccessVBOM non restaure a sa valeur initiale.' }
            }
            $report.SecuriteRestauree=-not (Test-Path -LiteralPath (Join-Path $Sortie 'acces-vba-a-restaurer.json'))
            if (-not $report.SecuriteRestauree) { throw 'Journal AccessVBOM encore present.' }
            Remove-Item -LiteralPath (Join-Path $Sortie 'securite-macros-a-restaurer.json') -ErrorAction SilentlyContinue
        } catch { $cleanupError=$_.Exception.Message }
    } else { $cleanupError='Office encore actif ; aucun processus termine de force. Restaurer avec les journaux apres fermeture.' }
    if ($cleanupError) { $report.Erreur=([string]$report.Erreur+' | Nettoyage : '+$cleanupError).Trim(' ','|') }
    $report.Statut=if ($report.SuitesTerminees -and $report.OfficeFerme -and $report.NormalRestaure -and $report.SecuriteRestauree -and -not $report.Erreur) {'SUCCESS'} else {'ECHEC'}
    $report['DateUTC']=[DateTime]::UtcNow.ToString('o')
    [IO.File]::WriteAllText((Join-Path $Sortie 'qualification-office-isolee.json'),($report|ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))
    $report|ConvertTo-Json -Depth 5
}
if ($report.Statut -ne 'SUCCESS') { throw 'Qualification isolee interrompue. Conserver les journaux de suivi.' }
