Set-StrictMode -Version 2.0
$ErrorActionPreference='Stop'
function Empreintes-Sources([string]$Racine) {
    $result=[ordered]@{}
    foreach ($folder in @('Src','Build')) {
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $Racine $folder) -Recurse -File | Sort-Object FullName) {
            $relative=$file.FullName.Substring($Racine.Length).TrimStart([char[]]@('\','/')).Replace('\','/')
            $result[$relative]=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        }
    }
    return $result
}
function Ecrire-Preparation([string]$Dossier,[string]$Profil,[string]$Racine) {
    $binaries=[ordered]@{}
    foreach ($name in @('CabinetUnifie.dotm','Cabinet.xlsm','sqlite3.exe')) {
        $path=Join-Path $Dossier $name
        if (Test-Path -LiteralPath $path) { $binaries[$name]=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }
    }
    $receipt=[ordered]@{version='2026.09.12';profil=$Profil;sources=(Empreintes-Sources $Racine);binaires=$binaries;compilationOffice='A effectuer sur ce PC'}
    $receipt | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $Dossier 'preparation.json') -Encoding UTF8
}
function Verifier-Preparation([string]$Dossier,[string]$Profil,[string]$Racine) {
    $receipt=Get-Content -LiteralPath (Join-Path $Dossier 'preparation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($receipt.version -ne '2026.09.12' -or $receipt.profil -ne $Profil) { throw 'Preparation absente ou incompatible avec ce profil.' }
    $current=Empreintes-Sources $Racine
    if ($current.Count -ne @($receipt.sources.PSObject.Properties).Count) { throw 'La liste des sources a change depuis la preparation.' }
    foreach ($item in $receipt.sources.PSObject.Properties) {
        if ($current[$item.Name] -ne $item.Value) { throw "Source modifiee depuis la preparation : $($item.Name)" }
    }
    $required=@()
    if ($Profil -in @('Domicile','CabinetMedecin')) { $required+=@('CabinetUnifie.dotm','sqlite3.exe') }
    if ($Profil -in @('Domicile','CabinetSecretariat')) { $required+='Cabinet.xlsm' }
    foreach ($name in $required) {
        $property=$receipt.binaires.PSObject.Properties[$name]
        if ($null -eq $property -or (Get-FileHash -LiteralPath (Join-Path $Dossier $name) -Algorithm SHA256).Hash -ne $property.Value) { throw "Binaire prepare absent ou modifie : $name" }
    }
}
function Proteger-FichierLocal([string]$Path) {
    $sid=[Security.Principal.WindowsIdentity]::GetCurrent().User
    $acl=New-Object Security.AccessControl.FileSecurity
    $acl.SetAccessRuleProtection($true,$false)
    foreach ($identity in @($sid,(New-Object Security.Principal.SecurityIdentifier('S-1-5-18')),(New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')))) {
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($identity,'FullControl','Allow')))
    }
    Set-Acl -LiteralPath $Path -AclObject $acl
}
