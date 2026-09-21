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
    foreach ($name in @('CabinetUnifie.dotm','Cabinet.xlsm')) {
        $path=Join-Path $Dossier $name
        if (Test-Path -LiteralPath $path) { $binaries[$name]=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }
    }
    $receipt=[ordered]@{version='2026.09.12';profil=$Profil;sources=(Empreintes-Sources $Racine);binaires=$binaries;compilationOffice='A effectuer sur ce PC'}
    $receipt['release']='2026.09.21-u2c'
    $receipt['commitSources']=$env:CABINET_SOURCE_COMMIT
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
    if ($Profil -in @('Domicile','CabinetMedecin')) { $required+=@('CabinetUnifie.dotm') }
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
    # DACL seule, y compris sur un fichier deja protege lors d une mise a jour.
    [IO.File]::SetAccessControl($Path,$acl)
}

function Fusionner-IniPoste([string]$Ancien,[System.Collections.IDictionary]$Valeurs) {
    $lignes=New-Object 'System.Collections.Generic.List[string]'
    $ecrits=@{};$section=''
    foreach ($line in ($Ancien -split '\r?\n')) {
        if ($line -match '^\s*\[([^]]+)\]') { $section=$Matches[1].ToLowerInvariant() }
        if ($line -match '^\s*([^;#=]+?)\s*=(.*)$') {
            $key=$section+'|'+$Matches[1].Trim().ToLowerInvariant()
            if ($Valeurs.Contains($key)) {
                if (-not $ecrits.ContainsKey($key)) { $lignes.Add($Matches[1].Trim()+'='+[string]$Valeurs[$key]);$ecrits[$key]=$true }
                continue
            }
        }
        $lignes.Add($line)
    }
    foreach ($key in $Valeurs.Keys) {
        if (-not $ecrits.ContainsKey($key)) {
            $parts=$key.Split('|');$lignes.Add('['+$parts[0]+']');$lignes.Add($parts[1]+'='+[string]$Valeurs[$key])
        }
    }
    return ($lignes -join "`r`n").TrimEnd()+"`r`n"
}

function Restaurer-FichierAvecDroits($Item) {
    if ($Item.backup) {
        $saved=$null
        if ($Item.PSObject.Properties['sddl'] -and $Item.sddl) {
            # Get-Acl sauvegarde les droits d acces, pas l audit (SACL).
            # La surcharge sans sections marque pourtant aussi Audit comme modifie
            # et demande SeSecurityPrivilege lors de Set-Acl sur un compte standard.
            $saved=New-Object Security.AccessControl.RawSecurityDescriptor([string]$Item.sddl)
            if ($null -eq $saved.DiscretionaryAcl) { throw 'Sauvegarde des droits sans DACL explicite : restauration interrompue.' }
        }
        [IO.File]::Copy($Item.backup,$Item.destination,$true)
        if ($null -ne $saved) {
            $current=Get-Acl -LiteralPath $Item.destination
            $sections=[Security.AccessControl.AccessControlSections]::Access
            $sidType=[Security.Principal.SecurityIdentifier]
            if ($null -ne $saved.Owner -and $saved.Owner.Value -ne $current.GetOwner($sidType).Value) {
                $sections=$sections -bor [Security.AccessControl.AccessControlSections]::Owner
            }
            if ($null -ne $saved.Group -and $saved.Group.Value -ne $current.GetGroup($sidType).Value) {
                $sections=$sections -bor [Security.AccessControl.AccessControlSections]::Group
            }
            $acl=New-Object Security.AccessControl.FileSecurity
            $acl.SetSecurityDescriptorSddlForm([string]$Item.sddl,$sections)
            # Persister seulement les sections marquees modifiees. Le fournisseur
            # Set-Acl de Windows PowerShell 5.1 peut redemander l audit.
            [IO.File]::SetAccessControl([string]$Item.destination,$acl)
        }
    } elseif ([IO.File]::Exists($Item.destination)) { [IO.File]::Delete($Item.destination) }
}
