$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'Build/outils_installation.ps1')
. (Join-Path $root 'Build/outils_telechargement.ps1')
function Exiger($condition,$nom) { if (-not $condition) { throw "ECHEC U1 : $nom" };Write-Output "PASS U1 : $nom" }
$ini="[CERFA]`r`nCalageValide=1`r`nImprimante=FICTIVE`r`n[ECG]`r`nDossierGdt=C:\Ancien`r`n[ECG]`r`nDossierGdt=C:\Second`r`n"
$settings=[ordered]@{'poste|profil'='CabinetMedecin';'ecg|dossiergdt'='C:\Fictif'}
$merged=Fusionner-IniPoste $ini $settings
Exiger ($merged -eq (Fusionner-IniPoste $merged $settings)) 'mise a jour INI idempotente'
Exiger ($merged.Contains('CalageValide=1') -and $merged.Contains('Imprimante=FICTIVE')) 'reglages CERFA preserves'
Exiger ([regex]::Matches($merged,'DossierGdt=').Count -eq 1) 'cle remplacee sans doublons'
$temp=Join-Path ([IO.Path]::GetTempPath()) ('cabinet-u1-paquet-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
$archive=''
try {
    $file=Join-Path $temp 'attendu.txt';[IO.File]::WriteAllText($file,'FICTIF')
    $expected=[pscustomobject]@{'attendu.txt'=(Get-FileHash -LiteralPath $file).Hash}
    Exiger ((Isoler-PaquetInvalide $temp $expected) -eq '') 'paquet conforme conserve'
    [IO.File]::WriteAllText((Join-Path $temp 'note-locale.txt'),'NOTE LOCALE FICTIVE A CONSERVER')
    $archive=Isoler-PaquetInvalide $temp $expected
    Exiger (-not (Test-Path -LiteralPath $temp) -and (Test-Path -LiteralPath $archive)) 'paquet invalide isole pour reprise propre'
    Exiger ([IO.File]::ReadAllText((Join-Path $archive 'note-locale.txt')) -eq 'NOTE LOCALE FICTIVE A CONSERVER') 'fichier local conserve'
} finally {
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Recurse -Force }
    if ($archive -and (Test-Path -LiteralPath $archive)) { Remove-Item -LiteralPath $archive -Recurse -Force }
}

# Retour arriere reel des octets et droits, uniquement sur fichiers fictifs.
if([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT){
    $aclRoot=Join-Path ([IO.Path]::GetTempPath()) ('cabinet-u1-acl-'+[guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($aclRoot)
    try{
        $cible=Join-Path $aclRoot 'fictif.txt';$copie=Join-Path $aclRoot 'sauvegarde.txt'
        [IO.File]::WriteAllText($cible,'FICTIF AVANT')
        Proteger-FichierLocal $cible
        $sddl=(Get-Acl -LiteralPath $cible).Sddl
        Proteger-FichierLocal $cible
        Exiger ((Get-Acl -LiteralPath $cible).AreAccessRulesProtected) 'fichier deja protege reste protege apres mise a jour'
        Exiger ((Get-Acl -LiteralPath $cible).Sddl -eq $sddl) 'protection des droits idempotente'
        [IO.File]::Copy($cible,$copie,$false);Proteger-FichierLocal $copie
        Exiger ((Get-Acl -LiteralPath $copie).AreAccessRulesProtected) 'sauvegarde fictive protegee'
        [IO.File]::WriteAllText($cible,'FICTIF APRES')
        # Modifier seulement la DACL : Get-Acl puis Set-Acl avec toutes les sections
        # peut demander SeSecurityPrivilege, absent du compte standard de recette.
        $acl=New-Object Security.AccessControl.FileSecurity
        $acl.SetSecurityDescriptorSddlForm($sddl,[Security.AccessControl.AccessControlSections]::Access)
        $acl.SetAccessRuleProtection($false,$true)
        [IO.File]::SetAccessControl($cible,$acl)
        Exiger (-not (Get-Acl -LiteralPath $cible).AreAccessRulesProtected) 'droits modifies avant retour arriere'
        $refuse=$false
        $sansDacl='O:'+([Security.Principal.WindowsIdentity]::GetCurrent().User.Value)
        try { Restaurer-FichierAvecDroits ([pscustomobject]@{backup=$copie;destination=$cible;sddl=$sansDacl}) }
        catch { $refuse=$_.Exception.Message -like '*sans DACL explicite*' }
        Exiger $refuse 'recu sans DACL explicite refuse'
        Exiger ([IO.File]::ReadAllText($cible) -eq 'FICTIF APRES') 'recu invalide refuse avant modification des octets'
        Restaurer-FichierAvecDroits ([pscustomobject]@{backup=$copie;destination=$cible;sddl=$sddl})
        Exiger ([IO.File]::ReadAllText($cible) -eq 'FICTIF AVANT') 'octets restaures apres echec simule'
        Exiger ((Get-Acl -LiteralPath $cible).Sddl -eq $sddl) 'droits restaures apres echec simule'
    }finally{if(Test-Path -LiteralPath $aclRoot){Remove-Item -LiteralPath $aclRoot -Recurse -Force}}
}
