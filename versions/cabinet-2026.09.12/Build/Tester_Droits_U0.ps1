[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$RacineRecette,
    [Parameter(Mandatory=$true)][string]$Docx,
    [Parameter(Mandatory=$true)][string]$Pdf,
    [Parameter(Mandatory=$true)][string]$Sortie
)
$ErrorActionPreference='Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Controle des droits SMB depuis Windows requis.' }
if (-not (Test-Path -LiteralPath (Join-Path $RacineRecette 'U0-RECETTE-SEULEMENT.txt'))) { throw 'Partage de recette non identifie.' }
$root=[IO.Path]::GetFullPath($RacineRecette).TrimEnd('\')
$archives=Join-Path $root 'Documents'
foreach ($path in @($Docx,$Pdf)) {
    $full=[IO.Path]::GetFullPath($path)
    if (-not $full.StartsWith($archives+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Fichier hors des archives de recette.' }
}
# Demander un droit sans ecrire, supprimer ni renommer le fichier.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CabinetU0Acl {
  [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  static extern IntPtr CreateFile(string p, uint access, uint share, IntPtr sa, uint mode, uint flags, IntPtr t);
  [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
  public static int Probe(string path, uint access, bool directory) {
    IntPtr h=CreateFile(path,access,7,IntPtr.Zero,3,directory ? 0x02000000u : 0u,IntPtr.Zero);
    if(h==new IntPtr(-1)) return Marshal.GetLastWin32Error();
    CloseHandle(h); return 0;
  }
}
'@
$checks=@()
foreach ($path in @($Docx,$Pdf)) {
    $read=[CabinetU0Acl]::Probe($path,2147483648,$false)
    $write=[CabinetU0Acl]::Probe($path,1073741824,$false)
    $delete=[CabinetU0Acl]::Probe($path,65536,$false)
    $checks+=@{fichier=[IO.Path]::GetFileName($path);lecture=$read;ecriture=$write;suppression=$delete;
        valide=($read -eq 0 -and $write -eq 5 -and $delete -eq 5)}
}
$dirWrite=[CabinetU0Acl]::Probe($archives,2,$true)
$dirDelete=[CabinetU0Acl]::Probe($archives,65536,$true)
$childDelete=[CabinetU0Acl]::Probe($archives,64,$true)
$rootChildDelete=[CabinetU0Acl]::Probe($root,64,$true)
$valid=(@($checks | Where-Object {-not $_.valide}).Count -eq 0 -and $dirWrite -eq 5 -and $dirDelete -eq 5 -and $childDelete -eq 5 -and $rootChildDelete -eq 5)
@{dateUtc=[DateTime]::UtcNow.ToString('o');poste=$env:COMPUTERNAME;fichiers=$checks;
    creationArchive=$dirWrite;suppressionDossier=$dirDelete;suppressionEnfants=$childDelete;suppressionDepuisParent=$rootChildDelete;valide=$valid} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Sortie -Encoding UTF8
if (-not $valid) { throw 'Droits de recette non valides ou indetermines. Codes 0=accorde, 5=refuse ; tout autre code doit etre examine.' }
Write-Host 'Lecture accordee ; ecriture et suppression refusees pour ces archives et leurs dossiers. Aucun contenu modifie.'
