[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$Sortie)
$ErrorActionPreference='Stop'
# Seulement sur un creneau autorise, Office ferme ; copies dans une sortie neuve.
& (Join-Path $PSScriptRoot 'Tester_U1_Office.ps1') -Sortie $Sortie -RecetteU2
