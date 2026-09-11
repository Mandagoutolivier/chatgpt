# CabinetCardio unifie — branche de travail

## Identification du modele historique

`ModeleCourrierChatGPT_PROD(6).dotm` a pour SHA-256
`A8CCB702E8DEFC9FCE19B9AFB968C4F7CD5B25D59DA562DBA172BC6B51B85524`.
Il est identique a `PROD(4)` et porte les marqueurs
`DOMICILE-DOUBLE-API-DETAILLE-1` et `DOMICILE-GRAS-1Z`.

Ce modele est la branche fonctionnelle **R7 / double API detaillee**, utilisee
comme point de depart de R12. Ce n'est pas le R12 complet « profils 3 API » :
les modules `modProfilsExamensR12`, `modMoteurLettresProfilsR12`,
`modRevisionFinaleR12` et `modZonesMedicalesR12` en sont absents. Ce n'est pas
non plus R13.1.

## Architecture retenue

- `Cabinet.xlsm` reste l'application du poste secretariat : patients, agenda,
  arrivees, courriers a traiter, facturation, feuille de soins et journal.
- un seul `CabinetUnifie_TEST.dotm` est charge par Word sur le poste medecin ;
  il rassemble `Cabinet(1).dotm` et le moteur courrier de `PROD(6)` ;
- les bases maitresses restent sous `\\DS224\CabinetCardio` ;
- `%LOCALAPPDATA%\CabinetCardio\attente_medecin.sqlite` est uniquement un cache
  des patients arrives du jour. Il est efface/reconstruit depuis
  `Echange\Arrives` et ne constitue jamais une base patient de reference ;
- `C:\Mandagout\IMPORT.GDT` est produit sur le poste medecin apres selection du
  patient. Le GDT contient nom, prenom, DDN (`3103`) et sexe (`3110`, 1/2).

## Construction sous Windows / Word 2016

Fermer Word puis activer temporairement « Acces approuve au modele d'objet du
projet VBA ». Depuis PowerShell :

```powershell
.\Build\construire_modele_unifie.ps1 `
  -Prod6 'C:\Mandagout\ModeleCourrierChatGPT_PROD(6).dotm' `
  -Cabinet1 'C:\Mandagout\Cabinet(1).dotm' `
  -Sortie 'C:\Mandagout\CabinetUnifie_TEST.dotm'
```

Installer ensuite l'executable officiel `sqlite3.exe` sur le poste medecin :

```powershell
.\Build\installer_sqlite_medecin.ps1 -SqliteExe 'C:\chemin\sqlite3.exe'
```

Ouvrir `CabinetUnifie_TEST.dotm`, lancer **Debogage > Compiler TemplateProject**,
enregistrer, fermer Word, puis tester exclusivement sur des patients fictifs.

Construire aussi la copie de test du poste secretariat, qui publie les arrivees :

```powershell
.\Build\construire_cabinet_secretariat.ps1 `
  -CabinetXlsm '.\Donnees\Modeles\Deploy\Cabinet.xlsm' `
  -Sortie 'C:\Mandagout\CabinetSecretariat_TEST.xlsm'
```

## Affectation PowerMic

| Touche | Macro |
|---|---|
| A | `Unifie_A_NouvelleLettre` |
| B | `Unifie_B_FormuleAppel` |
| C | `Unifie_C_InsererPatient` |
| D | `Unifie_D_Finaliser` |

Le bouton D utilise le moteur API et les lettres complementaires de `PROD(6)`.
Le constructeur remplace son moteur de gras par celui de `Cabinet(1)` et ajoute
la transmission du document final a `Echange\AEnvoyer` avant sa fermeture.

## Validation obligatoire avant production

Verifier successivement : arrivee secretaria -> cache SQLite -> fenetre des
patients arrives -> GDT Resting12Lead -> A/B/C -> D -> document principal et
lettres complementaires -> file des courriers du secretariat -> impression,
feuille de soins, journal et nouveau rendez-vous. Conserver les trois modeles
actuels en sauvegarde et ne remplacer la production qu'apres ces essais.
