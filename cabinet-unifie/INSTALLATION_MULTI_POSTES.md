# Installation multi-postes

Le meme script installe les trois configurations. Les bases de reference ne
sont jamais placees sur les PC : elles restent sous `\\DS224\CabinetCardio`.

Le dossier `ModelesSource` du depot contient :

- `ModeleCourrierChatGPT_PROD(6).dotm` ;
- `Cabinet(1).dotm` ;
- `Cabinet.xlsm` ;

Il peut aussi contenir `sqlite3.exe`. S'il est absent, le script telecharge
automatiquement le paquet Windows x64 depuis `sqlite.org`.

Fermer completement Word et Excel. L'acces approuve au modele d'objet VBA doit
etre active temporairement dans Word et Excel pour permettre la construction.

## PC domicile — deux parties

Connecter d'abord le VPN `Cabinet Freebox Pro`, puis :

```powershell
PowerShell -ExecutionPolicy Bypass -File .\Build\installer_multi_postes.ps1 -Profil Domicile
```

## PC secretariat du cabinet

```powershell
PowerShell -ExecutionPolicy Bypass -File .\Build\installer_multi_postes.ps1 -Profil CabinetSecretariat
```

## AX8_MAX — poste medecin du cabinet

```powershell
PowerShell -ExecutionPolicy Bypass -File .\Build\installer_multi_postes.ps1 -Profil CabinetMedecin
```

Le script verifie l'acces en lecture/ecriture au NAS, sauvegarde les modeles
locaux existants, sauvegarde `Config\config.ini`, construit les copies Word et
Excel, installe la partie demandee et conserve un journal d'installation sous
`%APPDATA%\CabinetCardio`.

Il ne remplace aucun fichier `Patients.xlsx`, `Agenda_*.xlsx` ou
`Journal_*.xlsx`. La migration eventuelle d'anciennes bases doit rester une
operation distincte, precedee d'une sauvegarde et d'un essai de restauration.
