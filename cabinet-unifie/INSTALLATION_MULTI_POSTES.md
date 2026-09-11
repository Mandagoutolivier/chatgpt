# Installation domicile et cabinet

## Ce que fait le script

`Build/installer_multi_postes.ps1` utilise le même code pour les trois profils :

| Profil | Partie secrétaire | Partie médecin |
|---|---|---|
| `Domicile` | Oui | Oui |
| `CabinetSecretariat` | Oui | Non |
| `CabinetMedecin` | Non | Oui |

Word **et** Excel de bureau sont nécessaires pour les trois profils : Word lit les classeurs par Excel et Excel utilise Word pour la feuille de soins. Dragon et Resting12Lead doivent être installés/configurés séparément sur le poste médecin. Le script ne les installe pas.

Les modèles source originaux sont conservés. Le script contrôle leurs empreintes, reconstruit les composants depuis `Src`, conserve les dessins des formulaires et crée les fichiers locaux. Les bases maîtres restent sur le Synology.

## Préparer le PC et le NAS

1. Sur Windows, installer Word et Excel de bureau et les ouvrir une première fois. La compatibilité cible du code est Office 2016 ou ultérieur ; elle reste à vérifier sur les éditions exactes du cabinet.
2. Depuis le compte Windows qui utilisera le logiciel, vérifier l'accès en lecture/écriture à `\\DS224\CabinetCardio` et à la sortie `\\DS224\home\sortiedragon`. À domicile, activer le VPN Cabinet Freebox Pro. Le script ne crée ni partage Synology, ni compte réseau, ni VPN.
3. Télécharger l'archive du dépôt privé depuis **GitHub → Code → Download ZIP**, puis l'extraire entièrement. Ne pas lancer les scripts directement depuis le ZIP. Vérifier la provenance du téléchargement avant de débloquer localement les fichiers téléchargés si Windows le demande.
4. Fermer Word et Excel sur ce PC. Pour une installation ou migration du NAS, fermer les anciennes applications du cabinet sur **tous les postes** : les anciens verrous ne sont pas compatibles avec le protocole corrigé.
5. Pour la construction, dans le Centre de gestion de la confidentialité d'Office, autoriser l'accès approuvé au modèle d'objet du projet VBA. Les scripts ne modifient pas automatiquement la politique de macros. Désactiver à nouveau cet accès après la construction si son usage permanent n'est pas nécessaire. Une politique d'entreprise bloquante doit être traitée par l'administrateur.

## Préparation sans remplacement des applications actives

Ouvrir PowerShell dans `cabinet-unifie` :

```powershell
.\Build\installer_multi_postes.ps1 -Profil Domicile
```

Ou, au cabinet :

```powershell
.\Build\installer_multi_postes.ps1 -Profil CabinetSecretariat
.\Build\installer_multi_postes.ps1 -Profil CabinetMedecin
```

Exécuter chaque commande sur le PC correspondant. `Preparation` est le mode par défaut. Il vérifie l'accès au NAS mais ne l'initialise pas et ne remplace pas le complément Word actif. Les fichiers construits et les journaux sont conservés sous `%APPDATA%\CabinetCardio\Versions` et `Sauvegardes`.

Le profil médecin télécharge SQLite depuis le site officiel avec une version et une empreinte fixées dans `Build/sqlite.lock.json`. Le paquet par défaut est Windows x64. Pour fournir un exécutable déjà disponible :

```powershell
.\Build\installer_multi_postes.ps1 -Profil Domicile `
  -SqliteExe 'C:\Installation\sqlite3.exe' `
  -SqliteSha256 'REMPLACER_PAR_EMPREINTE_SHA256_VERIFIEE'
```

Le script vérifie l'exécution de SQLite et une requête élémentaire. Le cache patient est créé au premier rafraîchissement de la file. Le modèle Office généré n'est pas automatiquement déclaré valide : effectuer la [recette Windows](RECETTE_WINDOWS.md).

Pour une première recette isolée, utiliser un partage de test du Synology et initialiser ses ressources avec `Build/initialiser_nas.ps1 -RacineNas '\\DS224\CabinetCardioTest'`. Dans **sa** configuration, remplacer aussi la destination `[SORTIE] Dossier` par un dossier de test sur le NAS, afin de ne pas publier dans la file réelle. Un simple changement de `RacineNas` ne modifie pas cette destination indépendante.

## Activer après la recette

Faire d'abord une sauvegarde du partage NAS. Installer le secrétariat avant le médecin, ou utiliser le profil domicile qui contient les deux :

```powershell
.\Build\installer_multi_postes.ps1 -Profil Domicile -Mode Installation
```

Au cabinet :

```powershell
.\Build\installer_multi_postes.ps1 -Profil CabinetSecretariat -Mode Installation
.\Build\installer_multi_postes.ps1 -Profil CabinetMedecin -Mode Installation
```

Pour changer la racine NAS et le dossier GDT local :

```powershell
.\Build\installer_multi_postes.ps1 -Profil CabinetMedecin -Mode Installation `
  -RacineNas '\\DS224\CabinetCardio' -DossierGdt 'C:\Mandagout'
```

Le mode Installation reconstruit les fichiers depuis les sources présentes ; il ne reprend pas automatiquement le binaire d'une préparation précédente. Conserver exactement le même état du dépôt entre recette et installation et recontrôler les binaires activés dans Office.

L'initialisation copie uniquement les ressources NAS absentes et ajoute les colonnes nécessaires à `Patients.xlsx`, avec sauvegarde préalable. Elle ne remplace pas une configuration, un annuaire ou un dictionnaire existant. Contrôler donc leurs chemins et leurs données si le cabinet possède une ancienne installation personnalisée. Les tarifs fournis sont des données de départ à vérifier, pas une mise à jour réglementaire.

## Après installation

- **Médecin** : `CabinetUnifie.dotm` est placé dans le vrai dossier de démarrage configuré dans Word. Les anciens compléments nommés `Cabinet.dotm` ou `ModeleCourrierChatGPT*.dotm` de ce dossier sont sauvegardés puis retirés. `Normal.dotm` est conservé.
- **Secrétariat** : `Cabinet.xlsm` est installé dans `Documents\CabinetCardio`, avec raccourci sur le bureau.
- **Configuration commune** : `Config\config.ini` sur le NAS. Vérifier coordonnées, RPPS/numéro AM, imprimante, sortie et dictionnaires. Les champs RPPS/numéro AM sont vides dans le fichier initial.
- **Configuration locale** : `%APPDATA%\CabinetCardio\chemin.txt` désigne le NAS ; `poste.ini` porte le profil, le dossier ECG et les réglages de poste. Les autres paramètres locaux existants sont conservés.
- **API** : renseigner `OPENAI_API_KEY` pour le compte Windows du médecin, ou créer `%APPDATA%\CabinetCardio\openai.key` contenant uniquement la clé. Relancer Word après modification d'une variable d'environnement. Ne pas envoyer cette clé au secrétariat ni la placer dans GitHub.
- **Dragon** : affecter A/B/C/D aux quatre macros `Unifie_*` de `modPowerMicUnifie`. Vérifier que les anciens raccourcis de `Normal.dotm` n'interceptent pas ces commandes.
- **ECG** : configurer l'import GDT de Resting12Lead vers `C:\Mandagout\IMPORT.GDT`, puis tester avec une identité fictive, notamment les accents, le sexe et la date de naissance.

## En cas d'échec

Une construction échouée conserve la version active. Lors d'un échec pendant l'activation, le script tente de restaurer les fichiers locaux remplacés et signale les restaurations incomplètes. Les ressources NAS déjà créées et les ajouts de colonnes ne sont pas annulés automatiquement.

Chaque installation conserve `installation.log`, les copies précédentes et `restauration.json` dans `%APPDATA%\CabinetCardio\Sauvegardes\<date-identifiant>`. Pour une restauration manuelle, fermer Word et Excel, restaurer chaque `backup` à sa `destination`, en ordre inverse des modifications. Une entrée sans sauvegarde désigne un fichier créé par cette installation ; ne le retirer qu'après avoir vérifié qu'il ne contient pas de modifications ultérieures. Ne pas restaurer aveuglément une ancienne base patients par-dessus les nouvelles saisies.

Aucune migration de données réelles ni installation sur les PC du cabinet n'a été exécutée pendant l'audit.
