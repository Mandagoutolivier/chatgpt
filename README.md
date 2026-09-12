## Installation en un fichier

Télécharger **[Demarrer_Installation_Cabinet.cmd](Installateur/Demarrer_Installation_Cabinet.cmd)**, puis double-cliquer : téléchargement du dépôt privé, vérification, choix domicile/secrétariat/cabinet, préparation et activation guidée. Le serveur Synology doit être disponible. Le lanceur ouvre Office pour la compilation et attend la confirmation des essais avant activation. [Mode d'emploi](Installateur/README.md).

## Nouvelle version dédiée

La [version 2026.09.12 avec service NAS](versions/cabinet-2026.09.12/README.md) reprend les suites de l’audit et fournit un [installateur à trois choix](versions/cabinet-2026.09.12/Installer.ps1) : domicile, secrétariat et cabinet médecin. Voir son guide et sa recette Windows avant activation. Le dossier ci-dessous conserve la livraison précédente.

# Cabinet Cardio — Word, Excel, Synology

Version de travail **2026.09-audit1**, issue de la revue du dépôt du 11 septembre 2026.

Le cabinet conserve ses données maîtres sur le Synology. Le poste secrétaire utilise `Cabinet.xlsm` ; le poste médecin utilise le complément Word `CabinetUnifie.dotm`. Le profil domicile installe les deux parties. La petite base SQLite du médecin sert uniquement de cache de la file d'attente et ne remplace pas les bases du NAS.

**Les corrections sont livrées sous forme de sources et de scripts de reconstruction. Les nouveaux modèles Office n'ont pas été compilés ni testés dans Windows pendant cet audit. Une recette sur données fictives est nécessaire avant leur utilisation avec les patients.**

- [Rapport d'audit et recommandations](cabinet-unifie/AUDIT.md)
- [Installation domicile et cabinet](cabinet-unifie/INSTALLATION_MULTI_POSTES.md)
- [Architecture et parcours utilisateur](cabinet-unifie/INTEGRATION_UNIFIEE.md)
- [Recette Windows à exécuter](cabinet-unifie/RECETTE_WINDOWS.md)
- [Manifeste des composants](cabinet-unifie/Build/manifest.json)
- [Inventaire des sources et empreintes](cabinet-unifie/Tests/inventaire_sources.json)

## Préparer une installation

Télécharger ce dépôt privé depuis GitHub, **Code → Download ZIP**, puis extraire l'archive. Sur le PC Windows, fermer Word et Excel et ouvrir PowerShell dans le dossier `cabinet-unifie` :

```powershell
.\Build\installer_multi_postes.ps1 -Profil Domicile
```

Le mode par défaut est `Preparation` : il construit les fichiers dans un dossier de version, sans remplacer les compléments actifs. Il faut Windows, Word et Excel de bureau, l'accès autorisé au projet VBA pour la construction, et l'accès au NAS. À domicile, connecter d'abord le VPN du cabinet. Voir la procédure détaillée avant le mode `Installation`.

## Vérifications reproductibles sans Office

Depuis la racine du dépôt, avec Python 3 et PowerShell 7 :

```powershell
python cabinet-unifie/Tests/audit_statique.py
python cabinet-unifie/Tests/test_sqlite.py
pwsh -NoProfile -File cabinet-unifie/Tests/test_construction.ps1
```

Ces contrôles ne compilent pas le VBA et ne simulent pas Word, Dragon, l'ECG ou l'imprimante.

Ne placer dans GitHub ni base patients, ni courriers, ni clé API. `DonneesInitiales` contient les ressources de démarrage et l'annuaire de correspondants du projet ; l'initialisation conserve les fichiers déjà présents sur le NAS.
