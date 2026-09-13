# Installation guidée Cabinet Cardio

Télécharger **Demarrer_Installation_Cabinet.cmd** avec le bouton « Download raw file » de GitHub, puis double-cliquer. Ce fichier autonome contient le PowerShell lisible fourni à côté ; il n'a besoin d'aucun autre fichier local au départ.

1. Le navigateur télécharge la version précise du dépôt privé. Se connecter à GitHub si nécessaire. Le lanceur repère le ZIP, l'extrait et vérifie tous les fichiers de la version.
2. Choisir **1 Domicile**, **2 Secrétariat**, **3 Cabinet médecin**.
3. Suivre les indications pour le NAS, la compilation Office et les essais. Après validation, l'activation se poursuit automatiquement, avec sauvegarde de l'installation locale précédente.

Le lanceur demande maintenant le **dossier réel des données sur le NAS**. Il ne suppose plus l'existence de `\\DS224\CabinetCardio`. Vous pouvez saisir son chemin UNC ou taper **B** pour parcourir le réseau. Un chemin déjà configuré est proposé, puis vérifié. Pour un emplacement neuf, saisir un sous-dossier dans un partage existant et confirmer par **CREER** ; cela ne crée pas de partage DSM et ne migre pas les anciennes bases.

Les dossiers contenant le code GitHub ne doivent pas être choisis comme dossier des données. Avec `home`, vérifier que les comptes NAS des deux postes accèdent réellement au même dossier : ce nom peut désigner un dossier personnel différent selon le compte connecté. Le lanceur demande **COMMUN** pour confirmer cette vérification.

Le même fichier fonctionne sur les trois postes. Utiliser la session Windows habituelle, sans « Exécuter en tant qu'administrateur ». Windows, Word et Excel de bureau sont nécessaires. À domicile, le VPN doit permettre l'accès au NAS. Le service NAS doit déjà être déployé ; son adresse HTTPS et le jeton du compte du poste sont demandés à l'activation.

La compilation dans Word/Excel et les essais réels nécessitent votre intervention. Le script ouvre les projets et le guide, recueille les confirmations, puis poursuit jusqu'à l'activation ; il ne certifie pas un essai non réalisé. **PAUSE** conserve la préparation pour une reprise avec le même profil. Les données partagées restent sur le NAS.

Si Windows affiche une interdiction relevant d'une stratégie administrateur, le lanceur s'arrête : il ne modifie pas cette stratégie. L'autorisation VBA par utilisateur est temporaire et journalisée, sans activation globale des macros. Après une coupure brutale, relancer le même fichier pour permettre sa restauration.

Détails : [guide des trois postes](../versions/cabinet-2026.09.12/INSTALLATION_MULTI_POSTES.md).

## Maintenance du lanceur

`generer_lanceur.py COMMIT` assemble `outils_telechargement.ps1`, le modèle et les empreintes des fichiers locaux de `versions/cabinet-2026.09.12`. Le commit indiqué doit contenir exactement ces fichiers. Publier d'abord la version, puis générer et publier le lanceur : cette séparation évite une référence circulaire. Les contrôles GitHub Actions testent l'extraction, le refus des fichiers altérés, la reprise et la syntaxe sur PowerShell 5.1 et 7. Ils ne remplacent pas une exécution sous Word/Excel.

## Word ou Excel encore actif

Le lanceur laisse cinq secondes aux processus Office pour terminer leur fermeture. Si Word ou Excel reste actif, il affiche son nom, son PID et sa session Windows, puis attend **Entrée** après fermeture ou **Q** pour reprendre plus tard. Enregistrer les documents avant de fermer les applications. Le lanceur ne termine jamais un processus de force.

Si aucune fenêtre Office n'est visible, consulter le Gestionnaire des tâches. Un redémarrage du PC, après enregistrement et fermeture des applications, peut libérer un processus resté actif ; relancer alors le même profil avant d'ouvrir Word ou Excel. La ligne de version affichée au démarrage permet de vérifier quelle copie du lanceur est utilisée. En cas d'échec persistant, transmettre le message et la ligne « Étape » affichée.
