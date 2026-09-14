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

`generer_lanceur.py COMMIT` assemble les outils, le modèle et les empreintes **depuis les objets Git de ce commit**, sans utiliser les fichiers de travail. Publier d'abord la version, puis générer et publier le lanceur : cette séparation évite une référence circulaire. La révision U0 est destinée à la [recette isolée](../versions/cabinet-2026.09.12/U0_RECETTE.md). Les contrôles GitHub Actions ne remplacent pas une exécution sous Word/Excel.

## Word ou Excel encore actif

Le lanceur laisse cinq secondes aux processus Office pour terminer leur fermeture. Si Word ou Excel reste actif, il affiche son nom, son PID et sa session Windows, puis attend **Entrée** après fermeture ou **Q** pour reprendre plus tard. Enregistrer les documents avant de fermer les applications. Le lanceur ne termine jamais un processus de force.

Si aucune fenêtre Office n'est visible, consulter le Gestionnaire des tâches. Un redémarrage du PC, après enregistrement et fermeture des applications, peut libérer un processus resté actif ; relancer alors le même profil avant d'ouvrir Word ou Excel. La ligne de version affichée au démarrage permet de vérifier quelle copie du lanceur est utilisée. En cas d'échec persistant, transmettre le message et la ligne « Étape » affichée.

## Correction « Module document inattendu : Feuil2 »

Le classeur source contient les feuilles **Accueil** et **Agenda**. Le constructeur associe désormais chacune à son module déclaré, à partir du nom de la feuille, y compris si Excel lui attribue automatiquement `Feuil2` ou `Sheet2`. La feuille Agenda est conservée. Les autres modules document non déclarés restent refusés.

Après cet échec, télécharger le nouveau lanceur et relancer le même profil. Le modèle Word déjà préparé reste dans son ancien dossier ; la nouvelle version construit une nouvelle préparation complète. Ne pas supprimer une feuille ni enlever le contrôle du constructeur pour poursuivre. La correction inclut la validation des modules VBA vides et l'attente guidée de fermeture d'Office.

## Correction « ActiveVBProject » après la compilation Word

L'affectation à `ActiveVBProject` a été supprimée : cette propriété est documentée en lecture seule dans le [modèle VBA Microsoft](https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model#activevbproject). L'assistant affiche désormais le volet de code de `ThisDocument` dans le modèle Word préparé, puis celui de `ThisWorkbook` dans le classeur Excel préparé. Il indique le fichier complet et le projet à compiler. Une erreur mentionne désormais aussi Word ou Excel et le fichier concerné.

Si l'affichage automatique échoue, utiliser **Alt+F11**, puis **Ctrl+R**, sélectionner le module indiqué dans le fichier préparé et choisir **Débogage > Compiler**. Répondre **OUI** uniquement après compilation sans erreur, sinon **NON**. Un refus d'accès VBA, un module absent ou un échec d'enregistrement reste bloquant. Le repli manuel n'active pas les macros et ne valide aucun essai à votre place.

Télécharger le lanceur nouvellement publié, fermer Word et Excel, puis reprendre le même profil et la même racine NAS. Un nouveau paquet crée une nouvelle préparation et redemande les validations. Les contrôles automatiques simulent les objets Office ; les compilations et essais réels doivent toujours être effectués sur le poste.

## Word fermé avant la réponse OUI : erreur RPC 0x800706BA

Fermer Word et Excel **avant de lancer l'installation**. Pendant la compilation guidée, au contraire, **laisser l'application et le fichier ouverts** : choisir Débogage > Compiler, revenir à la console avec **Alt+Tab**, puis répondre **OUI** si la compilation est sans erreur. L'assistant doit encore enregistrer le fichier ; il ferme ensuite lui-même l'application. La même consigne s'applique à Excel.

Le code `0x800706BA` pendant cet enregistrement peut correspondre à une application fermée ou arrêtée. Le journal seul ne permet pas de distinguer une fermeture volontaire d'un plantage. Le nouveau lanceur propose **R** pour rouvrir le même fichier préparé et refaire la compilation, ou **Q** pour mettre en pause. Il ne réutilise jamais le OUI de la tentative interrompue. Il exige que l'enregistrement réussisse et qu'Office confirme l'état enregistré. Les erreurs autres que les déconnexions reconnues restent bloquantes.

Avec le lanceur `17c84b9`, une fermeture volontaire se résout déjà en relançant le même fichier, en conservant le profil, puis en laissant Word et Excel ouverts pendant leurs confirmations respectives. Le correctif de reprise ne répare pas un éventuel plantage d'Office ; si l'erreur revient avec l'application laissée ouverte, relever si celle-ci a disparu ou affiche une boîte de dialogue.
