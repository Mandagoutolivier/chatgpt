# Protection de la consultation des archives — recette du 28 septembre 2026

## Correctif client

- Le bouton « Ouvrir le courrier » consulte une copie locale, jamais l'archive NAS.
- Le chemin attendu est exactement Documents/<SHA-256>.docx ou .pdf dans la racine configurée.
- L'empreinte de la source puis de la copie est contrôlée. Une copie préexistante altérée est refusée.
- Les copies portent l'attribut lecture seule ; Word les ouvre avec ReadOnly=True, sans ajout aux fichiers récents, macros désactivées et actualisation des liens désactivée pendant l'ouverture.
- Les options de Word sont rétablies. Le mode Lecture de Word est fermé avant cette restauration, car Word 2016 y refuse UpdateLinksAtOpen. Cette correction s'applique aussi aux instantanés de révision ouverts en lecture seule.
- Les erreurs d'ouverture sont affichées par le formulaire, sans lancer une ouverture de secours non protégée.
- Les corrections d'annexes passent par Patients/_EditionsSecretariat et publication.revise.

Référence de l'API Word : https://learn.microsoft.com/en-us/office/vba/api/word.view.readinglayout

## Vérifications réellement exécutées

Sur le compte Windows standard de recette RDC, Word/Excel 2016 :

- Compilation VBA réelle, puis nouvelle compilation après retrait du code de test.
- Fermeture/réouverture du classeur et comparaison des trois modules au texte source ; références présentes.
- 12/12 assertions : copie locale, empreinte, attribut lecture seule, ouverture Word locale ReadOnly, options Word restaurées, absence de doublon, empreinte invalide/absente refusée, traversée de chemin refusée, copie altérée refusée, source altérée refusée, repli PDF copié et contrôlé. Le lancement d'un lecteur PDF n'est pas couvert.
- Bouton « Ouvrir le courrier » : copie locale des deux pages en lecture seule.
- Bouton « Modifier les annexes dans Word » : copie éditable distincte ; modification fictive limitée à l'annexe ; publication d'une nouvelle version ; archive initiale inchangée.
- Réouverture de la nouvelle version via le bouton de consultation : deux pages en lecture seule.
- Aucune facturation ni impression. Office fermé ; Normal.dotm restauré ; AccessVBOM retiré après fermeture complète d'Office, conformément à l'état initial.

Le fragment Tests/RecetteArchives_lecture.vba doit être ajouté temporairement au module modEchange d'une COPIE de classeur, puis retiré avant recompilation et livraison. U2TesterArchives reçoit une racine UNC isolée contenant Documents/<shaDocx>.docx et Documents/<shaPdf>.pdf. Utiliser des fichiers fictifs aux empreintes distinctes des archives : le test altère volontairement sa source DOCX de test et sa copie locale. Créer une instance Word dédiée sans document avant l'appel. Ne jamais passer la racine réelle des archives à ce test. Les essais ne remplacent pas les autres lignes de RECETTE_WINDOWS.md.

## Constat NAS initial — qualification clinique NON acquise

Les sondes CreateFile réalisées depuis RDC et AX8_MAX demandent un droit puis ferment immédiatement le handle, sans écrire ni supprimer. Résultat sur les archives DOCX/PDF de recette :

| Droit demandé | Résultat réel |
|---|---|
| Lecture | Accordée |
| Écriture | Accordée — échec de protection |
| Suppression des fichiers | Accordée — échec de protection |
| Création dans Documents | Accordée — échec de protection |
| Suppression de Documents ou de ses enfants | Accordée — échec de protection |
| Suppression via le dossier parent | Accordée — échec de protection |

La copie locale protège le parcours applicatif contre un enregistrement sur l'original. Elle ne corrige pas les permissions du partage et ne constitue pas une immutabilité des archives.

Avant toute activation clinique :
1. Relever dans DSM l'identité SMB réellement utilisée par chaque compte de recette, les ACL du partage, de sa racine, de Documents et des fichiers, ainsi que les permissions héritées.
2. Sauvegarder ces droits. Réserver l'écriture des archives au compte du service API. Pour les postes : lecture des archives, aucune écriture/suppression/remplacement, y compris via les droits du parent. Conserver l'écriture nécessaire dans Patients pour les brouillons et copies d'annexes. Ne pas appliquer un refus global au service API.
3. Refaire les sondes depuis les deux comptes de recette : lecture=0 ; écriture/suppression/création=5 (accès refusé). Tout autre code doit être examiné.
4. Publier une nouvelle révision fictive pour vérifier que le service conserve son accès en écriture.
5. Qualifier séparément le compte applicatif strictement secrétariat. La recette présente utilisait le compte historique de test à deux rôles.
6. Poursuivre GDT, calage papier, restauration isolée et les autres contrôles du guide avant toute activation.

Lors de cette première intervention, aucune ACL NAS n'avait été modifiée. Le complément ci-dessous décrit la reprise via DSM. Aucun déploiement clinique ni aucune validation globale de préparation n'a été effectué.

## Complément DSM et SMB — 28 septembre 2026, vers 23 h 10

### Cause et corrections effectuées

Le moniteur DSM a identifié les connexions SMB de RDC et AX8_MAX sous le compte administrateur Mandagout. Documents était déjà réservé en écriture au propriétaire cabinet-api-u2 ; cabinet-u2-archives n'y avait que la lecture. La racine autorisait la lecture à ce groupe, mais Mandagout et administrators avaient lecture/écriture. Le compte effectif expliquait l'échec des sondes.

- RDC : compte NAS existant cabinet-rdc-u2, non administrateur et membre de cabinet-u2-archives, réutilisé. Son mot de passe de recette a été renouvelé et conservé sous forme chiffrée dans le seul profil Windows de recette ; l'identifiant DS224 a été remplacé dans ce profil. Les métadonnées de l'ancien identifiant ont été conservées ; Windows n'en restituait pas le secret. Après renouvellement de la seule connexion SMB RDC desservant U2, DSM confirme cabinet-rdc-u2.
- AX8_MAX : compte NAS cabinet-ax8-u2 créé, membre de cabinet-u2-archives et users, sans appartenance administrateur. Accès aux autres partages explicitement refusé. Identifiant conservé chiffré et enregistré dans le seul profil CabinetU2Medecin.
- Une autorisation de lecture/écriture héritée a été ajoutée pour cabinet-ax8-u2 dans Patients, sans changement des permissions de Documents ni de la racine. Les propriétaires sont conservés. Les ACL de 34 objets ont été relevées avant et après ; DSM a normalisé huit entrées héritées redondantes de cabinet-rdc-u2, dont les droits restent couverts par l'autorisation plus large préexistante.
- Les comptes des services API et les modèles cliniques n'ont pas été modifiés.

### Résultats des sondes

| Session réellement testée | Lecture DOCX/PDF | Écriture, suppression, création d'archives et suppression via le parent |
|---|---|---|
| RDC, session Windows courante après reconnexion | Accordée (0) | Refusées (5) |
| AX8_MAX, nouvelle session réseau isolée sous cabinet-ax8-u2 | Accordée (0) | Refusées (5) |
| AX8_MAX, session Windows courante conservant Mandagout | Accordée (0) | Encore accordées (0) — blocage restant |

Les mêmes contrôles ont été refaits sur les archives d'une nouvelle publication. La création dans Patients et l'écriture du brouillon fictif sont accordées aux comptes limités. Les sondes n'écrivent et ne suppriment aucun contenu.

### Publication réelle avec les droits limités RDC

Depuis l'interface du classeur corrigé : copie d'édition distincte, changement d'une phrase de l'annexe fictive, courrier principal inchangé, sauvegarde et publication.revise réussies. Le service a créé un nouveau DOCX et un nouveau PDF dans Documents. Les empreintes des deux nouvelles archives et des quatre archives précédentes sont conformes à leurs noms. La nouvelle publication a été rouverte par le bouton de consultation : copie locale, ReadOnly=True, deux pages et phrase corrigée présentes. Aucune facturation ni impression.

Publication active : 19a96f877b1d413fb0fa620fff54caee, révision du 28/09/2026 à 23:06:48. La date de validation médicale source reste 21:45:31 ; l'API renvoie Relu=false pour cette correction du secrétariat. Cela ne vaut pas nouvelle validation médicale.

### Point de reprise

La session habituelle AX8_MAX doit encore être reconnectée avec le compte limité puis retestée. Sa connexion DSM observée sous Mandagout dessert simultanément CabinetCardio et CabinetCardioTestU2 ; elle n'a pas été interrompue afin de préserver le travail clinique en cours. Enregistrer et fermer les fichiers NAS concernés avant de renouveler cette connexion, puis confirmer l'identité cabinet-ax8-u2 dans DSM et refaire les sondes dans la session habituelle, sans impersonation.

La protection est validée pour le compte limité RDC et pour une connexion isolée du compte limité AX8 ; elle n'est pas encore effective dans la session habituelle AX8. Il ne s'agit pas d'une immutabilité face à un administrateur NAS. La qualification clinique globale, notamment les rôles applicatifs séparés et les autres contrôles du guide, reste à terminer.
