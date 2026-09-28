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

## Blocage NAS observé — qualification clinique NON acquise

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

Aucune modification des ACL NAS, aucun déploiement clinique et aucune validation globale de préparation n'ont été effectués dans cette intervention.
