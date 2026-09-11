# Audit du dépôt Cabinet Cardio — 11 septembre 2026

## Conclusion

Le dépôt initial comportait plusieurs défauts susceptibles de bloquer la construction, de perdre une arrivée, de publier un courrier incomplet ou de doubler un enregistrement comptable. Les corrections de cet audit sont intégrées aux sources et aux scripts. **La version obtenue doit encore être compilée et essayée dans Word/Excel sur Windows avant utilisation avec les patients.**

L'architecture Word + Excel + classeurs NAS peut servir de transition. Elle conserve une fragilité structurelle : plusieurs fichiers et applications participent à une consultation sans transaction commune. Le chantier suivant recommandé est un service métier sur le Synology avec une base serveur, en conservant Word pour la dictée et SQLite comme simple cache local.

## Périmètre et méthode

Point de départ : dépôt privé `Mandagoutolivier/chatgpt`, commit `c61e1ea4086a78ffb233ff2447de96d1feb54955`, 18 fichiers suivis. Les scripts, la configuration et les trois conteneurs Office ont été examinés. Le VBA embarqué a été extrait afin de relire les procédures et leurs appels, puis de rendre ces sources visibles dans Git. Les formulaires ont été examinés côté code ; leur rendu et leurs événements restent à valider dans Office.

La livraison contient **73 fichiers de source VBA distincts, 24 735 lignes**, déclarés dans **79 composants** répartis entre Word et Excel. Six modules communs sont utilisés par les deux hôtes ; le total par projet est donc de 25 803 lignes. L'[inventaire](Tests/inventaire_sources.json) donne chaque fichier, son nombre de lignes et son SHA-256. Ce périmètre inclut les branches anciennes encore présentes, dont plusieurs entrées sont désormais désactivées. L'inventaire atteste le contenu, pas l'absence de tous les bugs possibles.

Ont également été examinés : structure ZIP/OOXML et rubans, schémas des classeurs, annuaires et dictionnaires de démarrage, 23 formules de l'annuaire, signets et contenu du modèle `LETTRE TYPE.dot`, ainsi que les verrous, transitions de file, appels API, exports GDT et scripts d'installation. `Normal.dotm` fourni dans la conversation a servi à rechercher les conflits de macros ; il n'est pas ajouté ni remplacé dans ce dépôt.

Les ressources manquantes viennent du dépôt public `Mandagoutolivier/claud-ai-cabinet`, état `402e6cef` consulté pendant la revue. Les trois dictionnaires DDE proviennent de ses données R12 ; cela ne transforme pas le moteur PROD6 en moteur R12. Aucun patient réel n'a été importé dans les tests. Les originaux Office ne sont pas modifiés ; la nouvelle version est reconstruite à partir du manifeste.

## Quelle version pour PROD(6) ?

**Le nom du fichier ne suffit pas et aucun numéro global R11/R12/R13 n'est établi par les sources inspectées.**

| Fichier | SHA-256 de référence |
|---|---|
| `ModeleCourrierChatGPT_PROD(6).dotm` | `a8ccb702e8defc9fce19b9afb968c4f7cd5b25d59da562dba172bc6b51b85524` |
| `Cabinet(1).dotm` | `493ed179d97ec4f15cafe39e44f9a6d2c0ce56a42927cf0fd4acafd860dabdb5` |
| `Cabinet.xlsm` | `714e9458066bd28b8b6e6f6e9f65b2d673d3285c1915b57216e780d5e75f04ff` |

Les sources de PROD6 portent notamment `DOMICILE-DOUBLE-API-DETAILLE-1`, `DOMICILE-GRAS-1Z`, `v22-P5` et des mentions DDE2-R4/R5. Elles correspondent à une branche de courrier avec double appel API. La documentation du paquet R12 décrit une évolution depuis une base R7 ; le rattachement de ce fichier précis à R7 reste un **indice de filiation**, pas une preuve de numéro de livraison. Il faut un ancien manifeste ou une empreinte de livraison pour trancher définitivement. La nouvelle intégration est nommée `2026.09-audit1`, sans lui attribuer artificiellement R13 ou R14.

Un contrôle heuristique d'oletools signale aussi des identifiants présents dans le désassemblage du code compilé mais absents du texte source. Cabinet contient principalement des noms de contrôles ou suffixes de type ; PROD6 contient davantage d'écarts, dont une mention `v23`. Les noms désassemblés sont parfois incohérents. Ce résultat **ne prouve ni modification malveillante, ni numéro de version, ni divergence sémantique précise** : les limitations du désassembleur et les états de compilation doivent être départagés dans Office. Les constructeurs réécrivent tous les modules déclarés depuis les sources revues ; la compilation puis la réouverture dans Office figurent explicitement dans la recette. Aucun ancien p-code n'a été exécuté pendant l'audit.

## Défauts corrigés

« Corrigé » désigne une modification du code examinée et soumise aux contrôles disponibles. Les comportements dépendant d'Office, du NAS Windows ou d'un périphérique attendent la recette décrite plus bas.

| Priorité | Constat et conséquence | Correction / fichiers concernés |
|---|---|---|
| Bloquant | Fusion de macros publiques homonymes : résolution ambiguë possible | Renommage de `LireCleOpenAI` côté Cabinet et de `FichierExiste` côté PROD ; contrôle des doublons et des appels qualifiés dans `Tests/audit_statique.py` |
| Bloquant | Construction partielle : des corrections de source pouvaient ne jamais atteindre les modèles | `Build/manifest.json` couvre tous les composants ; les deux constructeurs injectent toutes les sources, contrôlent modèles et références, conservent les dessins des formulaires |
| Majeur | Import de fichiers VBA UTF-8 par une voie ANSI et modification fragile du ruban | Injection Unicode par `CodeModule`, fusion des relations et types OOXML, contrôle du SHA-256 du projet VBA après insertion du ruban — `outils_construction.ps1` |
| Majeur | Verrous fondés sur l'âge du fichier, suppression d'un verrou encore actif, copies temporaires partagées | Verrous maintenus ouverts, pas de vol automatique, temporaires GUID et renommages atomiques — `modFichiers`, `modBaseIO` |
| Majeur | Écritures NAS et classeurs en lecture seule pouvant être présentés comme sauvegardés | Vérification `ReadOnly`, propagation des erreurs, sauvegardes préalables et fermeture sous verrou — `modBaseIO`, formulaires de correspondants |
| Majeur | Ouverture automatisée de classeurs/modèles pouvant déclencher des macros ou liens inattendus | `AutomationSecurity=3`, événements désactivés et liens non actualisés dans les ouvertures techniques corrigées |
| Majeur | Cache identifié par patient : un second rendez-vous écrasait le premier | Cache `attentes_v2` identifié par `rdv_id`, reconstruction dans `BEGIN IMMEDIATE`, arrêt SQLite sur erreur — `modAttenteLocale` |
| Majeur | Sélection implicite ou ancienne file locale après défaut réseau | Sélection explicite, relecture NAS obligatoire, réservation exclusive de l'arrivée, entrée de reprise pour réservation interrompue — modules d'intégration |
| Majeur | Identité du document ou export ECG dissociés du patient sélectionné | Métadonnées figées et comparaison avec la fiche NAS ; absence d'inférence du sexe ; dates strictes — `modIntegrationUnifie`, `modCourrier`, `modEcg`, `modGdt` |
| Majeur | Annonces d'arrivée multiples et collisions d'identifiants entre années | Publication déterministe par rendez-vous/date, transitions sous verrou, identifiant de consultation incluant l'année et propagation de `AnneeAgenda` |
| Majeur | Le brouillon de reprise pouvait rester pointé vers un ancien fichier | Actualisation atomique du chemin du brouillon dans l'événement `EnCours` après chaque sauvegarde — `modIntegrationUnifie` |
| Majeur | Réponse API incomplète, refusée ou mal extraite pouvant être utilisée comme résultat | Parseur JSON strict, statut `completed` exigé, traitement explicite des refus/erreurs HTTP, limites de taille/profondeur — `modJson`, `modOpenAI_v22_corrige` |
| Majeur | Échec de la seconde API assimilé à « aucune annexe » ; anciennes annexes retirées trop tôt | Échec propagé, validation des sorties avant remplacement, protection contre réentrée et brouillon préalable — `modProdRapide`, moteur API |
| Majeur | Plusieurs transports et commandes historiques contournaient le flux unifié | Ancien transport Cabinet et anciens générateurs directs bloqués ou redirigés ; un seul parcours A/B/C/D actif |
| Majeur | Publication au secrétariat avant l'obtention de toutes les sorties | DOCX et PDF obligatoires avant événement, révision distincte, consultation stable — `modValidation` |
| Majeur | Double validation, reprise ou révision de courrier pouvant doubler les actes | Ajout de séance idempotent sous verrou ; refus si un identifiant existant appartient à un autre patient — `modBaseIO`, `modActes` |
| Majeur | Réimpression à partir des choix actuels du formulaire, potentiellement différents de la séance facturée | Réimpression des actes/montants relus dans le journal ; marquage après retour de `PrintOut` — `modActes`, `ufChoixActe` |
| Majeur | Lignes de soins supplémentaires ignorées, tiers payant traité comme encaissement | Refus explicite de plus de quatre lignes, distinction du paiement ; date de l'acte conservée — `modCerfaPrint`, `modActes` |
| Majeur | Créneaux concurrents acceptés après une vérification faite avant verrou | Nouvelle vérification des chevauchements sous le verrou d'écriture ; durée bornée — `modBaseIO`, `modAgenda` |
| Majeur | Destinataire annexe choisi malgré une même clé pour des personnes/adresses différentes | Rejet des clés ambiguës et des lignes à valider — `modCabinetTestDestinations` |
| Majeur | Formules d'annuaire limitées à la ligne 499 : nouvelles structures invisibles | Extension des plages `Structures` avant calcul et sauvegarde — `modEnregistrementCorrespondants`, `modSaisieCorrespondants` |
| Modéré | Recherche textuelle d'un corps vide dans le modèle d'annexe | Repérage prioritaire par signet ; le modèle initial a un corps vide valide — `modLettresComplementairesModele` |
| Modéré | Accents dégradés après ajout aux dictionnaires, gras incomplet, texte de patient interprétable comme formule Excel | Réécriture UTF-8 atomique des dictionnaires, moteur Cabinet commun, cellules de saisie explicitement textuelles |
| Modéré | Données GDT invalides : caractères de contrôle, encodage ou longueur non compatibles | Validation CP1252, taille de ligne, identité et remplacement atomique de `IMPORT.GDT` — `modGdt` |
| Majeur | Installation écrasant des fichiers actifs avant construction terminée | Mode préparation par défaut, étapes de construction avant activation, sauvegardes et tentative de restauration locale — `installer_multi_postes.ps1` |
| Majeur | Ressources manquantes et initialisation risquant de remplacer une base personnalisée | Dix ressources initiales avec empreintes, copie seulement si absentes, ajout de colonnes avec sauvegarde, SQLite fixé par empreinte — `initialiser_nas.ps1` |
| Modéré | Anciennes procédures de test susceptibles d'écrire dans les données de travail | Points d'entrée historiques bloqués ; nouvelles fixtures VBA hors réseau et tests statiques/SQLite distincts |

La documentation des [ouvertures Excel](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open), de la [sécurité d'automatisation Excel](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity) et de [Word](https://learn.microsoft.com/en-us/office/vba/api/word.application.automationsecurity) a servi à vérifier les paramètres COM. Le vrai chemin de démarrage Word est obtenu par `Options.DefaultFilePath(8)`, conformément à [WdDefaultFilePath](https://learn.microsoft.com/en-us/office/vba/api/word.wddefaultfilepath).

## Risques restants et conseils de conception

| Priorité | Limite restante | Action conseillée |
|---|---|---|
| Avant mise en service | Compilation VBA et comportements COM non exécutés ici ; formulaires hérités des binaires | Compiler Word/Excel sur les versions du cabinet, rouvrir les fichiers et dérouler la recette ; conserver les anciens modèles |
| Avant mise en service | Pas de transaction unique entre brouillon, DOCX/PDF, arrivée, file et agenda | Tester les coupures à chaque étape ; documenter les reprises. À terme, journal transactionnel et service NAS |
| Avant mise en service | Rapprochement du destinataire principal encore partiellement fondé sur le nom ; annuaires Cabinet/PROD distincts | Vérification des homonymes ; unifier les correspondants par identifiant stable, puis utiliser cet identifiant dans les lettres |
| Avant mise en service | Masquage de texte libre incomplet et génération médicale non déterministe | Vérifier le texte transmis et les résultats avec données fictives ; relecture du médecin avant impression/envoi |
| Avant mise en service | CERFA papier : patient/assuré, ayants droit, NIR et identifiants du praticien pas entièrement couverts | Valider les cas réellement utilisés ; compléter configuration. Aucun connecteur FSE/CPS/SESAM-Vitale n'est livré |
| Court terme | Lecture complète de classeurs, copies temporaires locales, latence VPN et recherches répétées | Mesurer avec le volume réel ; centraliser les lectures dans un service, limiter les champs transmis, gérer explicitement la durée de cache |
| Court terme | États multi-fichiers et protocoles de verrou incompatibles avec les anciennes versions | Mettre à jour tous les clients hors consultation, tester restauration NAS, éviter le fonctionnement simultané ancien/nouveau |
| Court terme | Parseur JSON VBA et détection DDE restent du code spécifique ; segmentation clinique heuristique | Introduire un schéma de sortie validé et un corpus de réponses/lettres de régression ; ne pas confondre détection de mot et décision médicale |
| Court terme | Beaucoup de code mort et noms trompeurs (`modClaude`, modules `Test` utilisés en production) | Retirer progressivement les branches inaccessibles après recette, nommer les modules par leur responsabilité, documenter les interfaces |
| Court terme | Édition directe des classeurs/dictionnaires en dehors de l'application | Encadrer l'édition pendant l'arrêt des clients ; remplacer ensuite les écritures directes par le service NAS |
| Moyen terme | Pas de serveur d'authentification, rôles et journal métier centralisés ; sauvegarde périodique NAS hors application | Définir comptes par rôle, droits du partage, stratégie de restauration et journal des transitions sans contenu médical inutile |
| Moyen terme | Tests actuels surtout structurels ; absence de compilation/recette Office automatisée | Ajouter un poste Windows de validation et des jeux fictifs ; séparer métier et COM pour tester les règles sans Office |

Un passage à PostgreSQL/MariaDB sur le NAS est une recommandation d'architecture, pas une migration déjà effectuée. Mettre simplement SQLite sur SMB ne résout pas les écritures concurrentes ; voir [la documentation SQLite/WAL](https://www.sqlite.org/wal.html).

Pour l'écriture du code : remplacer les `On Error Resume Next` de métier par des erreurs typées et propagées ; les limiter aux nettoyages dont l'échec est traité. Réduire les `ActiveDocument` implicites, passer le document et l'identifiant de consultation explicitement. Utiliser une seule responsabilité pour les fichiers, les montants, l'identité et le transport API. Conserver des commentaires décrivant les invariants plutôt que l'historique de chaque tentative. Une partie de ce travail est faite dans les chemins actifs ; le retrait intégral du code ancien attend la recette pour éviter de casser des commandes Dragon non recensées.

## Vérifications réellement exécutées

| Contrôle | Résultat dans cet environnement |
|---|---|
| Inventaire et contrats statiques Python | 79 composants, aucun module déclaré absent, aucun doublon public standard détecté, appels qualifiés contrôlés, procédures équilibrées |
| Empreintes / conteneurs | Trois modèles source et dix ressources vérifiés ; CRC des trois archives Office valides |
| Tests PowerShell 7.6.6 sous Linux | **28 contrôles réussis**, dont parsing des six scripts, conservation des types/relations OOXML, ruban idempotent, SHA-256 VBA inchangé par le seul ajout du ruban et sauvegarde de publication |
| Tests SQLite via Python, SQLite 3.53.1 | **4 tests réussis** : deux RDV pour un patient, annulation d'une suppression sur erreur, isolation lecteur/transaction, exclusion de deux écrivains |
| Classeurs initiaux | 23 formules inspectées dans l'annuaire ; pas de liens externes ni macrosheets trouvés dans les XLSX examinés ; pas d'erreur Excel mise en cache trouvée |
| Recherche ciblée de secrets | 95 fichiers texte/désassemblages examinés ; aucune clé OpenAI/GitHub ni clé privée détectée par les motifs recherchés ; ce contrôle n'est pas une preuve exhaustive |
| Tests VBA `Audit_LancerTests` | Écrits, **non exécutés** : ils nécessitent Word Windows |
| Compilation Office, PowerShell Windows 5.1, Dragon, Resting12Lead, imprimante, NAS réel | **Non exécutés** dans cet environnement |

Les tests SQLite emploient les fragments de schéma/transaction extraits du VBA, mais pas le lancement de `sqlite3.exe` depuis Word ; celui-ci reste à tester sur Windows. Le paquet SQLite Windows téléchargé par l'installateur est fixé séparément dans `Build/sqlite.lock.json`.

Commandes reproductibles depuis la racine :

```powershell
python cabinet-unifie/Tests/audit_statique.py
python cabinet-unifie/Tests/test_sqlite.py
pwsh -NoProfile -File cabinet-unifie/Tests/test_construction.ps1
```

Le [plan de recette Windows](RECETTE_WINDOWS.md) fournit les résultats attendus pour l'installation, les identités, A/B/C/D, les erreurs API, la concurrence NAS, la facturation et les reprises. Une réussite des tests structurels n'autorise pas à annoncer que l'ensemble du cabinet est testé ou prêt pour la production.
