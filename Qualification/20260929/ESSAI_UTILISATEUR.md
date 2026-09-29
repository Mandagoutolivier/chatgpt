# Essai utilisateur du 29 septembre 2026

## État observé

Essai uniquement sur les profils Windows dédiés et le partage de recette.
Les binaires qualifiés de la livraison `c06d360` restent installés.

- Patient fictif `PESSAI2909A`, consultation `consult-RESSAI2909A`.
- Destinataire principal fictif `CESSAI2909A`.
- Dictée réelle au PowerMic effectuée. Identité ensuite réinsérée depuis la fiche
  avec la commande U2 C ; l'identité avait été dictée dans l'ordre prénom/nom,
  incompatible avec le remplacement exact attendu par l'anonymisation.
- Courrier corrigé et annexe de demande de test d'effort présents sur deux pages.
- Relecture des deux pages confirmée par l'utilisateur, puis transmission
  réussie à 18 h 31 min 57 s (heure locale) via la commande U2 D.
- Publication `5a672c59-b75d-4883-8b31-bdabecd8c7ce` reçue une seule fois dans
  la file du secrétariat. Les empreintes du DOCX et du PDF archivés ont été
  vérifiées depuis RDC ; `Relu=true`, `RelectureEnAttente=0` côté médecin.
- La liste « Courriers validés par le médecin » est ouverte sur RDC. Le
  traitement par la secrétaire, la facturation et l'impression ne sont pas
  encore validés par cet essai.
- Annexe adressée à `CESSAI2909E`, « TEST NORD PARISIEN », fiche explicitement
  fictive sans adresse électronique et sans statut de destinataire par défaut.

## Difficultés relevées

- Curseur observé hors du signet DESTINATAIRE à l'ouverture ; repositionnement
  ponctuel effectué, cause du déplacement encore à établir.
- Commande Dragon « PM Formule Appel » : appel de l'ancienne macro
  `PowerMic_B_FormuleAppel`, inexécutable dans la session U2. La commande
  `Cabinet_B_FormuleAppel` fonctionne. Le profil Dragon actif est historique,
  hors du profil Windows isolé ; il n'a pas été modifié.
- Majuscule de la formule d'appel corrigée dans le document d'essai.
- Les trois fiches importées correspondant à Nord Parisien ont `AValider=1` :
  elles sont exclues de la sélection. Elles n'ont pas été validées implicitement.
- La comparaison des négations signale ici des changements de contexte liés
  aux virgules : les formulations « sans signe d'insuffisance cardiaque » et
  « sans valvulopathie significative » restent présentes dans le résultat.
  Ce constat sur cet exemple ne supprime pas le besoin de relecture.

## Changement demandé pour la prochaine version

Supprimer la fenêtre systématique « Relecture du courrier » pour éviter un clic
de fermeture. Les sources préparent désormais le courrier sans boîte modale,
affichent son état dans la barre d'état et rendent les différences consultables
par « Cabinet > Points de relecture ». Les différences restent enregistrées
dans le brouillon. La commande D conserve la séparation entre préparation et
transmission après relecture ; la préparation ne transmet rien.

Validation de cette modification : 11 contrôles de source de la recette du
20 septembre réussis, audit statique de 85 composants sans erreur, XML du ruban
valide. La compilation Office, le chargement du nouveau ruban et l'affichage
réel de la barre d'état restent à qualifier avant remplacement des binaires.
La modification n'est pas installée dans la session d'essai en cours.

## Enregistrement et fermeture après transfert

Demande complémentaire : supprimer également la fenêtre annonçant le transfert,
enregistrer puis fermer le courrier et appliquer `Nom Prénom aaaammjjhhmm`.

Application ponctuelle sur le courrier déjà transmis :

- Enregistrement vérifié sous `TESTCOURRIER Camille 202609291831.docx` dans le
  dossier du patient fictif ; texte conservé, document enregistré et fermé.
- Copie du PDF archivé sous le même nom, extension `.pdf`, empreinte vérifiée.
- Aucun second envoi, aucune modification des archives internes du service.

Sources de la prochaine version :

- Fermeture du seul courrier après réponse positive de `publish` et sauvegarde
  finale confirmée ; aucune boîte de succès. En cas d'échec, pas de fermeture
  automatique ; un échec après réception est distingué d'un échec de transfert.
- DOCX et PDF nommés à partir de l'identité vérifiée et de l'heure de validation
  mémorisée, avec année sur quatre chiffres et minutes (`yyyymmddhhnn` en VBA).
  Un sous-dossier par publication distingue deux courriers de la même minute
  sans ajouter de suffixe au nom et sans écraser la publication précédente.
- Une publication préparée par l'ancienne version garde ses anciens chemins
  lors de la reprise, pour préserver la commande idempotente.
- La copie de lecture sur RDC porte également le nom habituel. Les archives
  internes restent nommées par empreinte et vérifiées avant ouverture.
- Mode d'export facultatif `IdentiteHorodatage` ajouté et proposé par défaut,
  sans activer l'export. Les configurations historiques restent acceptées.

Les changements automatiques ne sont pas encore installés : la compilation
Office et l'essai complet avec les nouveaux binaires restent à effectuer une
fois la session Office d'essai terminée. Ne pas arrêter les applications actives
pour contourner les garde-fous du constructeur isolé.

Recette à effectuer avec ces binaires : transfert réussi sans dialogue puis
fermeture ; panne de transfert laissant le courrier ouvert ; reprise conservant
le nom et une seule publication ; deux publications dans la même minute sans
écrasement ; lecture sur RDC et concordance du DOCX/PDF. Neuf contrôles VBA du
nommage ont été ajoutés à la recette Office (non exécutés à ce stade).

Contrôles exécutés sur les sources : 11 tests de la recette du 20 septembre,
22 tests d'architecture et d'annexes, audit statique de 85 composants sans
erreur. Inventaires des sources actualisés. Sur AX8MAX, le validateur PowerShell
accepte les trois modes de nommage et refuse un mode inconnu (4 contrôles),
sans modifier la configuration installée.

## Sélection du patient et fiche médecin

Demandes supplémentaires du 29 septembre : fermer la sélection après ouverture
du patient et ajouter « Fiche patient » dans le ruban Cabinet.

- L'ouverture renvoie maintenant le document effectivement créé. La liste des
  arrivées est déchargée seulement en cas de réussite ; Word reprend le premier
  plan et le curseur revient au destinataire. En cas d'échec, la liste reste
  disponible. La reprise réussie d'un brouillon ferme également la liste.
- Le bouton ouvre la fiche du patient lié au courrier actif, avec identité,
  coordonnées, NIR, mutuelle, ALD, notes et choix du médecin traitant. L'identifiant
  du dossier est conservé. Fermer sans enregistrer n'écrit rien sur le NAS.
- Nouvelle opération médecin `patient.update`, limitée au patient d'une
  consultation appartenant à ce médecin ; pas de droit générique `table.update`
  ni de création. La révision de la fiche empêche d'écraser une modification
  concurrente. Commande durable, rejouable et journalisée par le service.
- Une modification d'identité ne remplace pas aveuglément le texte médical :
  la transmission reste bloquée jusqu'à correction/vérification du courrier
  et des annexes. La fiche comporte alors « Identité du courrier vérifiée ».
  La confirmation relit la révision serveur, invalide la relecture et démarre
  un nouveau cycle de correction ; les archives publiées restent inchangées.

Contrôles locaux : 3 tests de contrat/formulaire, 11 tests de recette, 22 tests
d'architecture ; audit de 87 composants de production et 94 avec recette.
Ajout de tests PostgreSQL (rôles, propriété de consultation, révisions,
idempotence, données invalides et archives) et de 7 contrôles VBA d'identité.
Compilation Office, contrôle visuel du formulaire et déploiement du service
`patient.update` nécessaires avant activation du bouton sur les postes.

### Qualification exécutée à 18 h 57–58 (heure de Paris)

Source produit : `06506de063c76cc6811da46b5c503e5b549fc571` (GitHub).

- Sur AX8MAX, fermeture de l'instance Word vide uniquement, après vérification
  qu'aucun document ni modification de Normal n'était en attente.
- Compilation VBA réelle et recette isolée réussies : 415 contrôles Office
  (50 + 52 + 88 + 56 + 67 + 74 + 17 + 11). Aucun accès au NAS ni impression
  pendant cette recette ; paramètres Office et Normal restaurés.
- Construction séparée des binaires sans modules de recette, compilation,
  réouverture et concordance avec les sources réussies : Word 60 composants,
  Excel 27 composants. Constructeur reproductible : `Construire-Correctifs.ps1`.
- GitHub Actions [36601330164](https://github.com/Mandagoutolivier/chatgpt/actions/runs/36601330164)
  réussi : 238 tests serveur PostgreSQL, 5 tests d'interruption, contrôles
  statiques, construction et restauration. Une attente de compteur VBA restée
  à 36 dans un test serveur a été actualisée à 52 avec les nouveaux contrôles.
- SHA-256 Word : `C8DC1F1EF7040FC6DCBAEC3DE553C5BFC64CDD38153846529C1FDC2BE21A8A67`.
- SHA-256 Excel : `5293E7E3D9AB416BDCCE5680F22B9F044F17540879BF3594993FA3476926DC79`.
- Binaires et rapports copiés avec vérification d'empreinte dans
  `\\DS224\CabinetCardioTestU2\Patients\_Qualification20260929\CorrectifsFichePatient-06506de`.

**Pas encore installés.** L'accès SSH non interactif essayé depuis le profil
RDC de recette vers `Mandagout@DS224` retourne 255 sans diagnostic ; le partage
SMB reste accessible. Le service de recette n'a donc pas reçu `patient.update`.
Ne pas installer une fiche présentée comme modifiable tant que cette opération
n'est pas déployée et essayée avec les rôles réels. L'affichage du formulaire,
la sélection réelle avec fermeture de la liste et un aller-retour de modification
de fiche fictive restent à valider ensemble après cette mise à jour.

### Accès SSH confirmé et mise à jour préparée

- Capture utilisateur du 29 septembre à 19 h 11 : connexion SSH interactive
  `Mandagout@DS224`, `hostname`, `whoami`, `sudo -v` et `sudo docker ps`
  réussis. L'API U2 utilise bien `127.0.0.1:8766`, l'autre projet `8765`.
  L'échec du client SSH lancé par l'agent n'est donc pas un défaut d'accès NAS.
- `MettreAJour-ServiceEssai.py` prépare uniquement `patient.update` sur U2.
  Le dossier actif est déduit des labels Docker, pas du chemin générique du
  guide. Contrôles des ports, volumes, rôle du conteneur, configuration et
  concordance des sources sur disque et dans l'image avant modification.
- Archive GitHub 06506de : SHA-256
  `2a3c05875a0cbc6dee71e68019fed2561366e0defe4d41b9f8356834aff17707`.
  Seuls `contract.py` et `service.py` changent côté serveur. L'image est
  reconstruite depuis l'image locale vérifiée, sans téléchargement, puis
  comparée aux sources qualifiées. Les sources sur disque sont aussi mises
  à jour pour conserver la reproductibilité des constructions ultérieures.
- Conservation privée de l'image précédente, des deux sources remplacées
  et d'un dump PostgreSQL ; ce dump seul n'est pas une sauvegarde complète
  base et documents. Redémarrage de la seule API avec `--no-deps`, contrôle
  de santé et du code actif, puis vérification des autres conteneurs.
  Retour automatique de code et d'image en cas d'échec après modification.
- Syntaxe Python contrôlée, refus hors DS224 vérifié. Quatre simulations
  locales avec commandes Docker substituées : contrôle seul, succès,
  image candidate non conforme (arrêt avant modification), échec du
  redémarrage (restauration des sources et de l'image) : réussies.
  Ces simulations ne remplacent pas l'exécution réelle sur le NAS.
- Sans argument, le script ne fait que les contrôles. `--appliquer` exécute
  la mise à jour. Aucune lecture de jeton, migration, installation Office ni
  validation clinique n'est effectuée par ce script. Le résultat NAS et
  les essais authentifiés des deux postes restent à obtenir.
