# Recette avant activation

Effectuer cette recette sur un PC secondaire, avec le projet NAS `cabinetcardio-test-u2`, le port local `8766`, le partage d'exemple `\\NAS-RECETTE\CabinetCardioTestU2` et des données fictives. L'installation clinique actuelle doit rester disponible et inchangée pour les consultations du lendemain. Noter date, révision du service, versions Windows/Office/Dragon/Resting12Lead, compte, résultat obtenu et captures éventuelles sans données réelles. Une ligne non exécutée reste non validée.

| Essai | Résultat attendu |
|---|---|
| Projet Compose et volumes | Projet `cabinetcardio-test-u2`, port `8766`, données, PostgreSQL, sauvegardes et secrets distincts |
| Service `/health` | Protocole 2, schéma 2 et révision `2026.09.16-u2b` |
| Service clinique pendant la recette | Toujours accessible séparément ; aucun arrêt, changement de port ou changement de volume |
| Préparation des profils Domicile, Secretariat et Cabinet | Binaires dans un dossier de préparation ; anciens fichiers actifs conservés |
| Word/Excel : Débogage > Compiler, fermer/rouvrir | Aucune erreur, aucune référence manquante, sources conformes au manifeste |
| Recette automatisée Word | Socle Word au moins `34`, puis extension U2 exactement `17/17`, aucun échec |
| Recette automatisée Excel | Exactement `25/25`, aucun échec |
| Total Office annoncé faux, incomplet ou incohérent | Validation refusée |
| Word : `Audit_TestsSansReseau()` puis recette U2 | Socle au moins `34`, puis U2 `17/17` ; aucun appel API/NAS/imprimante par la suite hors réseau |
| `[SORTIE]` avec `ExportActif=0`, `Dossier` et `NomFichier` | Configuration acceptée, aucune copie secondaire |
| Clé `[SORTIE]` absente ou valeur inconnue | Erreur explicite ; aucun export silencieusement désactivé |
| Source ou binaire modifié après validation | Activation refusée ; nouvelle préparation/validation nécessaire |
| Jeton médecin utilisé comme secrétariat | Mutation des patients/agenda refusée par le serveur |
| NAS ou HTTPS inaccessible | Message d'erreur ; aucune sélection d'une ancienne arrivée locale comme si elle était actuelle |
| Nouvelle fiche : nom, prénom, DDN, sexe, coordonnées et médecin | Une fiche NAS ; sexe inconnu non inventé, date impossible refusée |
| Même fiche ouverte sur deux postes | La seconde sauvegarde avec ancienne révision est refusée |
| Deux prises du même créneau | Une seule réussit |
| Deux RDV distincts du même patient le même jour | Deux arrivées distinctes dans le cache et le service |
| Deux médecins sélectionnent la même arrivée | Une seule réservation réussit |
| A puis dictée du destinataire, B puis corps, C | Signets et formule d'appel corrects ; identité/âge du patient réservé |
| Export GDT, accents et identité fictive | Bon patient dans Resting12Lead ; aucune lettre ni import attribué à une autre fiche |
| D : sortie complète de l'API, puis relecture et D | Texte et annexes proposés ; rien dans la file avant confirmation ; toutes les pages relues |
| Réponse API incomplète, refus, JSON invalide, annexe vide | Erreur explicite ; aucun courrier incomplet publié |
| Négation/dose/nombre changé | Différence signalée ; relire même si aucun signal n'apparaît |
| Homonymes ou même alias de spécialistes | Pas de choix implicite ; sélection d'un ID précis et vérification de l'adresse |
| Coupure après sauvegarde brouillon | Reprise depuis le compte propriétaire, sans nouvelle consultation |
| Échec copie DOCX ou export PDF | Aucun nouvel élément exploitable dans la file du secrétariat |
| Coupure avant/après commit de publication | Une seule version en attente pour la consultation ; aucune double séance ; rapprochement des fichiers non référencés |
| Nouvelle correction après publication | Nouvelle version du courrier ; ancienne file remplacée ; montant déjà facturé conservé |
| Tarif ancien ou altéré côté client | Première facturation refusée jusqu'au rechargement de la nomenclature |
| Réimpression après modification de la sélection d'actes | Actes et montants de la séance enregistrée, pas de la nouvelle sélection |
| Réimpression après correction de l'identité courante | Identité et libellés CERFA figés lors de la facturation, sans substitution par la fiche ou la nomenclature actuelle |
| Réimpression demandée depuis un autre patient | Refus explicite |
| Facturation Patient/Organisme différente selon les actes | Chaque bundle conserve son payeur ; une répartition incomplète est refusée |
| Tiers payant puis encaissement | Non payé avant règlement, payé après commande d'encaissement ; date/mode enregistrés |
| Fin de traitement secrétariat | File, consultation et agenda « Honore » mis à jour ensemble |
| CERFA patient assuré / assuré distinct | Identités et NIR corrects ; positions absentes et calage non validé bloquent l'impression |
| CERFA plus de quatre lignes | Refus explicite, aucune ligne ignorée silencieusement |
| Deux imprimantes/postes | Calage local indépendant ; essai papier vérifié |
| Sauvegarde, contrôle et restauration sur volumes neufs | Base et fichiers récupérés, droits corrects, rapprochement puis lecture d'échantillons réussis |
| Arrêt du projet U2b | Seul `cabinetcardio-test-u2` s'arrête ; le service clinique reste disponible |
| Restauration du PC d'essai | Simulation puis application réussies ; ancien modèle de nouveau utilisable |

Avec le lanceur autonome, fermer les applications après les essais et saisir **RECETTE** dans sa console. Il exécute la validation et l'activation sans commande à recopier. Pour tester le modèle Word, utiliser **Fichier > Ouvrir** sur le `.dotm` préparé ; un double-clic dans l'Explorateur crée un nouveau document. Les macros des fichiers ouverts par le constructeur sont désactivées : fermer puis rouvrir le fichier pour les essais, en respectant les autorisations Office du poste.

La recette U2b n'est recevable que si les deux compilations Office réussissent, si le socle Word annonce au moins `34`, si l'extension Word U2 annonce `17/17`, si Excel annonce `25/25`, si le parcours fictif complet est validé et si le retour à l'installation précédente a été testé sur le PC d'essai.

Sur une installation neuve, les essais en réseau exigent auparavant un environnement d'essai configuré (service, partage, compte et configuration locale du client). Choisir **PAUSE** si cet environnement n'est pas disponible ; le lanceur ne crée pas un serveur de test sur le PC et ne confond pas essai et production.

En mode manuel, après compilation et essais, fermer les applications et lancer `Build/valider_preparation.ps1` avec les commutateurs correspondant au profil, comme indiqué dans le guide d'installation. Il compare le code des binaires rouverts à la source et fige leurs empreintes ; il ne remplace pas les essais du tableau.

Les essais de l'API externe se font avec des textes fictifs. Le corpus doit couvrir : absence d'annexe, une annexe, plusieurs annexes vers le même spécialiste, homonyme, identité absente, négation, dose décimale, unités, examen seulement évoqué et examen effectivement demandé. La décision médicale de demander un examen doit rester celle du médecin.
