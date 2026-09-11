# Recette Windows — à exécuter avant utilisation

**Statut : non exécutée pendant l'audit.** Les tests Linux du dépôt ne remplacent pas cette recette. Utiliser des patients fictifs et un partage NAS de test, y compris pour la sortie des courriers. Garder les résultats avec le commit testé, les versions Office/Windows et la version Resting12Lead.

## Construction et compilation

| Essai | Résultat attendu |
|---|---|
| Construire Word et Excel avec les trois profils | Fichiers générés, originaux inchangés, sauvegardes présentes |
| Alt+F11 → Débogage → Compiler dans les deux projets | Aucune erreur de syntaxe, symbole, type ou référence ; noter Office 32/64 bits |
| Inspecter Outils → Références | Aucune référence « MANQUANT » |
| Fermer, rouvrir et compiler de nouveau les modèles construits | Projet issu des sources relu correctement ; ne pas réutiliser directement les anciens binaires |
| Ouvrir chaque formulaire et le ruban | Contrôles, événements et accents corrects ; pas de collision avec Normal.dotm |
| Exécuter `Audit_LancerTests` dans Word | Réussite des tests de date, sexe, JSON, GDT, noms de fichiers et repérage du corps, sur ressources temporaires |
| Provoquer une erreur d'installation avant puis après une copie active | Ancienne version conservée/restaurée, échec annoncé, journal exploitable |

Les anciennes procédures `Test_*` des modules `modTests_Word` et `modTests_Excel` sont volontairement bloquées : elles reposaient sur d'anciennes fixtures et pouvaient écrire dans les fichiers de travail. Les modules PROD dont le nom contient `Test` ne sont pas tous des tests : certains contiennent encore la logique de production des destinations.

## Identité, arrivée et ECG

| Essai | Résultat attendu |
|---|---|
| Date impossible (31/02), sexe vide/inconnu, identité incomplète | Message de validation, aucune identité inventée |
| Noms accentués, prénom composé, anniversaire aujourd'hui/demain | Identité et âge exacts dans Word et GDT |
| Un patient avec deux rendez-vous dans la même journée | Deux entrées distinctes dans le cache |
| Deux médecins sélectionnent simultanément la même arrivée | Un seul peut réserver l'arrivée |
| Annuler la sélection A | Aucune consultation créée pour le premier patient de la liste |
| Patient modifié sur le NAS après ouverture de la lettre | Envoi bloqué jusqu'à résolution de la discordance |
| Coupure VPN/NAS, puis nouvelle sélection | Pas d'utilisation silencieuse de l'ancienne file locale |
| Interruption entre réservation et création du brouillon | Réservation visible dans la reprise, pas de disparition définitive de la file |
| A puis raccourci destinataire, B, dictée, C | Bon destinataire, bons signets, identité du patient courant |
| GDT importé dans Resting12Lead | Nom, prénom, date de naissance et sexe exacts ; pas de conservation du patient précédent |

## Courriers, API et correspondants

| Essai | Résultat attendu |
|---|---|
| Correction simple avec médicaments/anatomie au début et à la fin du corps | Texte conservé, gras selon Cabinet, formule de politesse et signature correctes |
| Lettre nécessitant plusieurs examens/spécialistes | Une annexe par demande effectivement retenue et validée, bon destinataire, sans duplication |
| Formule de politesse, en-tête, signature et modèle à corps vide | Insertion dans les signets du modèle, aucun effacement de l'en-tête |
| Deux correspondants homonymes ou même clé avec adresses différentes | Ambiguïté signalée ; vérification explicite du destinataire principal et de chaque annexe |
| Ajouter une structure après la ligne 499 de l'annuaire | Formules de recherche étendues ; adresse correcte dans aperçu et enregistrement |
| API : 401, 429, 5xx, réponse vide/incomplète/refus, JSON invalide | Erreur lisible ; pas de faux succès ni de dépôt d'un lot incomplet au secrétariat |
| Échec du deuxième appel API | Brouillon préalable récupérable ; vérifier le document avant toute reprise |
| Double pression D pendant le traitement | Un traitement actif, pas de double facturation ni d'annexes doublées |
| Réponse modifiant négation, dose ou identité | Détecter à la relecture médicale ; le code ne garantit pas ce contrôle sémantique |
| Échec de l'export PDF ou coupure NAS pendant la publication | Pas d'événement de succès avant les sorties ; vérifier les fichiers orphelins et la reprise |
| Reprise puis nouvelle version du courrier | Révision distincte, identifiant de consultation stable, pas de deuxième séance |

## Secrétariat, agenda et journal

| Essai | Résultat attendu |
|---|---|
| Deux réservations simultanées d'un créneau se chevauchant | La deuxième écriture est refusée sous verrou |
| Rendez-vous annulé, puis patient signalé arrivé | Pas de réapparition d'un événement d'arrivée annulé |
| Rendez-vous décembre puis traitement en janvier | Agenda et journal utilisent l'année pertinente ; pas de collision d'identifiant |
| Classeur NAS ouvert en lecture seule ou réseau indisponible | Échec explicite ; aucune annonce trompeuse de sauvegarde |
| Double clic de validation, réouverture du même événement | Une seule séance enregistrée ; un autre patient ne peut pas réutiliser son identifiant |
| Nouvelle impression d'une séance existante | Montants et actes relus dans le journal, pas dans une nouvelle sélection de formulaire |
| Erreur `PrintOut` / imprimante hors service | Pas de marquage automatique en cas d'erreur renvoyée ; contrôler physiquement le papier même si le spouleur accepte |
| Plus de quatre lignes de soins | Refus explicite ; pas de lignes perdues silencieusement |
| Tiers payant et paiement ultérieur | Pas de recette considérée encaissée sans encaissement réel |
| Mineur/ayant droit, NIR, RPPS et numéro AM | Vérifier la distinction patient/assuré et le formulaire papier ; les cas non gérés doivent rester manuels |

Le module fourni ne réalise pas la télétransmission SESAM-Vitale/CPS. La feuille papier, le journal et les tarifs doivent être confrontés au fonctionnement réel du cabinet.

## Acceptation

Consigner chaque résultat, toute anomalie et sa correction. Refaire le scénario complet secrétariat → A/B/C/D → ECG → courriers → journal → nouveau rendez-vous avec deux comptes Windows, puis depuis le domicile par VPN. Conserver une sauvegarde NAS restaurable et une copie des modèles précédents. Un échec de compilation, une discordance d'identité, une erreur de destinataire ou un double acte empêche l'acceptation de cette version.
