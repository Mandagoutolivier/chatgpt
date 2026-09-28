# Qualification technique CabinetCardio U2c — 29 septembre 2026

## Décision et périmètre

Travaux exécutés dans la nuit du 28 au 29 septembre, heure de Paris, sur les profils Windows de recette AX8_MAX et RDC, le partage CabinetCardioTestU2 et PostgreSQL isolé sur le NAS. Aucune activation clinique. Les préparations restent sans `recetteValidee`.

Les contrôles logiciels et la restauration décrits ici ont réussi. Les essais matériels ECG, le calage papier CERFA et la validation médicale complète restent à effectuer avant usage clinique. Ce document décrit des preuves techniques, pas une certification médicale.

## Sources et binaires

- Branche : `codex/u2c-qualification-20260929`.
- Source produit qualifiée : `b40e7d94aabd8536544f0d23138042fbe9bf7160`.
- Intégration Word et Excel : `7e3a6a0f288e691309ccec276a0d4503766c4402`.
- Correctif restauration et tests Windows : `a1524658e6fd7080ed924d96d36faef827cfbdab`, puis correction d'une attente du test UNC dans `b40e7d9`.
- CI finale du produit : [run 36496767722](https://github.com/Mandagoutolivier/chatgpt/actions/runs/36496767722), réussite.
- Les sources `Src` et `Build` sont identiques entre la compilation `7e3a6a0` et `b40e7d9`. Les nouveaux changements concernent les tests et le moteur Python de restauration.
- Les lanceurs fournis sont régénérés depuis les objets Git de `b40e7d9` ; vérificateur des lanceurs réussi.

| Binaire | Modules de production | SHA-256 |
|---|---:|---|
| CabinetUnifie.dotm | 58 | `AC941632E4C5A16EC202E6067BB4E78CF978154FB4A7FDA4BB7FA6FD96770DF2` |
| Cabinet.xlsm | 27 | `3A975BDF3879016DE64C3495EED0A6C23050EAA1D774B2389D34D48FEC4355FD` |

Compilation VBA réelle, références Office/VBA/MSForms présentes, fermeture/réouverture et comparaison des sources normalisées réussies. Aucun module de test dans ces binaires. Les passerelles VBA utilisées pour piloter les parcours sont limitées à des copies de recette distinctes.

## Résultats

| Domaine | Exécution et preuve |
|---|---|
| Office sur chacun des deux postes | 399 contrôles : Word U1 50, U2 36, AuditGdt 88, ModeleCourrier 56, PresentationAnnexe 67, Excel 74, FileAnnexes 17, EditionSecretariat 11 ; aucun échec |
| Serveur | 226 tests réussis, aucun ignoré, sur PostgreSQL natif isolé ; 122,32 s pour la suite finale |
| Interruptions et concurrence complémentaires | 5 tests réussis : échec après copie DOCX, annulation avant commit SQL, réponse perdue après commit, reprise du brouillon après recréation du service, deux prises simultanées du même créneau |
| Contrôles statiques | 22 contrôles d'architecture et 11 de recette ; audits de 85 et 92 modules sans erreur |
| Windows | Suites PowerShell en CI réussies ; correctifs PS 5.1 de chargement Compression et de restauration DACL testés sur RDC ; restauration poste : 39 contrôles par PC |
| Rôles | Médecin seul sur AX8_MAX et secrétariat seul sur RDC ; 16 refus HTTP 403 réellement obtenus, 9 côté médecin et 7 côté secrétariat |
| Parcours fictif | RDV, arrivée, réservation, brouillon, correction de texte, courrier et annexe, relecture fictive, publication, changement de destinataire par le secrétariat, nouvelle version, facturation et clôture |
| Facturation | 14 contrôles HTTPS réels ; 26,50 € Patient et 5,00 € Organisme, total 31,50 € conservé ; paiements CB/Virement et audit, rejets des versions et empreintes périmées, du montant partiel, rejeux sans doublon |
| Archives | Anciennes versions conservées ; nouvelle version secrétariat `Relu=false` ; copies de consultation en lecture seule et contrôlées par empreinte |
| ECG GDT | Fichiers effectivement produits par le VBA : 4 dates valides, 4 dates absentes/impossibles/futures refusées sans modifier le fichier précédent |
| Sauvegarde/restauration | Sauvegarde cohérente puis restauration réelle dans des chemins et une base neufs ; 10 publications, 20 archives référencées, 6 brouillons |
| Droits restaurés | 226 objets avec propriétaire, groupe, modes, xattrs et indicateurs ACL Synology identiques ; sondes ordinaires Windows sur les 20 archives depuis les deux postes |
| Retour arrière PC | Simulation puis application sur copies des anciens binaires ; anciennes empreintes et DACL retrouvées ; réouverture Office et exécution d'une macro de contrôle sur chaque PC |

Les suites automatiques couvrent notamment le verrouillage concurrent, les révisions obsolètes, la reprise du brouillon propriétaire, le PDF manquant, les publications idempotentes et les migrations atomiques. Les 5 essais d'interruption complémentaires sont conservés dans `test_interruptions.py` et ajoutés à la CI ; leur journal accompagne la livraison. Il s'agit d'injections de fautes logicielles et de concurrence PostgreSQL, sans coupure électrique matérielle.

### Parcours du 29 septembre

Une nouvelle identité entièrement fictive et un nouveau RDV ont été utilisés. Publication initiale `a654e1cd-47bb-42b6-8c21-f436434b0fa7`, révision secrétariat `31a427ce152443d7ab314d5b7ca1aa13`. Le courrier principal et l'annexe ont été relus techniquement ; dose et négation conservées. La révision change le destinataire d'annexe et conserve le courrier principal. La consultation a été clôturée et le RDV est « Honore ».

Aucune impression papier déclarée réussie : `FeuilleSoinsImprimee=N`. Les publications historiques de recette restées en attente n'ont pas été traitées à la place du nouveau cas.

### ECG : limites précises

Chaque fichier valide contient 136 octets et 9 enregistrements : longueurs incluant CRLF, total `8100` exact, `8000=6302`, `9206=3`, `9218=02.00`, encodage Windows-1252 sans BOM. Accents `Œ` et `É` vérifiés sur les octets. `3103` contient `JJ.MM.AAAA` pour janvier, novembre, le 29 février 1980 et le 31 décembre. Le sexe `3110` n'est pas émis par cette implémentation.

Cela ne valide pas l'import dans Resting12Lead27, le recalcul manuel de l'âge, la saisie du sexe ni la stabilité du récepteur Bluetooth. Un dossier GDT séparé n'isole pas, à lui seul, la base patients du logiciel ECG. Aucune acquisition ni identité réelle n'a été utilisée.

## Restauration et ACL

Sauvegarde privée NAS : `cabinet-20260928T230245Z-2c6dca309556` dans le dossier de sauvegardes du projet de recette. Marqueur `TERMINE` et SHA-256 vérifiés. Arrêt bref de la seule API de recette pendant la sauvegarde, puis redémarrage confirmé. Base PostgreSQL et fichiers/configuration ont été sauvegardés sans activité Office.

La restauration retenue est `_RestaurationQualification20260929B`, base isolée `cabinet_restore_q29b`, aucun port publié. La première cible, sans suffixe B, conserve un diagnostic d'échec et ne constitue pas une restauration exploitable.

Le premier essai a révélé un défaut réel : une nouvelle racine UNC située sous l'ancien partage était prise à tort pour un ancien chemin non migré. Le correctif valide la nouvelle cible avec le contrôle de chemin existant, en maintenant les refus des chemins hors racine, traversées et préfixes voisins. Deux tests de régression sont intégrés au produit.

Le moteur générique ne restaure pas seul les ACL Synology : son indicateur `acl_synology_restaures=false` reste exact. Une étape native distincte a restauré les xattrs et indicateurs via `synoacltool`, puis comparé les 226 objets. La capsule ACL a pour SHA-256 `f9b833b38c48c6ee91b9fa74f79ceac8fb1065013d8f725340a34e22749f2bc7`.

Contrôles SQL après restauration : comptes 4, ressources 514, consultations 10, publications 10, commandes 95, séances 5, audits de règlements 2. Comptages identiques à la source ; contenu complet des séances et audits identique.

Sur chaque session Windows limitée : lecture et empreinte des 20 archives réussies ; écriture, suppression, modification de DACL et suppression via `Documents` refusées (code 5). Écriture/lecture/suppression d'une sonde fictive dans `Patients` réussies. Ces droits protègent contre les comptes utilisateurs testés ; ils ne rendent pas les archives immuables pour un administrateur NAS.

Un DOCX restauré a été ouvert dans Word RDC en lecture seule : identité, dose, annexe et SHA-256 exacts. Le corps et sa dernière page indiquent 2 pages. La statistique globale Word est restée à 4 après la sortie du mode Lecture ; cette anomalie de compteur n'a pas modifié le fichier. Vérifier la pagination imprimée lors du contrôle matériel. Les liens automatiques ont été désactivés dans cette ouverture de recette.

## Incidents de recette résolus et nettoyage

- Chargement PS 5.1 de `System.IO.Compression` explicité avant `FileSystem`.
- Application DACL de test corrigée pour éviter la demande indue du privilège SACL.
- Ancienne instance Excel de recette fermée avant la suite de restauration ; aucun arrêt forcé d'une application clinique.
- Adaptateurs de pilotage COM corrigés (`[ref]` et nom de variable PowerShell) ; formule d'appel ajoutée au scénario fictif avant sa correction. Les essais interrompus ne sont pas comptés comme réussis.
- L'historique des premiers journaux en échec est conservé ; les rapports de reprise et la CI finale font foi.
- En fin de recette : aucun Word/Excel, `AccessVBOM` absent pour les deux applications sur les deux profils, restaurations de `Normal.dotm` contrôlées. Jeton RDC strict vérifié ; ancien jeton gardé uniquement en sauvegarde DPAPI privée. Matériel RSA et fichiers de transfert temporaires supprimés.

## Preuves et livraison

Sur chaque PC : `%LOCALAPPDATA%\U2Q29` (Office1, roles-refus.json, RetourBinaire, restauration-smb.json, etat-final.json). AX8_MAX : Binaires1, gdt-octets.json. RDC : facturation-fictive.json, restauration-lecture-word.json.

Sur le partage de recette : `Patients\_Qualification20260929` (rapports sauvegarde/restauration, SQL, tests PostgreSQL, interruptions et copie du paquet). Les secrets, configurations privées, archives signées et données nominatives ne sont pas déposés dans GitHub.

Le paquet comprend les sources produit figées, les deux profils préparés, les lanceurs, cette qualification, la procédure de bascule et un manifeste SHA-256. `livraison.json` indique explicitement `Activation=false` et `RecetteClinique=false`. Voir [DEPLOIEMENT.md](DEPLOIEMENT.md) pour les étapes restantes.
