# U1 — correctifs et qualification du 14 septembre 2026

Livraison de développement sur `codex/urgence-1-20260914`, issue de U0 `10dcccbcee2ad926a6dbcc6235f2ad7d2095e84f`. [PR en brouillon nº 3](https://github.com/Mandagoutolivier/chatgpt/pull/3).

**Les correctifs U1 sont implémentés et les essais automatisés ci-dessous ont été exécutés. La qualification du parcours complet au cabinet reste à faire avant toute activation clinique.** L'installation habituelle est conservée.

## Correctifs livrés

| Domaine | Comportement livré |
|---|---|
| Contrat NAS et client | Contrats stricts par opération avant transaction, erreurs explicites, JSON typé UTF-8, pagination bornée, commandes persistées par identifiant pour retrouver un résultat après interruption. Arrivées filtrées et triées en mémoire ; anciens scripts SQLite retirés du dossier actif, historique Git conservé. |
| ACTES et reprise serveur | Identifiants stables, codes explicites, alias contrôlés, conservation du dépassement existant ; blocage d'un dépassement non nul dont la règle n'est pas qualifiée. Migration de schéma 2 précédée d'une simulation et de l'approbation de son empreinte ; audit des valeurs avant/après. Sauvegarde/restauration des schémas 1 et 2, historique ACTES et règlements compris. |
| Courriers et reprise | Source immuable et état local chiffrés par DPAPI, écriture atomique et verrou par consultation. Reprise explicite en cas d'appel IA incertain. Destinations conservées par annexe, identité réhydratée, révision documentaire stable, publication préparée DOCX/PDF et relecture imposée. Contrôles des nombres, unités, négations et marqueurs masqués. |
| Facturation, impression et règlements | Publication attendue vérifiée avant facturation, lignes enregistrées faisant référence, contrôles CERFA avant écriture, impression déclarée incertaine avant envoi physique puis confirmation explicite. Règlement protégé par empreinte optimiste et confirmation de la séance, avec audit ; exercice explicite. |
| Agenda et intégration Office | Transitions atomiques et réactivation contrôlée, anciennes entrées d'arrivées conservées en lecture seule, un seul rappel Excel. Recherche Word avec restauration des paramètres et conservation des critères de format ; contexte des raccourcis restauré. |
| Installation et qualification | Fusion INI préservant les réglages locaux, sauvegarde protégée et retour arrière des octets/droits, précontrôle des composants inconnus, conservation des paquets modifiés. Lanceur lié à un commit précis et vérifié contre ses objets Git. Recette Office sur copies avec restauration de l'accès VBA et de Normal.dotm. |

La présence de ces protections ne constitue pas une validation du sens médical d'un courrier, du calage papier ni de la compatibilité complète des périphériques.

## Essais effectués sur AX8_Max

Remote Desktop Commander, session habituelle, copies dans :

`%LOCALAPPDATA%\CabinetCardioTest\U1-20260914\Preparation-U1-essai9`

| Contrôle | Résultat |
|---|---|
| Construction et compilation du modèle Word | Réussies |
| Exécution réelle VBA Word : SHA-256, DPAPI, JSON, états persistants, révision, contrôles de texte et recherche | **34 contrôles réussis**, aucune erreur |
| Construction et compilation du classeur Excel | Réussies |
| Exécution réelle VBA Excel : règles locales, sélections et contrôles de saisie | **10 contrôles réussis**, aucune erreur |
| Installation sur fichiers fictifs : fusion INI, conservation du paquet, sauvegarde et retour arrière des droits/octets | **9 contrôles réussis** |
| Comparaison SHA-256 de l'installation existante après recette | **45/45 fichiers conformes** : 5 configurations et 40 fichiers Word |
| Fermeture et restauration après recette | Aucun Word/Excel restant ; accès VBA restauré, journal de restauration retiré |

La recette ne déclenche ni appel IA, ni écriture NAS, ni impression clinique. Le journal final est `validation-office-finale-4482aecaa0f34ae0998ee8df8b516be1.log`. La comparaison est conservée dans `verification-existant-finale-u1.json` à la racine de l'espace U1 sur AX8_Max (2026-09-14 21:53 UTC).

Office a réécrit Normal.dotm pendant les premiers essais. La copie modifiée a été conservée, puis les octets originaux sauvegardés avant U1 ont été rétablis et vérifiés. Le harnais final automatise cette protection ; l'essai 9 l'a exercée avec succès. Les modifications locales antérieures restent dans les fichiers actifs et dans `Sauvegarde-avant-U1`.

Les essais Windows ont révélé et permis de corriger les arguments COM par référence, la lecture non prise en charge de StatusBar, l'identifiant Windows SHA-256 signé, le traitement du texte vide, ainsi que les marqueurs Word de navigation et d'identification qui rendaient instable la révision documentaire.

## Validation GitHub

Le [contrôle du commit e17f9fd](https://github.com/Mandagoutolivier/chatgpt/actions/runs/34897323767) a réussi ses quatre tâches :

- **78 tests serveur sur PostgreSQL natif**, dont concurrence.
- Sauvegarde puis restauration réelle en conteneurs isolés, schémas 1 et 2 et migration ACTES sur données fictives.
- Construction et exécution de l'image avec droits restrictifs et sans utilisateur root.
- Contrôles PowerShell Windows 5.1 et lanceur autonome.

Ces preuves concernent ce commit intermédiaire. Les ajustements Office, la protection de Normal.dotm, les droits Windows et le lanceur final sont soumis à nouveau au même workflow sur la tête de la [PR nº 3](https://github.com/Mandagoutolivier/chatgpt/pull/3). Le statut du commit final doit être vert avant utilisation de la livraison ; ne pas déduire sa réussite de la seule présence du workflow.

`AUDIT.md` et `Tests/verification_livraison.json` décrivent une livraison antérieure. Le présent document et les exécutions liées à la PR portent la qualification U1.

## Ce qui reste avant une mise en service

- Qualifier le parcours complet Word → NAS/SMB → Excel sur données fictives, avec coupure/reprise réseau et les comptes du cabinet.
- Vérifier les commandes Dragon/PowerMic, les formulaires et l'échange ECG sur identité fictive.
- Qualifier NIR/assuré distinct, paramètres praticien, imprimante et calage CERFA, puis confirmer les sorties papier.
- Valider les données ACTES importées, les alias ambigus et la règle des dépassements ; ne pas appliquer une simulation présentant une anomalie non comprise.
- Tester la sauvegarde/restauration et le retour arrière sur le NAS de recette avant une bascule approuvée.

Aucune migration U1 du NAS existant, activation dans STARTUP, modification de l'ECG ou utilisation de données patient réelles n'a été effectuée pour cette recette. U2 reste un travail distinct à poursuivre avant de décider la mise en service.
