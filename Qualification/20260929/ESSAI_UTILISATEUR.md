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
