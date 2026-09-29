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
- `RelectureEnAttente=1` : préparation réussie ; aucune validation humaine du
  contenu ni transmission ne peut être déduite de cet état.
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
