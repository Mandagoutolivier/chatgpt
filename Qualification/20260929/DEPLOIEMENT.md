# Migration et bascule U2c — procédure préparée le 29 septembre 2026

## Point de départ

Source produit : `b40e7d94aabd8536544f0d23138042fbe9bf7160`. Binaires et résultats : [QUALIFICATION.md](QUALIFICATION.md). Les profils du paquet sont préparés mais non activés. Ne pas recopier la base de recette contenant des patients fictifs dans le cabinet.

L'URL `https://DS224:8444`, le partage CabinetCardioTestU2 et le projet `cabinetcardio-test-u2` sont des cibles de recette. Les noms, URL, volumes et comptes du futur service clinique doivent être relevés sur son installation effective avant toute commande de bascule. Aucun emplacement clinique n'est présumé ici.

## 1. Conditions d'activation encore ouvertes

| Contrôle | Critère de passage |
|---|---|
| ECG réel sur un profil de test isolé | Identité, quatre dates, âge après contournement manuel et sexe exacts ; fonctionnement du récepteur Bluetooth sans blocage pendant un essai représentatif ; aucune écriture dans la base clinique |
| CERFA matériel | Patient assuré et assuré distinct, NIR fictifs, alignement et marges corrects sur chaque imprimante/poste ; plus de quatre lignes et calage absent toujours bloqués |
| Dictée et ergonomie | Parcours au micro/Dragon et commandes A/B/C/D, reprise Word et correction RDC dans les conditions usuelles ; tous les documents et annexes visibles |
| Validation médicale | Corpus réel de recette fictive couvrant aucune/une/plusieurs annexes, homonymes, négations, doses, unités et examen seulement évoqué ; relecture de toutes les pages et acceptation par le médecin |
| Infrastructure clinique | Inventaire du service et des volumes, accès stricts, TLS, espace libre, sauvegarde hors du même NAS et politique de conservation vérifiés |

Les suites automatiques ne cochent pas ces lignes. Ne pas exécuter `valider_preparation.ps1 -RecetteValidee` ni saisir `RECETTE` dans le lanceur tant qu'elles ne sont pas réellement satisfaites. Ce verrou correspond à `RECETTE_WINDOWS.md`, pas à une demande supplémentaire d'autorisation administrative.

## 2. Préparer la migration sur une copie

1. Relever versions applicatives, schéma, projet Compose, volumes, URL, comptes NAS/SMB/API, profils Windows/Office, chemins des modèles et imprimantes. Ne pas exporter les secrets dans un rapport public.
2. Définir une fenêtre sans écriture. Fermer les applications de l'installation à migrer, arrêter uniquement son API, produire la sauvegarde base/fichiers/configuration avec `TERMINE`, inventaires et SHA-256. Capturer aussi les ACL Synology avec propriétaires, groupes, modes, xattrs et indicateurs. Redémarrer le service initial si la bascule n'est pas immédiate.
3. Restaurer sur une base vide, un dossier neuf et un réseau isolé sans port clinique. Employer le moteur corrigé. Restaurer les ACL par l'étape native puis vérifier les accès SMB depuis les vrais comptes limités.
4. Exécuter la simulation du module de migration approprié au schéma constaté. Lire toutes les erreurs et avertissements, contrôler identités, correspondants, doublons, historique, facturation et montants. Conserver le plan et son empreinte.
5. Appliquer exclusivement le plan dont l'empreinte a été relue. Pour l'import initial : `python -m cabinet.migration --appliquer --empreinte-validee EMPREINTE_DE_LA_SIMULATION`. Pour la conversion U1 : `python -m cabinet.migration_u1 --appliquer EMPREINTE_DU_PLAN`. Utiliser les commandes et le service Compose explicitement définis dans `Serveur/INSTALLATION_NAS.md` ; aucune commande générique ne doit viser un projet par défaut.
6. Comparer comptages, montants, règlements, références et empreintes avant/après ; ouvrir des échantillons ; vérifier qu'une relance ne duplique rien. Une modification des sources ou du plan impose une nouvelle simulation.

La migration technique a été testée sur données fictives. La migration des données du cabinet n'a pas été exécutée dans cette intervention.

## 3. Préparer les postes et le retour arrière

1. Copier le paquet qualifié et contrôler `SHA256SUMS.txt` avant toute utilisation. Conserver les anciens modèles/classeurs, leurs SHA-256, configurations privées et DACL dans une sauvegarde inaccessible aux autres utilisateurs.
2. Choisir `CabinetMedecin` pour AX8_MAX et `CabinetSecretariat` pour RDC. Configurer les URL et partages de la cible effective, le certificat, le dossier ECG et les imprimantes. Employer des comptes API/SMB distincts et limités ; ne pas reprendre le compte de recette à double rôle.
3. Comparer les sources et références dans Office et vérifier les profils. Toute recompilation modifie potentiellement le binaire : recalculer et figer les empreintes après la validation.
4. Effectuer les essais physiques et médicaux de la section 1. Fermer complètement Office, restaurer les permissions VBA et `Normal.dotm`, puis seulement valider la préparation avec les commutateurs correspondant aux résultats réellement obtenus.
5. Tester le retour arrière sur copies avant activation : simulation puis application de `Build/restaurer_poste.ps1`, anciennes empreintes/DACL, réouverture et macro témoin. Ce test a réussi sur les deux postes pour la livraison présente.

## 4. Bascule du cabinet

1. Suspendre les écritures et effectuer une sauvegarde finale cohérente avec capsule ACL ; consigner l'heure et les dernières opérations. Préserver l'installation précédente pour le retour arrière.
2. Appliquer la migration validée aux données finales. Contrôler à nouveau le plan et son empreinte si les données ont changé depuis la répétition. Ne pas réutiliser aveuglément une ancienne empreinte.
3. Démarrer le service U2c sur ses volumes cliniques définis. Vérifier `/health`, puis `whoami` authentifié : protocole 2, schéma 2, révision U2c et rôle attendu.
4. Activer les seuls dossiers préparés et validés avec l'installateur documenté. Conserver les anciens fichiers et les reçus d'installation.
5. Depuis chaque poste, contrôler la file, les identités, les références d'archives, les droits refusés, les montants, les imprimantes et l'ECG. Effectuer le test autorisé sans introduire de facturation ou acte réel fictif dans les systèmes tiers.
6. Consigner la décision de mise en service, les empreintes finales, l'heure, le responsable, les résultats et la localisation des sauvegardes. Réouvrir les écritures seulement après contrôles concordants.

## 5. Déclenchement du retour arrière

Arrêter les écritures en cas d'identité incohérente, archive altérée/manquante, montant changé, doublon, droits excessifs, perte d'accès ou panne bloquante du parcours.

- Avant toute nouvelle opération clinique : restaurer le service/base/fichiers/configuration/ACL cohérents de la sauvegarde finale, puis les anciens modèles des postes ; vérifier les mêmes contrôles de reprise.
- Si des opérations cliniques ont été créées après la bascule : sauvegarder d'abord cet état, inventorier les opérations nouvelles et organiser leur reprise contrôlée. Restaurer une vieille base sans ce rapprochement ferait perdre ces opérations ; ce n'est pas un retour arrière acceptable.
- Tester la lecture d'archives et la conservation des montants, puis consigner incident, période affectée et résultat. Garder l'état échoué pour diagnostic privé.

## 6. Exploitation après bascule

Prévoir une sauvegarde cohérente régulière, une copie protégée indépendante du NAS, une rétention définie, une alerte d'échec et des restaurations périodiques sur cible isolée. Ces automatisations de production et leur destination n'ont pas été créées par la présente recette. Après toute mise à jour Office/ECG/service, rejouer les contrôles concernés et tenir le mémo à jour.
