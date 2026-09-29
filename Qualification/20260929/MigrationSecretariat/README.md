# Migration du secrétariat vers RDC

## État du 29 septembre 2026

- Cible : `RDC`, adresse `192.168.10.102`, futur profil quotidien
  `RDC\PATRICIA`.
- Source identifiée sur le réseau : `ACCUEIL`, adresse `192.168.10.104`,
  adresse MAC `68-1D-EF-1B-BA-92`.
- `ACCUEIL` répond au ping ; SMB 445 et RDP 3389 sont ouverts. WinRM 5985
  est fermé.
- Le compte de recette `RDC\CabinetU2Test` voit les partages publiés par
  ACCUEIL mais n'a pas accès au profil personnel ni aux dossiers système.
- RDC possède déjà Office 2016, Thunderbird, Acrobat, les pilotes Brother,
  AccuWin, Quick BP/Quick Reader, Resting 12-Lead et EasyScope. La présence
  d'un logiciel ne prouve pas que sa licence, ses réglages ou ses données
  utilisateur sont prêts dans `RDC\PATRICIA`.

## Ordre obligatoire

1. Exécuter `Inventorier-PosteSecretariat.ps1` dans la session utilisée
   quotidiennement sur ACCUEIL. Le script ne copie aucune donnée et ne lit
   aucun secret.
2. Vérifier une sauvegarde complète et restaurable d'ACCUEIL sur le DS224+.
3. Ouvrir `RDC\PATRICIA`, inventorier son profil et sauvegarder son état.
4. Comparer applications, licences, messagerie, raccourcis, lecteurs réseau,
   imprimantes, scanner et réglages métier.
5. Sauvegarder puis reproduire le Bureau d'ACCUEIL : tous les éléments et
   raccourcis, leurs icônes, ainsi que leur disposition lorsque la résolution
   et la mise à l'échelle de RDC sont compatibles.
6. Copier seulement les autres éléments nécessaires, puis effectuer la bascule
   après arrêt des écritures sur ACCUEIL.
7. Conserver ACCUEIL intact et disponible comme retour arrière jusqu'à la
   validation des tâches quotidiennes.

Le profil `RDC\CabinetU2Test` et le partage `CabinetCardioTestU2` restent
réservés à la recette. Ils ne doivent pas devenir le profil clinique du
secrétariat.
