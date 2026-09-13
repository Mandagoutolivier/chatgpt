# Préparer les essais de la version déjà compilée

`Configurer_Recette_Domicile.ps1` configure la session Windows pour tester la préparation **Domicile 17c84b98374f54c076930583796e50602458d294** existante. Il répond au cas où Word et Excel ont été compilés, mais où les essais réseau n'ont pas pu être réalisés avant la création du service NAS.

Ce script ne constitue pas une nouvelle version du logiciel et n'active pas les compléments Office. Il s'utilise en dehors du cache immuable des sources. Le lanceur principal reste inchangé.

## Conditions

- Utiliser la session Windows habituelle, Word et Excel complètement fermés.
- Le fichier `%APPDATA%\CabinetCardio\Assistant\Domicile.json` doit décrire la préparation 17c84b9, en phase `valide` ou `compile`, avec les deux compilations mémorisées.
- Les empreintes des sources et des trois fichiers préparés doivent encore correspondre au reçu de préparation. Le script ne régénère aucune empreinte pour accepter un fichier modifié.
- Le service de test doit répondre en HTTPS à `https://192.168.10.1:8443`, avec un certificat accepté par Windows.
- Le compte `domicile-test` doit être actif et posséder les rôles `medecin` et `secretariat`.
- Le partage `\\DS224\CabinetCardioTest` doit être accessible et son réglage `[SORTIE] Dossier` doit désigner `\\DS224\CabinetCardioTest\sortiedragon`.

## Exécution

Enregistrer le script sur le Bureau, puis lancer dans PowerShell :

```powershell
& {
    $scriptRecette = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Configurer_Recette_Domicile.ps1'
    if (-not (Test-Path -LiteralPath $scriptRecette -PathType Leaf)) {
        throw 'Enregistrez le script sur le Bureau avant de continuer.'
    }
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptRecette
}
```

Copier le jeton depuis le gestionnaire de mots de passe **lorsque la saisie masquée apparaît**, après avoir copié la commande. Il ne faut pas mettre le jeton dans la ligne de commande ni le transmettre dans une capture.

Le script vérifie le compte avant de changer les réglages. Il prépare une sauvegarde des fichiers précédents et un journal `restauration.json`, puis configure :

| Élément | Valeur pour la recette |
|---|---|
| Racine des ressources | `\\DS224\CabinetCardioTest` |
| Service | `https://192.168.10.1:8443` |
| Compte | `domicile-test` |
| Profil | `Domicile` |
| Dossier GDT local | `%LOCALAPPDATA%\CabinetCardioTest\GDT` |
| SQLite | Copie du programme déjà vérifié vers `%APPDATA%\CabinetCardio\Tools\sqlite3.exe` |
| Recette | `false` dans l'état de l'assistant et le reçu ; phase `compile` |

Le jeton est enregistré dans le fichier attendu par le client VBA, avec des permissions Windows limitées à l'utilisateur courant, SYSTEM et aux administrateurs. Ce fichier n'est pas chiffré par ce script. Les sauvegardes de fichiers existants ont également des droits restreints.

Les données PostgreSQL et les documents du NAS ne sont pas modifiés. Les fichiers Word et Excel préparés ne sont pas reconstruits. Les changements concernent cependant la configuration **de cette session Windows** : tant qu'elle pointe sur le service de test, utiliser cette session pour les essais prévus.

En cas d'exception pendant les écritures, le script tente de rétablir tous les fichiers précédents et leurs ACL. Une interruption forcée de PowerShell ou de Windows peut empêcher cette restauration automatique : conserver alors le dossier de sauvegarde indiqué, avec son journal. Ne pas le publier, car il peut contenir un ancien jeton.

## Après configuration

Le fichier `service.url` doit contenir l'adresse seule, sans retour à la ligne. Le lecteur VBA 17c84b9 utilise `Trim$`, qui ne supprime pas CR/LF. Le script de recette corrigé écrit désormais cette forme exacte ; l'installateur de la branche de correction est corrigé également. Un ancien lanceur 17c84b9 utilisé ensuite pour l'activation réintroduira sa fin de ligne : normaliser à nouveau `service.url` après cette activation, sans modifier le cache des sources ni les binaires compilés.

Effectuer les essais applicables du guide `versions/cabinet-2026.09.12/RECETTE_WINDOWS.md` avec des identités fictives. Pour le modèle Word préparé, commencer par **Fichier > Ouvrir**, et non un double-clic qui crée un document dérivé. L'accès HTTP authentifié ne démontre pas encore le fonctionnement du transport WinHTTP utilisé par VBA, de Dragon, de l'ECG ou de l'impression.

Le nouveau dossier GDT évite d'envoyer les identités fictives dans `C:\Mandagout`. Le test d'import dans Resting12Lead demandera de sélectionner explicitement ce dossier de test dans son interface GDT et de noter le réglage précédent pour le rétablir ensuite.

Ne déclarer `RECETTE` dans le lanceur que lorsque les essais requis ont réellement été exécutés. La relance du script de configuration conserve les compilations mais remet toujours la recette à effectuer ; elle ne sert pas à valider les essais.

## Vérification du correctif

`Test-ConfigurerRecette.ps1` s'exécute sous Windows PowerShell 5.1 avec fichiers et ACL réels. Les réponses NAS/HTTP et la saisie du jeton sont simulées. Il contrôle la configuration réussie et sa relance, le refus HTTP avant modification, la restauration après un échec d'écriture tardif et le refus d'un binaire modifié avant la demande du jeton. Ces tests ne remplacent pas la recette Word/Excel sur le poste.

Le fichier `Fixtures/outils_installation_17c84b9.ps1` est une copie immuable du vérificateur de la version 17c84b9, réservée à ces tests.
