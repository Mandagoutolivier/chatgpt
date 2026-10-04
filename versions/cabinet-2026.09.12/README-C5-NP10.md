# CabinetCardio — recette C5 NP10

3 octobre 2026. NP10 reprend le parcours « un seul D » de NP9 et corrige le nom des copies lisibles conservées dans le dossier patient.

## Nom des courriers validés

Les nouvelles copies DOCX et PDF du dossier patient portent désormais :

`NOM Prénom aaaammjjhhmm.docx`

`NOM Prénom aaaammjjhhmm.pdf`

Les caractères interdits par Windows sont remplacés. Si plusieurs courriers du même patient sont validés dans la même minute, les suivants reçoivent les suffixes ` 2`, ` 3`, etc. Aucun fichier existant n'est écrasé. Les archives techniques immuables du dossier `Documents` conservent leurs noms SHA-256.

## Vérifications

- Qualification Office isolée : succès le 3 octobre 2026 à 21:47:47 UTC, 469 assertions, dont 70/70 ciblées sur le parcours Dragon et le nommage.
- Contrôle de démarrage connecté NP10 : succès ; modèle chargé, zéro document au démarrage.
- Essai connecté sur le patient fictif `RECETTE4 beta`, consultation `consult-R722ef13f71374f69b8c20df7df9f17cd`.
- Fichiers réellement créés : `RECETTE4 beta 202610032359.docx` et `RECETTE4 beta 202610032359.pdf`.
- Une publication reçue par le secrétariat : `4c151b4e-5d34-4b94-be5b-8b708e2929ec`.
- Archives immuables : DOCX `03ff545d11ad16c6201cc1cb8f4b6b872853d56abf51ea617a33258f7fd8b1ea`, PDF `998898dad527e9ec67ca50899445cc79bb4bd6ba85847206dc72a91c5db48bde`.

## Modèle

`CabinetDragonC5NP10.dotm` : 497 652 octets, SHA256 `FF183F42D591BE91CEE2146AF2260439157957C3A5D4FD13D9328B31F19C2C7C`.

La production clinique et le profil `AX8_MAX\olivi` n'ont pas été activés.
