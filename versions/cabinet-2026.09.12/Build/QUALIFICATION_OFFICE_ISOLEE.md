# Qualification Office sur copies locales

`Tester_U2_Office_Isole.ps1` consolide les suites Word U1 et U2, le contrat du
papier à lettres, la présentation du courrier principal et les 70 contrôles
Excel. Il doit être exécuté dans un processus Windows PowerShell 5.1 neuf,
avec Word et Excel fermés, depuis des sources publiques vérifiées :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\SourcesU2\Build\Tester_U2_Office_Isole.ps1 -RacineSources C:\SourcesU2 -Sortie C:\Recettes\QualificationNeuve
```

Les deux chemins sont des exemples. La sortie doit être inexistante et son
parent doit déjà exister sur un disque local fixe. Son chemin complet est limité
à 90 caractères : la suite U2 ajoute 164 caractères pour un fichier temporaire
identifié par GUID et SHA-256, et FileSystemObject VBA reste soumis à MAX_PATH.
Choisir par exemple un nom de sortie court `QI-12345678`. Le lanceur refuse les liens
symboliques, les dossiers de démarrage Office non vides, un Normal contenant
un projet VBA et les réglages de lancement personnalisés détectés. Les
chemins machine Office 16 et les deux vues de registre sont contrôlés.

La construction utilise les trois modèles publics et le manifeste du dépôt.
Avant l'exécution des suites, seules les copies construites sont instrumentées :

- appels NAS, appels OpenAI et lecture de clés remplacés par des erreurs ;
- chemins APPDATA, LOCALAPPDATA et temporaires remplacés par des chemins
  explicites sous la sortie, sans dépendre de l'héritage d'environnement COM ;
- racine de données par défaut locale et fictive ;
- procédures automatiques neutralisées et événements Excel désactivés ;
- appels d'impression bloqués.

Les tests métier ciblés restent présents. Les modèles historiques personnels,
les courriers réels et les secrets ne sont pas utilisés. La copie instrumentée
ne constitue pas un fichier à installer dans Word ou Excel.

Le lanceur conserve puis restaure Normal et les paramètres de sécurité VBA.
Il ne termine aucun processus de force. En cas de fermeture incomplète,
conserver la sortie et les journaux `*-a-restaurer.json` pour examiner l'état
avant une reprise. Aucun résultat ne vaut validation si le rapport final
`qualification-office-isolee.json` indique ECHEC.

Les résultats détaillés figurent dans `resultats-suites-office.json`, les
modules instrumentés dans `instrumentation-Word.json` et
`instrumentation-Excel.json`. Ces fichiers de suivi restent locaux. Ils ne
doivent pas être publiés avec une copie de Normal.

Cette qualification ne valide pas les accès NAS, la transmission connectée,
la restauration Synology, l'impression papier, Dragon ou le matériel ECG.
Le test `Tests/test_recette_isolee.ps1` vérifie séparément l'instrumentation
sans démarrer Office et sans prétendre effectuer une compilation VBA réelle.
