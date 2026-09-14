Attribute VB_Name = "modRecetteU0"
Option Explicit

' Recette SYNTHETIQUE. Aucun patient NAS, aucune reservation, aucun appel API.
' Les GDT restent dans TEMP, jamais dans le dossier surveille par l ECG.
Public Sub Unifie_U0_EssaisLocaux()
    On Error GoTo Echec
    Dim dossier As String, pat As Object, cor As Object, auteur As Object
    Dim ctx As Object, requete As String, corps As String, resultat As Object, k As Variant
    dossier = Environ$("TEMP") & "\Cabinet-U0-" & modFichiers.IdUnique()
    MkDir dossier
    Set pat = CreateObject("Scripting.Dictionary")
    For Each k In Array("ID", "Nom", "NomNaissance", "Prenom", "DDN", "Sexe", "NIR", "Adresse1", "Adresse2", "Tel", "Mobile", "Email")
        pat(CStr(k)) = ""
    Next k
    pat("ID") = "P123456789A00000000000000000000001"
    pat("Nom") = "TEST-" & ChrW$(201) & "CG-ALPHA": pat("Prenom") = "Ana" & ChrW$(239) & "s"
    pat("DDN") = "25/01/1946": pat("Sexe") = "F"
    modFichiers.EcrireTexteAnsi dossier & "\ECG_A.gdt", modGdt.ConstruireGdt(pat)
    pat("ID") = "P123456789B00000000000000000000002"
    pat("Nom") = "TEST-ECG-BETA": pat("Prenom") = "Beno" & ChrW$(238) & "t"
    pat("DDN") = "17/02/1952": pat("Sexe") = "M"
    modFichiers.EcrireTexteAnsi dossier & "\ECG_B.gdt", modGdt.ConstruireGdt(pat)

    pat("Nom") = "PAGE": pat("Prenom") = "LI": pat("NomNaissance") = "FICTIVE"
    pat("Adresse1") = "12 voie FICTIVE": pat("Tel") = "0100000000"
    pat("Email") = "patient-fictif@example.invalid"
    Set cor = CreateObject("Scripting.Dictionary")
    cor("Nom") = "DESTINATAIREFICTIF": cor("Prenom") = "ALICEFICTIVE"
    Set auteur = CreateObject("Scripting.Dictionary")
    auteur("Nom") = "AUTEURFICTIF": auteur("Prenom") = "MARCELFICTIF"
    Set ctx = modAnonymise.Construire(pat, Nothing)
    modAnonymise.AjouterCorrespondant ctx, cor, "DEST"
    modAnonymise.AjouterCorrespondant ctx, auteur, "AUTEUR"
    corps = "Je revois PAGE LI, nee FICTIVE le 17/02/1952, 12 voie FICTIVE, " & _
        "0100000000, patient-fictif@example.invalid. Adresse par ALICEFICTIVE DESTINATAIREFICTIF. " & _
        "Suivi par MARCELFICTIF AUTEURFICTIF. Aucun dopage. Absence de douleur. Dose 5 mg."
    requete = modOpenAI_v22_corrige.PreparerRequeteSortante(ConstruirePromptReecritureMedicale(corps), ctx)
    Set resultat = modJson.JsonParse(requete)
    For Each k In Array("{{PAT_NOM}}", "{{PAT_PRENOM}}", "{{PAT_DDN}}", "{{PAT_ADRESSE}}", "{{PAT_TEL}}", "{{PAT_EMAIL}}", "{{DEST_NOM}}", "{{AUTEUR_NOM}}")
        If InStr(1, CStr(resultat("input")), CStr(k), vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 1207, , "Balise attendue absente de la requete fictive."
    Next k
    If Len(modAnonymise.ScanResiduel(CStr(resultat("input")), ctx)) > 0 Then Err.Raise vbObjectError + 1201, , "Residu dans la requete fictive."
    If InStr(1, CStr(resultat("input")), "dopage", vbTextCompare) = 0 Then Err.Raise vbObjectError + 1202, , "Nom court : faux remplacement dans dopage."
    If InStr(1, modAnonymise.Reinjecter(CStr(resultat("input")), ctx), corps, vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 1203, , "Reinjection fictive non reversible."
    If InStr(1, requete, "Olivier Mandagout", vbTextCompare) > 0 Then Err.Raise vbObjectError + 1204, , "Identite auteur dans une consigne."
    modFichiers.EcrireTexteUTF8 dossier & "\requete_principale_fictive.json", requete
    requete = modOpenAI_v22_corrige.PreparerRequeteSortante(ConstruirePromptDemandeExamenSeule(corps), ctx)
    modFichiers.EcrireTexteUTF8 dossier & "\requete_annexes_fictive.json", requete
    If Len(modAnonymise.ScanResiduel("Appeler +33 6 12 34 56 78 ou inconnu@example.invalid", ctx)) = 0 Then Err.Raise vbObjectError + 1205, , "Residu inconnu non detecte."
    If Len(modAnonymise.VerifierBalisesRetour("Texte sans identite", ctx, "{{PAT_NOM}}")) = 0 Then Err.Raise vbObjectError + 1206, , "Perte de balise non detectee."
    MsgBox "Essais locaux termines, aucun envoi API ou ECG." & vbCrLf & dossier & vbCrLf & _
        "Inspecter les deux JSON. Tester ensuite A, B, puis A dans Resting12Lead selon U0_RECETTE.md.", vbInformation, "Cabinet U0"
    Exit Sub
Echec:
    MsgBox "Recette U0 interrompue : " & Err.Description, vbExclamation, "Cabinet U0"
End Sub
