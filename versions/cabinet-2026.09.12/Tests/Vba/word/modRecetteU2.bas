Attribute VB_Name = "modRecetteU2"
Option Explicit
Private Const NOMBRE_ATTENDU_U2 As Long = 36
Private mNombre As Long

Private Sub ExigerU2(ByVal condition As Boolean, ByVal nom As String)
    If Not condition Then Err.Raise vbObjectError + 1192, "Recette U2", nom
    mNombre = mNombre + 1
End Sub

Private Function DestinataireIARefuse(ByVal ident As String, ByVal cor As Object) As Boolean
    Dim numero As Long
    On Error Resume Next
    modOpenAI_v22_corrige.VerifierDestinataireAvantIA ident, cor
    numero = Err.Number: Err.Clear
    On Error GoTo 0
    DestinataireIARefuse = (numero = vbObjectError + 432)
End Function

Private Sub TesterDestinataireAvantIA()
    Dim cor As Object
    Set cor = modServiceNas.Parametres()
    cor("ID") = "DEST_FICTIF": cor("Actif") = "1": cor("AValider") = "0"
    modOpenAI_v22_corrige.VerifierDestinataireAvantIA " DEST_FICTIF ", cor
    ExigerU2 True, "destinataire explicite actif et valide accepte"
    ExigerU2 DestinataireIARefuse(" ", cor), "IA refuse un ID destinataire absent"
    ExigerU2 DestinataireIARefuse("DEST_FICTIF", Nothing), "IA refuse un destinataire introuvable"
    ExigerU2 DestinataireIARefuse("AUTRE_FICTIF", cor), "IA refuse un destinataire different de la selection"
    cor("Actif") = "0"
    ExigerU2 DestinataireIARefuse("DEST_FICTIF", cor), "IA refuse un destinataire desactive"
    cor("Actif") = "1": cor("AValider") = "1"
    ExigerU2 DestinataireIARefuse("DEST_FICTIF", cor), "IA refuse un destinataire restant a valider"
    cor("AValider") = "0": cor.Remove "Actif"
    ExigerU2 DestinataireIARefuse("DEST_FICTIF", cor), "IA refuse un statut actif absent"
    cor("Actif") = "1": cor.Remove "AValider"
    ExigerU2 DestinataireIARefuse("DEST_FICTIF", cor), "IA refuse un statut de validation absent"
    cor("AValider") = ""
    ExigerU2 DestinataireIARefuse("DEST_FICTIF", cor), "IA refuse un statut de validation vide non normalise par le serveur"
    cor("AValider") = "inconnu"
    ExigerU2 DestinataireIARefuse("DEST_FICTIF", cor), "IA refuse un statut de validation inconnu"
    cor("AValider") = "0": cor("Actif") = ""
    ExigerU2 DestinataireIARefuse("DEST_FICTIF", cor), "IA refuse un statut actif vide non normalise par le serveur"
    cor("Actif") = "inconnu"
    ExigerU2 DestinataireIARefuse("DEST_FICTIF", cor), "IA refuse un statut actif inconnu"
End Sub

Public Function ExecuterU2(ByVal sortie As String) As String
    Dim p As Object, q As Object, fso As Object, ts As Object, cree As Boolean
    Dim dossier As String, cle As String, ancien As String, chemin As String, id As String, id2 As String
    Dim reprise As Boolean, numero As Long, description As String
    On Error GoTo Echec
    mNombre = 0
    Set p = modServiceNas.Parametres(): p("z") = "FICTIF": p("a") = "1"
    Set q = modServiceNas.Parametres(): q("a") = "1": q("z") = "FICTIF"
    cle = modCommandesLocales.CleCommandeLocale("https://fictif", "compte-a", "claim", p)
    ExigerU2 cle = modCommandesLocales.CleCommandeLocale("https://fictif", "compte-a", "claim", q), "cle canonique"
    ExigerU2 cle <> modCommandesLocales.CleCommandeLocale("https://fictif", "compte-b", "claim", p), "compte distinct"
    ExigerU2 cle <> modCommandesLocales.CleCommandeLocale("https://autre", "compte-a", "claim", p), "serveur distinct"
    ExigerU2 cle <> modCommandesLocales.CleCommandeLocale("https://fictif", "compte-a", "release", p), "operation distincte"
    q("a") = "2"
    ExigerU2 cle <> modCommandesLocales.CleCommandeLocale("https://fictif", "compte-a", "claim", q), "contenu distinct"
    Set fso = CreateObject("Scripting.FileSystemObject")
    dossier = sortie & "\commandes-U2-" & modFichiers.IdUnique()
    fso.CreateFolder dossier: cree = True
    id = modCommandesLocales.ObtenirCommandeLocale(cle, chemin, reprise, modServiceNas.SHA256("ancien jeton fictif"), dossier)
    id2 = modCommandesLocales.ObtenirCommandeLocale(cle, chemin, reprise, modServiceNas.SHA256("nouveau jeton fictif"), dossier)
    ExigerU2 reprise And id = id2, "rotation sans nouvel identifiant"
    cle = modServiceNas.SHA256("migration fictive"): ancien = modServiceNas.SHA256("ancienne cle fictive")
    Set ts = fso.CreateTextFile(dossier & "\" & ancien & ".pending", False, False): ts.Write String$(32, "a"): ts.Close
    id = modCommandesLocales.ObtenirCommandeLocale(cle, chemin, reprise, ancien, dossier)
    ExigerU2 reprise And id = String$(32, "a") And Not fso.FileExists(dossier & "\" & ancien & ".pending"), "migration U1 conserve identifiant"
    Set ts = fso.CreateTextFile(dossier & "\" & ancien & ".pending", False, False): ts.Write String$(32, "b"): ts.Close
    On Error Resume Next
    id2 = modCommandesLocales.ObtenirCommandeLocale(cle, chemin, reprise, ancien, dossier)
    numero = Err.Number: Err.Clear
    On Error GoTo Echec
    ExigerU2 numero = vbObjectError + 1108, "identifiants contradictoires refuses"
    ExigerU2 fso.FileExists(dossier & "\" & ancien & ".pending") And fso.FileExists(chemin), "fichiers contradictoires conserves"
    ExigerU2 modCycleCourrier.DocumentSource() Is Nothing And Not modCycleCourrier.CycleEnCours(), "contexte de correction libere"
    ExigerU2 modProdRapide.PR_DoitForcerCorrespondantACompleterCCN("Je l'adresse au CCN.", "Bilan de rythmologie"), "CCN apostrophe droite"
    ExigerU2 modProdRapide.PR_DoitForcerCorrespondantACompleterCCN("Je l" & ChrW(8217) & "adresse au CCN.", "Bilan de rythmologie"), "CCN apostrophe courbe"
    ExigerU2 modProdRapide.PR_DoitForcerCorrespondantACompleterCCN("Je l" & ChrW(8216) & "oriente au CCN.", "Bilan de rythmologie"), "CCN apostrophe gauche"
    ExigerU2 Not modProdRapide.PR_DoitForcerCorrespondantACompleterCCN("Antecedent d" & ChrW(8217) & "ablation au CCN.", "Bilan de rythmologie"), "CCN historique sans nouvelle orientation"
    ExigerU2 Not modProdRapide.PR_DoitForcerCorrespondantACompleterCCN("Je l'adresse au docteur FICTIF au CCN.", "Bilan de rythmologie"), "CCN correspondant nomme conserve"
    chemin = dossier & "\empreinte.bin"
    Set ts = fso.CreateTextFile(chemin, False, False): ts.Write "abc": ts.Close
    ExigerU2 modDonneesTransport.EmpreinteFichierSHA256(chemin) = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "empreinte du contenu binaire"
    Set ts = fso.CreateTextFile(chemin, True, False): ts.Close
    ExigerU2 modDonneesTransport.EmpreinteFichierSHA256(chemin) = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855", "empreinte du fichier vide"
    TesterDestinataireAvantIA
    TesterDestinataireRelecture
    If mNombre <> NOMBRE_ATTENDU_U2 Then Err.Raise vbObjectError + 1192, "Recette U2", "Nombre de controles inattendu."
    ExecuterU2 = "{""reussis"":" & CStr(mNombre) & ",""attendus"":" & CStr(NOMBRE_ATTENDU_U2) & ",""echec"":false}"
Sortie:
    On Error Resume Next
    If cree Then fso.DeleteFolder dossier, True
    On Error GoTo 0
    Exit Function
Echec:
    description = Err.Description
    ExecuterU2 = "{""reussis"":" & CStr(mNombre) & ",""attendus"":" & CStr(NOMBRE_ATTENDU_U2) & ",""echec"":true,""description"":" & modServiceNas.JsonValeur(description) & "}"
    Resume Sortie
End Function

Private Sub TesterDestinataireRelecture()
    Dim doc As Document, cor As Object, rng As Range
    Dim numero As Long, description As String
    On Error GoTo Echec
    Set doc = Documents.Add
    Set cor = modServiceNas.Parametres()
    cor("ID") = "DEST-FICTIF": cor("BlocDestinataire") = "Docteur FICTIF"
    doc.Content.Text = CStr(cor("BlocDestinataire"))
    Set rng = doc.Range(0, Len(CStr(cor("BlocDestinataire"))))
    doc.Bookmarks.Add "DESTINATAIRE", rng
    modControleCourrier.InvaliderRelecture doc
    ExigerU2 modIntegrationUnifie.VariableDoc(doc, "RelectureValidee") = "0" And modIntegrationUnifie.VariableDoc(doc, "RelectureEnAttente") = "0", "aucune relecture presumee"
    ExigerU2 Not modControleCourrier.DestinataireRelectureConforme(doc, cor), "relecture sans instantane refusee"
    modIntegrationUnifie.FixerVariable doc, "CorrespondantID", CStr(cor("ID"))
    modIntegrationUnifie.FixerVariable doc, "RelectureDestinataireID", CStr(cor("ID"))
    modIntegrationUnifie.FixerVariable doc, "RelectureAdresseSource", modServiceNas.SHA256(CStr(cor("BlocDestinataire")))
    modIntegrationUnifie.FixerVariable doc, "RelectureAdresseDocument", modServiceNas.SHA256(rng.Text)
    ExigerU2 modControleCourrier.DestinataireRelectureConforme(doc, cor), "destinataire inchange accepte"
    modIntegrationUnifie.FixerVariable doc, "CorrespondantID", "AUTRE"
    ExigerU2 Not modControleCourrier.DestinataireRelectureConforme(doc, cor), "changement identifiant refuse"
    modIntegrationUnifie.FixerVariable doc, "CorrespondantID", CStr(cor("ID"))
    cor("BlocDestinataire") = "Autre adresse"
    ExigerU2 Not modControleCourrier.DestinataireRelectureConforme(doc, cor), "changement adresse serveur refuse"
    cor("BlocDestinataire") = "Docteur FICTIF"
    modCourrier.RemplirSignet doc, "DESTINATAIRE", "Autre destinataire"
    ExigerU2 Not modControleCourrier.DestinataireRelectureConforme(doc, cor), "modification adresse document refusee"
    doc.Bookmarks("DESTINATAIRE").Delete
    ExigerU2 Not modControleCourrier.DestinataireRelectureConforme(doc, cor), "signet destinataire absent refuse"
Sortie:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    If numero <> 0 Then Err.Raise numero, "TesterDestinataireRelecture", description
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    Resume Sortie
End Sub

Public Function ExecuterFileAnnexes(ByVal sortie As String) As String
    Dim f As ufArriveesMedecin, items As New Collection, d As Object, i As Long
    Dim principal As Document, annexe As Document, r As Range, cor As Object
    Dim n As Long, debut As Long, description As String, ancien As String
    On Error GoTo Echec
    Set f = New ufArriveesMedecin
    For i = 1 To 2
        Set d = modServiceNas.Parametres()
        d("ConsultationID") = "FICTIF-S" & CStr(i)
        d("Nom") = "FICTIF" & CStr(i): d("Prenom") = "ESSAI"
        d("DDN") = "01/01/1980": d("HeureArrivee") = "09:00": d("HeureRdv") = "09:15"
        items.Add d
    Next i
    f.Charger items, "09:01:00"
    VerifierNouvellesFonctions f.NombreAffiche() = 2, "deux arrivees visibles", n
    VerifierNouvellesFonctions f.IdentifiantSelection() = "", "aucun patient selectionne automatiquement", n
    f.Controls("lstPatients").ListIndex = 1
    f.Charger items, "09:01:15"
    VerifierNouvellesFonctions f.IdentifiantSelection() = "FICTIF-S2", "selection stable apres actualisation", n
    f.Controls("txtFiltre").Text = "fictif1"
    VerifierNouvellesFonctions f.NombreAffiche() = 1, "filtre sans distinction de casse", n
    VerifierNouvellesFonctions f.IdentifiantSelection() = "", "filtre ne choisit pas un autre patient", n
    f.Controls("txtFiltre").Text = ""
    f.Indisponible "FICTIF : NAS indisponible"
    VerifierNouvellesFonctions Not f.Controls("lstPatients").Enabled, "liste perimee non cliquable", n
    f.Charger items, "09:01:30"
    VerifierNouvellesFonctions f.Controls("lstPatients").Enabled, "actualisation retablit la liste", n
    f.Show vbModeless
    Set principal = Documents.Add
    principal.Content.Text = "Docteur PRINCIPAL" & vbCr & "Corps fictif du courrier." & vbCr
    principal.Activate
    VerifierNouvellesFonctions f.Visible, "fenetre persiste avec un document Word actif", n
    f.Hide: Unload f: Set f = Nothing
    principal.Bookmarks.Add "DESTINATAIRE", principal.Range(0, Len("Docteur PRINCIPAL"))
    ancien = principal.Bookmarks("DESTINATAIRE").Range.Text
    For Each d In items
        ' Aucun appel reseau : uniquement des donnees fictives deja construites.
    Next d
    modIndexAnnexes.FixerMeta principal, "PatientID", "FICTIF-P1"
    modIndexAnnexes.FixerMeta principal, "ConsultationID", "FICTIF-S1"
    modIndexAnnexes.FixerMeta principal, "Patient_Nom", "FICTIF"
    modIndexAnnexes.FixerMeta principal, "Patient_Prenom", "ESSAI"
    modIndexAnnexes.FixerMeta principal, "Patient_DDN", "01/01/1980"
    modIndexAnnexes.FixerMeta principal, "Patient_Sexe", "M"
    Set annexe = Documents.Add
    annexe.Content.Text = "Docteur ANNEXE" & vbCr & "Corps fictif de l annexe." & vbCr
    annexe.Bookmarks.Add "U2_DEST_SOURCE", annexe.Range(0, Len("Docteur ANNEXE"))
    debut = principal.Content.End - 1
    principal.Range(debut, debut).FormattedText = annexe.Range(0, annexe.Content.End - 1).FormattedText
    modIndexAnnexes.IndexerAnnexe principal, annexe, debut, 1, "FICTIF-DEST-A"
    VerifierNouvellesFonctions principal.Bookmarks.Exists("U2ANN_DEST_001"), "repere annexe conserve", n
    VerifierNouvellesFonctions modIndexAnnexes.ListerAnnexes(principal).Count = 1, "une annexe indexee", n
    VerifierNouvellesFonctions modIndexAnnexes.LireMeta(principal, "AnnexeDestinataire_001") = "FICTIF-DEST-A", "identifiant du destinataire conserve", n
    principal.SaveAs2 sortie & "\annexe-fictive-avant.docx", 12
    Set cor = modServiceNas.Parametres()
    cor("ID") = "FICTIF-DEST-B": cor("Actif") = "1": cor("AValider") = "0"
    cor("BlocDestinataire") = "Docteur REMPLACEMENT" & vbLf & "ADRESSE FICTIVE"
    modIndexAnnexes.RemplacerDestinataireAnnexe principal, "001", cor
    VerifierNouvellesFonctions principal.Bookmarks("DESTINATAIRE").Range.Text = ancien, "destinataire principal intact", n
    VerifierNouvellesFonctions principal.Bookmarks("U2ANN_DEST_001").Range.Text = Replace(CStr(cor("BlocDestinataire")), vbLf, Chr$(11)), "adresse annexe remplacee", n
    VerifierNouvellesFonctions modIndexAnnexes.LireMeta(principal, "AnnexeDestinataire_001") = "FICTIF-DEST-B", "nouvel identifiant annexe", n
    VerifierNouvellesFonctions InStr(principal.Content.Text, "Corps fictif de l annexe.") > 0, "corps de l annexe conserve", n
    principal.SaveAs2 sortie & "\annexe-fictive-apres.docx", 12
    principal.Close 0: Set principal = Documents.Open(sortie & "\annexe-fictive-apres.docx", False, True, False)
    VerifierNouvellesFonctions principal.Bookmarks.Exists("U2ANN_DEST_001"), "repere persiste apres fermeture", n
    VerifierNouvellesFonctions modIndexAnnexes.LireMeta(principal, "AnnexeDestinataire_001") = "FICTIF-DEST-B", "identifiant persiste apres fermeture", n
    ExecuterFileAnnexes = "{""reussis"":" & CStr(n) & ",""echec"":false}"
Sortie:
    On Error Resume Next
    If Not f Is Nothing Then Unload f
    If Not principal Is Nothing Then principal.Close 0
    If Not annexe Is Nothing Then annexe.Close 0
    On Error GoTo 0
    Exit Function
Echec:
    description = Err.Description
    ExecuterFileAnnexes = "{""reussis"":" & CStr(n) & ",""echec"":true,""description"":" & modServiceNas.JsonValeur(description) & "}"
    Resume Sortie
End Function

Private Sub VerifierNouvellesFonctions(ByVal condition As Boolean, ByVal message As String, ByRef n As Long)
    If Not condition Then Err.Raise vbObjectError + 1290, "File et annexes", message
    n = n + 1
End Sub
