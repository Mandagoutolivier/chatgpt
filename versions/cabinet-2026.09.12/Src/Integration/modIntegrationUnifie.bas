Attribute VB_Name = "modIntegrationUnifie"
Option Explicit

Public Function VariableDoc(ByVal doc As Document, ByVal nom As String) As String
    On Error Resume Next
    VariableDoc = doc.Variables(nom).Value
End Function

Public Sub FixerVariable(ByVal doc As Document, ByVal nom As String, ByVal valeur As String)
    If Len(valeur) = 0 Then valeur = " "
    Dim v As Variable
    For Each v In doc.Variables
        If StrComp(v.Name, nom, vbTextCompare) = 0 Then
            v.Value = valeur
            Exit Sub
        End If
    Next v
    doc.Variables.Add Name:=nom, Value:=valeur
End Sub

Public Function PatientVerifie(ByVal doc As Document) As Object
    Dim pat As Object, id As String, k As Variant
    id = Trim$(VariableDoc(doc, "PatientID"))
    If Len(id) = 0 Then Err.Raise vbObjectError + 950, "modIntegrationUnifie", "Ce courrier n est pas rattache a un patient."
    Set pat = modBase.PatientParID(id, True)
    If pat Is Nothing Then Err.Raise vbObjectError + 951, "modIntegrationUnifie", "Patient introuvable sur le NAS."
    For Each k In Array("Nom", "Prenom", "DDN", "Sexe")
        If Len(Trim$(CStr(pat(k)))) = 0 Then Err.Raise vbObjectError + 952, "modIntegrationUnifie", "Fiche patient incomplete : " & k
        If Len(Trim$(VariableDoc(doc, "Patient_" & k))) > 0 Then
            If StrComp(CStr(pat(k)), VariableDoc(doc, "Patient_" & k), vbBinaryCompare) <> 0 Then
                Err.Raise vbObjectError + 953, "modIntegrationUnifie", "L identite en base a change depuis l ouverture. Verifiez le patient avant de reprendre ce courrier."
            End If
        Else
            FixerVariable doc, "Patient_" & k, CStr(pat(k))
        End If
    Next k
    If Not modTexte.DateFrValide(CStr(pat("DDN"))) Then Err.Raise vbObjectError + 954, "modIntegrationUnifie", "Date de naissance invalide."
    If modTexte.DateFr(CStr(pat("DDN"))) > Date Then Err.Raise vbObjectError + 955, "modIntegrationUnifie", "Date de naissance future."
    If Len(modTexte.SexeNormalise(CStr(pat("Sexe")))) = 0 Then Err.Raise vbObjectError + 956, "modIntegrationUnifie", "Sexe patient indetermine."
    If Len(Trim$(VariableDoc(doc, "ConsultationID"))) = 0 Then FixerVariable doc, "ConsultationID", modFichiers.IdUnique()
    If Len(Trim$(VariableDoc(doc, "DateActe"))) = 0 Then FixerVariable doc, "DateActe", Format$(Date, "dd/mm/yyyy")
    Set PatientVerifie = pat
End Function

Public Function InitialiserPatientProd(ByVal doc As Document) As Boolean
    Dim pat As Object
    Set pat = PatientVerifie(doc)
    gPatient.nom = UCase$(CStr(pat("Nom")))
    gPatient.prenom = CStr(pat("Prenom"))
    gPatient.civilite = modTexte.Civilite(CStr(pat("Sexe")))
    gPatient.Age = modCourrier.CalculerAge(CStr(pat("DDN")))
    gPatient.NomComplet = gPatient.civilite & " " & gPatient.nom & " " & gPatient.prenom
    gPatient.marqueur = MARQUEUR_PATIENT
    InitialiserPatientProd = True
End Function

Public Sub SauvegarderBrouillon(ByVal doc As Document)
    Dim pat As Object, chemin As String
    Set pat = PatientVerifie(doc)
    chemin = modPatient.DossierPatient(pat) & "\brouillon_" & modFichiers.IdUnique() & ".docx"
    doc.SaveAs2 FileName:=chemin, FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    If Not modFichiers.FichierExiste(chemin) Then Err.Raise vbObjectError + 957, "modIntegrationUnifie", "Brouillon non enregistre sur le NAS."
    Dim reservation As String, attente As Object
    reservation = Trim$(VariableDoc(doc, "ReservationNas"))
    If Len(reservation) > 0 Then
            Set attente = CreateObject("Scripting.Dictionary")
            attente("ReservationNas") = reservation
            modAttenteLocale.EnregistrerBrouillon attente, chemin
    End If
End Sub

' Reutilise la validation Cabinet : archive docx + PDF + identifiants stables.
Public Sub TransmettreSecretariat(ByVal doc As Document, ByVal cheminDocx As String)
    Dim typeCourrier As String
    typeCourrier = modValidation.ValiderDocument(doc, True)
End Sub
' Retrouver les consultations interrompues, sans les recreer ni les refacturer.
Public Sub Unifie_ReprendreBrouillon()
    On Error GoTo Echec
    Dim d As Object, items As New Collection, fichier As Variant, f As ufListe, choisi As Object
    Dim resultat As Object
    Set resultat = modServiceNas.Appeler("reprises", modServiceNas.Parametres())
    For Each d In resultat("items")
        d("Patient") = d("Nom") & " " & d("Prenom")
        d("Etat") = "Brouillon disponible"
        If Not d.Exists("CheminBrouillon") Then d("Etat") = "Reservation interrompue"
        items.Add d
    Next d
    If items.Count = 0 Then MsgBox "Aucun brouillon en cours.", vbInformation: Exit Sub
    Set f = New ufListe
    f.Configurer "Reprendre une consultation", items, Array("Patient", "DDN", "Etat"), "180 pt;70 pt;130 pt"
    f.Show vbModal
    If Not f.Annule Then Set choisi = f.Resultat
    Unload f
    If choisi Is Nothing Then Exit Sub
    If Not choisi.Exists("CheminBrouillon") Then
        MsgBox "L ouverture a ete interrompue avant l enregistrement du lien vers le brouillon. Faites verifier cette reservation et le dossier patient sur le NAS avant de la liberer.", vbExclamation, "Consultation a recuperer"
        Exit Sub
    End If
    Documents.Open FileName:=CStr(choisi("CheminBrouillon")), AddToRecentFiles:=False
    Exit Sub
Echec:
    MsgBox "Reprise impossible : " & Err.Description, vbExclamation, "Cabinet"
End Sub
Public Function LocaliserCorpsUnifie(ByVal doc As Document, ByRef corps As Range, ByRef premier As Range) As Boolean
    Dim p As Paragraph, debut As Long, fin As Long, limite As Long, texte As String, appelTrouve As Boolean
    limite = doc.Content.End - 1
    If doc.Bookmarks.Exists("PR_DEBUT_DEMANDES") Then limite = doc.Bookmarks("PR_DEBUT_DEMANDES").Range.Start
    ' La formule d appel est un paragraphe distinct ; le corps commence juste apres.
    For Each p In doc.Paragraphs
        If p.Range.Start >= limite Then Exit For
        texte = NettoyerTexteParagraphe(p.Range.Text)
        If Not appelTrouve Then
            If EstFormuleAppel(texte) Then
                debut = p.Range.End
                appelTrouve = True
            End If
        ElseIf EstSignature(texte) Then
            fin = p.Range.Start
            Exit For
        End If
    Next p
    If Not appelTrouve Or fin <= debut Then Exit Function
    ' Seule une formule reconnue en fin de corps est exclue. On ne retire
    ' jamais arbitrairement le dernier paragraphe medical avant la signature.
    Dim zone As Range, i As Long
    Set zone = doc.Range(debut, fin)
    For i = zone.Paragraphs.Count To 1 Step -1
        texte = NettoyerTexteParagraphe(zone.Paragraphs(i).Range.Text)
        If Len(texte) = 0 Then
            fin = zone.Paragraphs(i).Range.Start
        ElseIf EstPolitesseFinale(texte) Then
            fin = zone.Paragraphs(i).Range.Start
            Exit For
        Else
            Exit For
        End If
    Next i
    If fin <= debut Then Exit Function
    Set corps = doc.Range(debut, fin)
    Set premier = corps.Paragraphs(1).Range
    doc.Bookmarks.Add "CORPS", corps
    LocaliserCorpsUnifie = True
End Function

Private Function EstPolitesseFinale(ByVal texte As String) As Boolean
    Dim t As String
    t = modTexte.Plier(Trim$(texte))
    Select Case t
        Case "bien confraternellement.", "bien confraternellement,", "confraternellement.", _
             "bien cordialement.", "bien cordialement,", "cordialement.", "cordialement,", _
             "bien amicalement.", "amities.", "merci de votre confiance.", "merci de ta confiance."
            EstPolitesseFinale = True
    End Select
End Function
