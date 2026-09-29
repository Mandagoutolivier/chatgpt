Attribute VB_Name = "ufPatientMedecin"
Option Explicit
Private mDocument As Document
Private mPatient As Object
Private mMedecinID As String

Private Function Champs() As Variant
    Champs = Array("Nom", "Prenom", "NomNaissance", "DDN", "Sexe", "NIR", "Adresse1", "Adresse2", "CP", "Ville", "Tel", "Mobile", "Email", "Mutuelle", "ALD", "Notes")
End Function

Public Sub Charger(ByVal doc As Document, ByVal pat As Object)
    Set mDocument = doc
    Set mPatient = pat
    txtNotes.MultiLine = True
    txtNotes.EnterKeyBehavior = True
    Remplir
End Sub

Private Sub Remplir()
    Dim k As Variant, cor As Object
    Me.Caption = "Fiche patient - " & CStr(mPatient("Nom")) & " " & CStr(mPatient("Prenom"))
    lblPatient.Caption = "Dossier " & CStr(mPatient("ID"))
    For Each k In Champs()
        Me.Controls("txt" & CStr(k)).Text = CStr(mPatient(k))
    Next k
    mMedecinID = CStr(mPatient("MedTraitantID"))
    lblMedecin.Caption = "Medecin traitant : aucun"
    If Len(mMedecinID) > 0 Then
        Set cor = modBase.CorrespondantParID(mMedecinID)
        lblMedecin.Caption = "Medecin traitant : " & CStr(cor("Nom")) & " " & CStr(cor("Prenom"))
    End If
    AfficherEtatIdentite
End Sub

Private Sub AfficherEtatIdentite()
    btnConfirmerIdentite.Visible = modFichePatient.IdentiteCourrierDifferente(mDocument, mPatient) And Not mDocument.ReadOnly
    lblEtat.Caption = ""
    If modFichePatient.IdentiteCourrierDifferente(mDocument, mPatient) Then
        lblEtat.Caption = "L identite du courrier differe. Fermez cette fiche, corrigez l identite et l age dans le courrier et ses annexes, puis rouvrez la fiche pour confirmer leur verification."
    End If
End Sub

Private Function DonneesSaisies() As Object
    Dim d As Object, k As Variant
    Set d = modServiceNas.Parametres()
    d("ID") = CStr(mPatient("ID"))
    d("_revision") = CStr(mPatient("_revision"))
    For Each k In Champs()
        d(CStr(k)) = Trim$(Me.Controls("txt" & CStr(k)).Text)
    Next k
    d("Nom") = UCase$(CStr(d("Nom")))
    d("NomNaissance") = UCase$(CStr(d("NomNaissance")))
    d("Sexe") = modTexte.SexeNormalise(CStr(d("Sexe")))
    d("ALD") = UCase$(CStr(d("ALD")))
    d("MedTraitantID") = mMedecinID
    Set DonneesSaisies = d
End Function

Private Function SaisieModifiee() As Boolean
    Dim d As Object, k As Variant
    Set d = DonneesSaisies()
    For Each k In d.Keys
        If CStr(d(k)) <> CStr(mPatient(k)) Then SaisieModifiee = True: Exit Function
    Next k
End Function

Private Sub btnEnregistrer_Click()
    On Error GoTo Echec
    Dim d As Object, p As Object
    Set d = DonneesSaisies()
    If Len(d("Nom")) = 0 Or Len(d("Prenom")) = 0 Then Err.Raise vbObjectError + 1291, , "Nom et prenom obligatoires."
    If Not modTexte.DateFrValide(CStr(d("DDN"))) Then Err.Raise vbObjectError + 1291, , "Naissance : utilisez jj/mm/aaaa."
    If modTexte.DateFr(CStr(d("DDN"))) > Date Then Err.Raise vbObjectError + 1291, , "La naissance ne peut pas etre future."
    If Len(d("Sexe")) = 0 Then Err.Raise vbObjectError + 1291, , "Sexe : indiquez M ou F."
    If d("ALD") <> "" And d("ALD") <> "O" And d("ALD") <> "N" Then Err.Raise vbObjectError + 1291, , "ALD : indiquez O, N ou laissez vide."
    If Not SaisieModifiee() Then Me.Hide: Exit Sub
    d("DDN") = Format$(modTexte.DateFr(CStr(d("DDN"))), "dd/mm/yyyy")
    d("DateModif") = Format$(Date, "dd/mm/yyyy")
    Set p = modServiceNas.Parametres()
    p("consultation_id") = Trim$(modIntegrationUnifie.VariableDoc(mDocument, "ConsultationID"))
    Set p("data") = d
    btnEnregistrer.Enabled = False
    Set mPatient = modServiceNas.Appeler("patient.update", p)
    btnEnregistrer.Enabled = True
    Remplir
    If modFichePatient.IdentiteCourrierDifferente(mDocument, mPatient) Then
        If Not mDocument.ReadOnly Then
            modControleCourrier.InvaliderRelecture mDocument
            mDocument.Save
        End If
        lblEtat.Caption = "Fiche enregistree. " & lblEtat.Caption
    Else
        Me.Hide
    End If
    Exit Sub
Echec:
    btnEnregistrer.Enabled = True
    MsgBox "Enregistrement : " & Err.Description, vbExclamation, "Fiche patient"
End Sub

Private Sub btnMedecin_Click()
    On Error GoTo Echec
    Dim cor As Object
    Set cor = modPatient.ChoisirCorrespondant(mMedecinID, "Medecin traitant")
    If cor Is Nothing Then Exit Sub
    mMedecinID = CStr(cor("ID"))
    lblMedecin.Caption = "Medecin traitant : " & CStr(cor("Nom")) & " " & CStr(cor("Prenom"))
    Exit Sub
Echec:
    MsgBox Err.Description, vbExclamation, "Fiche patient"
End Sub

Private Sub btnConfirmerIdentite_Click()
    On Error GoTo Echec
    If SaisieModifiee() Then Err.Raise vbObjectError + 1291, , "Enregistrez d abord les modifications de la fiche."
    modFichePatient.ConfirmerIdentiteCourrier mDocument, mPatient
    AfficherEtatIdentite
    Exit Sub
Echec:
    MsgBox Err.Description, vbExclamation, "Identite du courrier"
End Sub

Private Sub btnFermer_Click()
    If AutoriserFermeture() Then Me.Hide
End Sub

Private Function AutoriserFermeture() As Boolean
    AutoriserFermeture = True
    If SaisieModifiee() Then AutoriserFermeture = (MsgBox("Fermer sans enregistrer les modifications ?", vbYesNo + vbQuestion + vbDefaultButton2, "Fiche patient") = vbYes)
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        If AutoriserFermeture() Then Me.Hide
    End If
End Sub
