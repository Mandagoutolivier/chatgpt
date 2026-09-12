Attribute VB_Name = "ufPatientEdit"
Attribute VB_Base = "0{27A97239-3F29-4D90-925A-4B401F1324FB}{45DEA35E-DDC8-48F3-95C4-BADDDF7CA47E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
' Fiche patient : creation / modification (poste secretaire).
Private mRevision As String
Private mID As String               ' vide = nouveau patient
Private mIDsMedecins() As String

Public Sub ChargerNouveau()
    mID = ""
    Me.Caption = "Nouveau patient"
    lblID.Caption = ""
    ChargerCombos
    cmbSexe.ListIndex = -1
End Sub

Public Sub ChargerExistant(ByVal pat As Object)
    mRevision = CStr(pat("_revision"))
    mID = pat("ID")
    Me.Caption = "Patient " & pat("Prenom") & " " & pat("Nom")
    lblID.Caption = "Dossier " & mID
    ChargerCombos
    txtNom.Text = pat("Nom")
    txtNomN.Text = pat("NomNaissance")
    txtPrenom.Text = pat("Prenom")
    txtDDN.Text = pat("DDN")
    SelectionnerCombo cmbSexe, pat("Sexe")
    txtNIR.Text = pat("NIR")
    txtAdr1.Text = pat("Adresse1")
    txtAdr2.Text = pat("Adresse2")
    txtCP.Text = pat("CP")
    txtVille.Text = pat("Ville")
    txtTel.Text = pat("Tel")
    txtMobile.Text = pat("Mobile")
    txtEmail.Text = pat("Email")
    SelectionnerMedecin pat("MedTraitantID")
    txtMutuelle.Text = pat("Mutuelle")
    chkALD.Value = (pat("ALD") = "O")
    txtNotes.Text = pat("Notes")
End Sub

Private Sub ChargerCombos()
    Dim cors As Collection, c As Object, i As Long
    cmbSexe.Clear
    cmbSexe.AddItem "M"
    cmbSexe.AddItem "F"
    cmbMed.Clear
    cmbMed.AddItem "(aucun)"
    Set cors = modBaseIO.LireTableX(modConfig.FichierPatients(), "CORRESPONDANTS")
    ReDim mIDsMedecins(0 To cors.Count)
    mIDsMedecins(0) = ""
    i = 1
    For Each c In cors
        cmbMed.AddItem c("Titre") & " " & c("Nom") & " " & c("Prenom") & " (" & c("Ville") & ")"
        mIDsMedecins(i) = c("ID")
        i = i + 1
    Next c
    cmbMed.ListIndex = 0
End Sub

Private Sub SelectionnerMedecin(ByVal id As String)
    Dim i As Long
    For i = 0 To UBound(mIDsMedecins)
        If mIDsMedecins(i) = id Then
            cmbMed.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub SelectionnerCombo(ByVal cmb As MSForms.ComboBox, ByVal Valeur As String)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = Valeur Then cmb.ListIndex = i: Exit Sub
    Next i
End Sub

Private Sub btnOK_Click()
    On Error GoTo Erreur
    Dim d As Object, nouveauID As String
    If Len(Trim$(txtNom.Text)) = 0 Or Len(Trim$(txtPrenom.Text)) = 0 Then
        MsgBox "Le nom et le prenom sont obligatoires.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    If Not DateValide(txtDDN.Text) Then
        MsgBox "Date de naissance invalide : utilisez le format jj/mm/aaaa.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    If Len(modTexte.SexeNormalise(cmbSexe.Text)) = 0 Then
        MsgBox "Choisissez le sexe utilise pour l ECG (M ou F).", vbExclamation, "Cabinet"
        Exit Sub
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d("Nom") = UCase$(Trim$(txtNom.Text))
    d("NomNaissance") = UCase$(Trim$(txtNomN.Text))
    d("Prenom") = Trim$(txtPrenom.Text)
    d("DDN") = Format$(modTexte.DateFr(txtDDN.Text), "dd/mm/yyyy")
    d("Sexe") = modTexte.SexeNormalise(cmbSexe.Text)
    d("NIR") = Trim$(txtNIR.Text)
    d("Adresse1") = Trim$(txtAdr1.Text)
    d("Adresse2") = Trim$(txtAdr2.Text)
    d("CP") = Trim$(txtCP.Text)
    d("Ville") = Trim$(txtVille.Text)
    d("Tel") = Trim$(txtTel.Text)
    d("Mobile") = Trim$(txtMobile.Text)
    d("Email") = Trim$(txtEmail.Text)
    If cmbMed.ListIndex >= 0 Then d("MedTraitantID") = mIDsMedecins(cmbMed.ListIndex)
    d("Mutuelle") = Trim$(txtMutuelle.Text)
    d("ALD") = IIf(chkALD.Value, "O", "")
    d("Notes") = Trim$(txtNotes.Text)
    If Len(mID) = 0 Then
        d("DateCreation") = Format$(Now, "dd/mm/yyyy")
        nouveauID = modBaseIO.AjouterLigne(modConfig.FichierPatients(), "PATIENTS", d, "P")
        MsgBox "Patient cree : " & d("Prenom") & " " & d("Nom") & " (" & nouveauID & ")", _
               vbInformation, "Cabinet"
    Else
        d("DateModif") = Format$(Now, "dd/mm/yyyy")
        d("_revision") = mRevision
        modBaseIO.ModifierLigne modConfig.FichierPatients(), "PATIENTS", "ID", mID, d
        MsgBox "Fiche mise a jour.", vbInformation, "Cabinet"
    End If
    Me.Hide
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Private Function DateValide(ByVal s As String) As Boolean
    If Not modTexte.DateFrValide(s) Then Exit Function
    DateValide = (modTexte.DateFr(s) <= Date)
End Function
Private Sub btnAgenda_Click()
    Dim d As Object
    If Len(mID) = 0 Then
        MsgBox "Enregistrez d'abord la fiche du patient, puis reprenez-la pour fixer le rendez-vous.", _
               vbInformation, "Cabinet"
        Exit Sub
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("ID") = mID
    d("Nom") = UCase$(Trim$(txtNom.Text))
    d("Prenom") = Trim$(txtPrenom.Text)
    d("DDN") = Format$(modTexte.DateFr(txtDDN.Text), "dd/mm/yyyy")
    Me.Hide
    modAgendaVue.AfficherAgenda Date, d
End Sub

Private Sub btnAnnuler_Click()
    Me.Hide
End Sub
