Attribute VB_Name = "ufCorrespEdit"
Attribute VB_Base = "0{70E11178-5DDC-4794-A787-CEFAE6676BD8}{6140D122-4C1C-428C-84B3-828351D91B6B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
' Fiche correspondant : creation / modification (poste secretaire).
Private mRevision As String
Private mID As String

Public Sub ChargerNouveau()
    mID = ""
    Me.Caption = "Nouveau correspondant"
    lblID.Caption = ""
    ChargerCombos
    txtAppel.Text = "Cher Confrère,"
    txtPolit.Text = "Confraternellement."
    chkActif.Value = True
End Sub

Public Sub ChargerExistant(ByVal cor As Object)
    mRevision = CStr(cor("_revision"))
    mID = cor("ID")
    Me.Caption = "Correspondant " & cor("Titre") & " " & cor("Nom")
    lblID.Caption = "Fiche " & mID
    ChargerCombos
    SelectionnerCombo cmbTitre, cor("Titre")
    txtNom.Text = cor("Nom")
    txtPrenom.Text = cor("Prenom")
    txtSpec.Text = cor("Specialite")
    txtAdr1.Text = cor("Adresse1")
    txtAdr2.Text = cor("Adresse2")
    txtCP.Text = cor("CP")
    txtVille.Text = cor("Ville")
    txtTel.Text = cor("Tel")
    txtEmail.Text = cor("Email")
    txtAppel.Text = cor("FormuleAppel")
    txtPolit.Text = cor("FormulePolitesse")
    chkActif.Value = (cor("Actif") <> "0")
End Sub

Private Sub ChargerCombos()
    cmbTitre.Clear
    cmbTitre.AddItem "Dr"
    cmbTitre.AddItem "Pr"
    cmbTitre.AddItem "M."
    cmbTitre.AddItem "Mme"
    cmbTitre.ListIndex = 0
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
    If Len(Trim$(txtNom.Text)) = 0 Then
        MsgBox "Le nom est obligatoire.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d("Titre") = cmbTitre.Text
    d("Nom") = UCase$(Trim$(txtNom.Text))
    d("Prenom") = Trim$(txtPrenom.Text)
    d("Specialite") = Trim$(txtSpec.Text)
    d("Adresse1") = Trim$(txtAdr1.Text)
    d("Adresse2") = Trim$(txtAdr2.Text)
    d("CP") = Trim$(txtCP.Text)
    d("Ville") = Trim$(txtVille.Text)
    d("Tel") = Trim$(txtTel.Text)
    d("Email") = Trim$(txtEmail.Text)
    d("FormuleAppel") = Trim$(txtAppel.Text)
    d("FormulePolitesse") = Trim$(txtPolit.Text)
    d("Actif") = IIf(chkActif.Value, "1", "0")
    If Len(mID) = 0 Then
        nouveauID = modBaseIO.AjouterLigne(modConfig.FichierPatients(), "CORRESPONDANTS", d, "C")
        MsgBox "Correspondant cree (" & nouveauID & ").", vbInformation, "Cabinet"
    Else
        d("_revision") = mRevision
        modBaseIO.ModifierLigne modConfig.FichierPatients(), "CORRESPONDANTS", "ID", mID, d
        MsgBox "Fiche mise a jour.", vbInformation, "Cabinet"
    End If
    Me.Hide
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Private Sub btnAnnuler_Click()
    Me.Hide
End Sub
