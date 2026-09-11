Attribute VB_Name = "ufListe"
Attribute VB_Base = "0{1E8E2075-E4BE-43F1-9956-2822C8B269E5}{F19BDF6F-F9E1-440A-933B-26A4893B64F0}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
' Selecteur generique : liste filtrable en tapant (casse/accents ignores)
Private mItems As Collection        ' Scripting.Dictionary par element
Private mCles As Variant            ' colonnes affichees (noms de champs)
Private mFiltre As Collection       ' elements actuellement affiches
Public Resultat As Object
Public Annule As Boolean
Public NouveauDemande As Boolean

Public Sub Configurer(ByVal titre As String, ByVal items As Collection, _
                      ByVal cles As Variant, Optional ByVal largeurs As String = "", _
                      Optional ByVal preselectionID As String = "", _
                      Optional ByVal afficherNouveau As Boolean = False)
    Me.Caption = titre
    Set mItems = items
    mCles = cles
    lstItems.ColumnCount = UBound(mCles) - LBound(mCles) + 1
    If Len(largeurs) > 0 Then lstItems.ColumnWidths = largeurs
    Annule = True
    NouveauDemande = False
    Set Resultat = Nothing
    btnNouveau.visible = afficherNouveau
    Remplir ""
    If Len(preselectionID) > 0 Then Preselectionner preselectionID
End Sub

Private Sub btnNouveau_Click()
    NouveauDemande = True
    Annule = True
    Me.Hide
End Sub

Private Sub Preselectionner(ByVal id As String)
    Dim i As Long
    For i = 1 To mFiltre.Count
        If mFiltre(i)("ID") = id Then
            lstItems.ListIndex = i - 1
            Exit Sub
        End If
    Next i
End Sub

Private Sub Remplir(ByVal filtre As String)
    Dim it As Object, j As Long, visible As Boolean, f As String
    lstItems.Clear
    Set mFiltre = New Collection
    f = modTexte.Plier(Trim$(filtre))
    For Each it In mItems
        visible = (Len(f) = 0)
        If Not visible Then
            For j = LBound(mCles) To UBound(mCles)
                If InStr(modTexte.Plier(CStr(it(mCles(j)))), f) > 0 Then
                    visible = True
                    Exit For
                End If
            Next j
        End If
        If visible Then
            mFiltre.Add it
            lstItems.AddItem ""
            For j = LBound(mCles) To UBound(mCles)
                lstItems.List(lstItems.ListCount - 1, j - LBound(mCles)) = CStr(it(mCles(j)))
            Next j
        End If
    Next it
    If lstItems.ListCount > 0 Then lstItems.ListIndex = 0
End Sub

Private Sub Valider()
    If lstItems.ListIndex >= 0 Then
        Set Resultat = mFiltre(lstItems.ListIndex + 1)
        Annule = False
        Me.Hide
    End If
End Sub

Private Sub txtRecherche_Change()
    Remplir txtRecherche.Text
End Sub

Private Sub txtRecherche_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lstItems.ListIndex < lstItems.ListCount - 1 Then lstItems.ListIndex = lstItems.ListIndex + 1
            KeyCode = 0
        Case vbKeyUp
            If lstItems.ListIndex > 0 Then lstItems.ListIndex = lstItems.ListIndex - 1
            KeyCode = 0
        Case vbKeyReturn
            Valider
            KeyCode = 0
    End Select
End Sub

Private Sub lstItems_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Valider
End Sub

Private Sub btnOK_Click()
    Valider
End Sub

Private Sub btnAnnuler_Click()
    Annule = True
    Me.Hide
End Sub
