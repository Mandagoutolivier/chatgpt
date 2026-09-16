Attribute VB_Name = "ufChoixActe"
Attribute VB_Base = "0{66A52417-0F69-46E9-8342-4E063EF83378}{D03DA56A-85A5-4249-839E-6D8013692B64}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
' Traitement d'un courrier valide : choix des actes, journal, feuille de soins.
Private mDrapeau As Object          ' dictionnaire du fichier-drapeau
Private mNomenclature As Collection
Private mInit As Boolean            ' vrai pendant le remplissage de la liste

Public Sub Charger(ByVal drapeau As Object)
    Dim modes As Variant, m As Variant
    Set mDrapeau = drapeau
    lblInfo.Caption = Valeur("Prenom") & " " & Valeur("Nom") & _
                      IIf(Len(Valeur("DDN")) > 0, " - " & Valeur("DDN"), "") & vbCrLf & _
                      "Courrier : " & Valeur("TypeCourrier") & "  (" & Valeur("DateValidation") & ")"
    lstItemsInit
    cmbPaiement.Clear
    modes = Array("CB", "Cheque", "Especes", "Virement", "Impaye")
    For Each m In modes
        cmbPaiement.AddItem CStr(m)
    Next m
    cmbPaiement.ListIndex = 0
    chkFds.Value = True
    MajTotal
End Sub

Private Sub lstItemsInit()
    Dim a As Object, i As Long, typeCourrier As String
    mInit = True
    Set mNomenclature = modActes.Nomenclature()
    lstActes.Clear
    lstActes.ColumnCount = 3
    lstActes.ColumnWidths = "60 pt;250 pt;60 pt"
    lstActes.MultiSelect = 1            ' fmMultiSelectMulti
    lstActes.ListStyle = 1              ' fmListStyleOption (cases a cocher)
    i = 0
    typeCourrier = LCase$(Valeur("TypeCourrier"))
    For Each a In mNomenclature
        lstActes.AddItem ""
        lstActes.List(i, 0) = a("Code")
        lstActes.List(i, 1) = a("Libelle") & IIf(Len(a("CodeAssocie")) > 0, " (+ " & a("CodeAssocie") & ")", "")
        lstActes.List(i, 2) = Format$(Val(Replace(a("Tarif"), ",", ".")) + Val(Replace(a("TarifAssocie"), ",", ".")), "0.00")
        ' preselection simple : consultation -> CSC
        If InStr(typeCourrier, "consultation") > 0 And a("Code") = "CSC" Then lstActes.Selected(i) = True
        i = i + 1
    Next a
    mInit = False
End Sub

Private Function Valeur(ByVal cle As String) As String
    If mDrapeau.Exists(cle) Then Valeur = mDrapeau(cle) Else Valeur = ""
End Function

Private Function ActesChoisis() As Collection
    Dim col As New Collection, j As Long, a As Object
    If mNomenclature Is Nothing Then Set ActesChoisis = col: Exit Function
    j = 0
    For Each a In mNomenclature
        If j < lstActes.ListCount Then
            If lstActes.Selected(j) Then col.Add a
        End If
        j = j + 1
    Next a
    Set ActesChoisis = col
End Function

Private Sub MajTotal()
    On Error Resume Next
    lblTotal.Caption = "Total : " & Format$(modActes.TotalActes(ActesChoisis()), "0.00") & " EUR"
End Sub

Private Sub lstActes_Change()
    If mInit Then Exit Sub
    MajTotal
End Sub

Private Sub btnOuvrir_Click()
    modEchange.OuvrirCourrier mDrapeau
End Sub

Private Sub AffecterPayeur(ByVal resultat As Object, ByVal code As String, ByVal organisme As Boolean)
    code = Trim$(code)
    If Len(code) > 0 Then resultat(code) = organisme
End Sub

Private Sub AffecterTousPayeurs(ByVal actes As Collection, ByVal resultat As Object, ByVal organisme As Boolean)
    Dim a As Object
    For Each a In actes
        AffecterPayeur resultat, CStr(a("Code")), organisme
        If Len(CStr(a("CodeAssocie"))) > 0 Then AffecterPayeur resultat, CStr(a("CodeAssocie")), organisme
    Next a
End Sub

Private Function ChoisirPayeurActe(ByVal acte As Object, ByRef annule As Boolean) As Boolean
    Dim reponse As VbMsgBoxResult, libelle As String, total As Double
    libelle = CStr(acte("Code"))
    total = Val(Replace(CStr(acte("Tarif")), ",", "."))
    If Len(CStr(acte("CodeAssocie"))) > 0 Then
        libelle = libelle & " (+ " & CStr(acte("CodeAssocie")) & ")"
        total = total + Val(Replace(CStr(acte("TarifAssocie")), ",", "."))
    End If
    reponse = MsgBox("Qui doit regler l acte " & libelle & " (" & Format$(total, "0.00") & " EUR) ?" & vbCrLf & _
                     "Oui : Organisme" & vbCrLf & "Non : Patient" & vbCrLf & "Annuler : aucune ecriture", _
                     vbYesNoCancel + vbQuestion, "Repartition du tiers payant")
    If reponse = vbCancel Then annule = True: Exit Function
    ChoisirPayeurActe = (reponse = vbYes)
End Function

Private Function ChoisirRepartitionTiers(ByVal actes As Collection, ByRef annule As Boolean) As Object
    Dim resultat As Object, reponse As VbMsgBoxResult, a As Object, organisme As Boolean, resume As String
    Set resultat = CreateObject("Scripting.Dictionary"): resultat.CompareMode = 1
    If Not CBool(chkTiers.Value) Then
        AffecterTousPayeurs actes, resultat, False
        Set ChoisirRepartitionTiers = resultat
        Exit Function
    End If
    If actes.Count = 1 Then
        AffecterTousPayeurs actes, resultat, True
        Set ChoisirRepartitionTiers = resultat
        Exit Function
    End If
    reponse = MsgBox("Le tiers payant concerne-t-il tous les actes selectionnes ?" & vbCrLf & _
                     "Oui : tous Organisme" & vbCrLf & "Non : choisir acte par acte" & vbCrLf & _
                     "Annuler : aucune ecriture", vbYesNoCancel + vbQuestion, "Repartition du tiers payant")
    If reponse = vbCancel Then annule = True: Set ChoisirRepartitionTiers = resultat: Exit Function
    If reponse = vbYes Then
        AffecterTousPayeurs actes, resultat, True
        Set ChoisirRepartitionTiers = resultat
        Exit Function
    End If
    For Each a In actes
        organisme = ChoisirPayeurActe(a, annule)
        If annule Then Set ChoisirRepartitionTiers = resultat: Exit Function
        AffecterPayeur resultat, CStr(a("Code")), organisme
        resume = resume & CStr(a("Code")) & " : " & IIf(organisme, "Organisme", "Patient") & vbCrLf
        If Len(CStr(a("CodeAssocie"))) > 0 Then
            AffecterPayeur resultat, CStr(a("CodeAssocie")), organisme
            resume = resume & "  + " & CStr(a("CodeAssocie")) & " : meme payeur" & vbCrLf
        End If
    Next a
    If MsgBox("Confirmez la repartition :" & vbCrLf & vbCrLf & resume, vbYesNo + vbQuestion, _
              "Repartition du tiers payant") <> vbYes Then annule = True
    Set ChoisirRepartitionTiers = resultat
End Function

Private Sub btnOK_Click()
    On Error GoTo Erreur
    Dim actes As Collection, seanceID As String, a As Object, tarifZero As Boolean, deja As Boolean
    Set actes = ActesChoisis()
    If actes.Count = 0 Then
        MsgBox "Cochez au moins un acte.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    For Each a In actes
        If Val(Replace(a("Tarif"), ",", ".")) = 0 Then tarifZero = True
    Next a
    If tarifZero Then
        If MsgBox("Au moins un acte a un tarif a 0 EUR (nomenclature a completer : " & _
                  "Config\Nomenclature.xlsx)." & vbCrLf & "Continuer quand meme ?", _
                  vbYesNo + vbExclamation, "Cabinet") <> vbYes Then Exit Sub
    End If

    Dim repartition As Object, annulePayeur As Boolean
    Set repartition = ChoisirRepartitionTiers(actes, annulePayeur)
    If annulePayeur Then Exit Sub
    seanceID = Valeur("ConsultationID")
    If Len(seanceID) = 0 Then seanceID = Valeur("SeanceID")
    If chkFds.Value Then
        If Not modActes.ExisteSeanceEnregistree(seanceID) Then _
            modCerfaPrint.VerifierAvantFacturation mDrapeau, modActes.LignesPourImpression(actes)
    End If
    seanceID = modActes.EnregistrerSeance(mDrapeau, actes, cmbPaiement.Text, _
                                          repartition, False, deja)

    Dim reponse As VbMsgBoxResult, impressionEnvoyee As Boolean
    If deja And chkFds.Value Then
        reponse = MsgBox("Cette consultation figure deja au journal. Reimprimer la feuille avec les actes et montants enregistres ?" & vbCrLf & _
                        "Oui : reimprimer. Non : terminer sans reimprimer. Annuler : conserver le courrier en attente.", vbYesNoCancel + vbQuestion, "Reprise d une consultation")
        If reponse = vbCancel Then Exit Sub
        If reponse = vbYes Then
            modActes.ImprimerSeanceEnregistree mDrapeau, seanceID
            impressionEnvoyee = True
        End If
    ElseIf Not deja And chkFds.Value Then
        modActes.ImprimerSeanceEnregistree mDrapeau, seanceID
        impressionEnvoyee = True
    End If

    If mDrapeau.Exists("_Chemin") Then modEchange.DeplacerVersTraites mDrapeau("_Chemin")

    MsgBox "Consultation traitee" & IIf(impressionEnvoyee, " + impression envoyee", "") & _
           "." & vbCrLf & "(" & Valeur("Prenom") & " " & Valeur("Nom") & ", " & actes.Count & _
           " acte(s) coche(s))", vbInformation, "Cabinet"
    Me.Hide
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Private Sub btnAnnuler_Click()
    Me.Hide
End Sub



Public Sub ReimprimerFeuille()
    On Error GoTo Echec
    modActes.ImprimerSeanceEnregistree mDrapeau, Valeur("ConsultationID")
    Exit Sub
Echec:
    MsgBox Err.Description, vbExclamation, "Reimpression"
End Sub
