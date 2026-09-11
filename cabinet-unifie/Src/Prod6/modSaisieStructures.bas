Attribute VB_Name = "modSaisieStructures"
Option Explicit

Private mNSHandlersTextes As Collection
Private mNSHandlersBoutons As Collection
Private mNSStructures As Collection

Private Const NS_XL_UP As Long = -4162

Public gNSDernierIDStructure As String
Public gNSDernierNomStructure As String

Public Sub NS_OuvrirNouvelleStructure()
    Dim frm As Object

    On Error GoTo GestionErreur

    gNSDernierIDStructure = ""
    gNSDernierNomStructure = ""

    On Error Resume Next
    Unload frmNouvelleStructure
    On Error GoTo GestionErreur

    Set frm = frmNouvelleStructure
    Set mNSHandlersTextes = New Collection
    Set mNSHandlersBoutons = New Collection
    Set mNSStructures = New Collection

    NS_ConstruireInterface frm
    NS_ChargerStructuresExistantes frm

    frm.Show
    Exit Sub

GestionErreur:
    MsgBox _
        "Impossible d'ouvrir le formulaire Nouvelle structure." & vbCrLf & _
        Err.Number & " - " & Err.description, _
        vbExclamation, _
        "Nouvelle structure"
End Sub

Private Sub NS_ConstruireInterface(ByVal frm As Object)
    Dim i As Long
    Dim ctl As Object

    On Error Resume Next
    For i = frm.Controls.Count - 1 To 0 Step -1
        frm.Controls.Remove frm.Controls(i).Name
    Next i
    On Error GoTo 0

    With frm
        .Caption = "Nouvelle structure"
        .Width = 560
        .Height = 515
        .StartUpPosition = 1
    End With

    Set ctl = NS_AjouterLabel(frm, "lblNSTitre", "Nouvelle structure", 16, 12, 500, 22, True)
    ctl.Font.Size = 12

    Set ctl = NS_AjouterLabel( _
        frm, "lblNSInfo", _
        "Renseignez la structure, puis choisissez Enregistrer.", _
        16, 38, 500, 30, False)
    ctl.WordWrap = True

    NS_AjouterLabel frm, "lblNSNom", "Nom de la structure", 16, 82, 190, 16, True
    NS_AjouterTexte frm, "txtNSNom", 16, 100, 500, 22, "NomStructure"

    NS_AjouterLabel frm, "lblNSAdr1", "Adresse", 16, 136, 120, 16, True
    NS_AjouterTexte frm, "txtNSAdresse1", 16, 154, 500, 22, "Adresse1"
    NS_AjouterTexte frm, "txtNSAdresse2", 16, 180, 500, 22, "Adresse2"

    NS_AjouterLabel frm, "lblNSCP", "Code postal", 16, 216, 90, 16, True
    NS_AjouterLabel frm, "lblNSVille", "Ville", 120, 216, 120, 16, True
    NS_AjouterTexte frm, "txtNSCodePostal", 16, 234, 90, 22, "CodePostal"
    NS_AjouterTexte frm, "txtNSVille", 120, 234, 396, 22, "Ville"

    NS_AjouterLabel frm, "lblNSTel", "Téléphone", 16, 270, 120, 16, True
    NS_AjouterTexte frm, "txtNSTelephone", 16, 288, 220, 22, "Telephone"

    NS_AjouterLabel frm, "lblNSDoublons", "Contrôle des doublons", 16, 326, 220, 16, True
    Set ctl = NS_AjouterTexte(frm, "txtNSDoublons", 16, 344, 500, 92, "")
    With ctl
        .MultiLine = True
        .WordWrap = True
        .Locked = True
        .TabStop = False
        .ScrollBars = 2
        .BackColor = &H8000000F
        .Text = "Saisissez au minimum le nom de la structure. Le contrôle peut être déclenché avec Tab ou avec le bouton."
    End With

    NS_AjouterBouton frm, "cmdNSVerifier", "Vérifier les doublons", 16, 450, 145, 28, "Verifier"
    NS_AjouterBouton frm, "cmdNSEnregistrer", "Enregistrer", 300, 450, 105, 28, "Enregistrer"
    NS_AjouterBouton frm, "cmdNSFermer", "Fermer", 420, 450, 95, 28, "Fermer"

    Set ctl = NS_AjouterLabel( _
        frm, "lblNSAucunEnregistrement", _
        "Enregistrement bloqué en cas de doublon certain.", _
        170, 446, 120, 38, False)
    ctl.WordWrap = True
End Sub

Private Function NS_AjouterLabel( _
    ByVal frm As Object, ByVal nom As String, ByVal texte As String, _
    ByVal gauche As Single, ByVal haut As Single, _
    ByVal largeur As Single, ByVal hauteur As Single, _
    Optional ByVal gras As Boolean = False) As Object

    Dim ctl As Object
    Set ctl = frm.Controls.Add("Forms.Label.1", nom, True)

    With ctl
        .Caption = texte
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        .Font.Size = 9
        .Font.Bold = gras
    End With

    Set NS_AjouterLabel = ctl
End Function

Private Function NS_AjouterTexte( _
    ByVal frm As Object, ByVal nom As String, _
    ByVal gauche As Single, ByVal haut As Single, _
    ByVal largeur As Single, ByVal hauteur As Single, _
    ByVal role As String) As Object

    Dim ctl As Object
    Dim h As clsNSTextBox

    Set ctl = frm.Controls.Add("Forms.TextBox.1", nom, True)

    With ctl
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        .Font.Size = 9
    End With

    If Trim$(role) <> "" Then
        Set h = New clsNSTextBox
        Set h.ctl = ctl
        h.role = role
        mNSHandlersTextes.Add h
    End If

    Set NS_AjouterTexte = ctl
End Function

Private Sub NS_AjouterBouton( _
    ByVal frm As Object, ByVal nom As String, ByVal texte As String, _
    ByVal gauche As Single, ByVal haut As Single, _
    ByVal largeur As Single, ByVal hauteur As Single, _
    ByVal role As String)

    Dim ctl As Object
    Dim h As clsNSBouton

    Set ctl = frm.Controls.Add("Forms.CommandButton.1", nom, True)

    With ctl
        .Caption = texte
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        .Font.Size = 9
    End With

    Set h = New clsNSBouton
    Set h.ctl = ctl
    h.role = role
    mNSHandlersBoutons.Add h
End Sub

Private Sub NS_ChargerStructuresExistantes(ByVal frm As Object)
    Dim excelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim colID As Long, colNom As Long, colAdresse1 As Long
    Dim colAdresse2 As Long, colCP As Long, colVille As Long
    Dim colTel As Long, colActif As Long
    Dim derniereLigne As Long, i As Long
    Dim actif As String
    Dim d As Object

    On Error GoTo GestionErreur

    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    excelApp.AutomationSecurity = 3
    excelApp.EnableEvents = False

    Set wb = excelApp.Workbooks.Open(CHEMIN_BASE_CORRESPONDANTS, 0, True)
    Set ws = wb.Worksheets("Structures")

    colID = NS_TrouverColonneEntete(ws, "ID_Structure")
    colNom = NS_TrouverColonneEntete(ws, "NomStructure")
    colAdresse1 = NS_TrouverColonneEntete(ws, "Adresse1")
    colAdresse2 = NS_TrouverColonneEntete(ws, "Adresse2")
    colCP = NS_TrouverColonneEntete(ws, "CodePostal")
    colVille = NS_TrouverColonneEntete(ws, "Ville")
    colTel = NS_TrouverColonneEntete(ws, "Telephone")
    colActif = NS_TrouverColonneEntete(ws, "Actif")

    If colID = 0 Or colNom = 0 Then
        Err.Raise vbObjectError + 601, , "Colonnes ID_Structure ou NomStructure absentes."
    End If

    Set mNSStructures = New Collection
    derniereLigne = ws.Cells(ws.Rows.Count, colID).End(NS_XL_UP).Row

    For i = 2 To derniereLigne
        If colActif > 0 Then
            actif = NS_LireCellule(ws, i, colActif)
            If Not NS_EstOui(actif) Then GoTo LigneSuivante
        End If

        If Trim$(NS_LireCellule(ws, i, colNom)) = "" Then GoTo LigneSuivante

        Set d = CreateObject("Scripting.Dictionary")
        d.Add "ID", NS_LireCellule(ws, i, colID)
        d.Add "Nom", NS_LireCellule(ws, i, colNom)
        d.Add "Adresse1", NS_LireCellule(ws, i, colAdresse1)
        d.Add "Adresse2", NS_LireCellule(ws, i, colAdresse2)
        d.Add "CodePostal", NS_LireCellule(ws, i, colCP)
        d.Add "Ville", NS_LireCellule(ws, i, colVille)
        d.Add "Telephone", NS_LireCellule(ws, i, colTel)
        mNSStructures.Add d

LigneSuivante:
    Next i

    frm.Controls("lblNSInfo").Caption = _
        "Structures : " & CStr(mNSStructures.Count) & _
        " structure(s) active(s) chargée(s)."

Sortie:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not excelApp Is Nothing Then excelApp.Quit
    Set wb = Nothing
    Set excelApp = Nothing
    On Error GoTo 0
    Exit Sub

GestionErreur:
    MsgBox _
        "Impossible de lire la feuille Structures." & vbCrLf & _
        Err.Number & " - " & Err.description, _
        vbExclamation, _
        "Nouvelle structure"
    Resume Sortie
End Sub

Public Sub NS_EvenementBouton(ByVal role As String)
    Select Case role
        Case "Verifier"
            NS_VerifierDoublonsFormulaire
        Case "Enregistrer"
            NS_EnregistrerNouvelleStructure
        Case "Fermer"
            Unload frmNouvelleStructure
    End Select
End Sub

Public Sub NS_EvenementTexteTab(ByVal role As String)
    If Trim$(frmNouvelleStructure.Controls("txtNSNom").Text) <> "" Then
        NS_VerifierDoublonsFormulaire
    End If
End Sub

Public Sub NS_VerifierDoublonsFormulaire()
    Dim frm As Object
    Dim d As Object
    Dim nomRecherche As String, adresseRecherche As String
    Dim cpRecherche As String, villeRecherche As String
    Dim telRecherche As String
    Dim nomExistant As String, adresseExistante As String
    Dim cpExistant As String, villeExistante As String
    Dim telExistant As String
    Dim certain As Boolean, probable As Boolean
    Dim nbCertains As Long, nbProbables As Long
    Dim txtCertains As String, txtProbables As String
    Dim resultat As String

    Set frm = frmNouvelleStructure

    nomRecherche = NS_NormaliserTexte(frm.Controls("txtNSNom").Text)
    adresseRecherche = NS_NormaliserTexte(frm.Controls("txtNSAdresse1").Text)
    cpRecherche = NS_NormaliserCP(frm.Controls("txtNSCodePostal").Text)
    villeRecherche = NS_NormaliserTexte(frm.Controls("txtNSVille").Text)
    telRecherche = NS_NormaliserTelephone(frm.Controls("txtNSTelephone").Text)

    If nomRecherche = "" Then
        frm.Controls("txtNSDoublons").Text = "Saisissez au minimum le nom de la structure."
        Exit Sub
    End If

    For Each d In mNSStructures
        nomExistant = NS_NormaliserTexte(CStr(d("Nom")))
        adresseExistante = NS_NormaliserTexte(CStr(d("Adresse1")))
        cpExistant = NS_NormaliserCP(CStr(d("CodePostal")))
        villeExistante = NS_NormaliserTexte(CStr(d("Ville")))
        telExistant = NS_NormaliserTelephone(CStr(d("Telephone")))

        certain = False
        probable = False

        If nomExistant <> "" And nomExistant = nomRecherche Then
            certain = True
        ElseIf telRecherche <> "" And telExistant <> "" And telRecherche = telExistant Then
            certain = True
        ElseIf adresseRecherche <> "" And adresseExistante <> "" _
        And adresseRecherche = adresseExistante _
        And villeRecherche <> "" And villeExistante = villeRecherche Then
            probable = True
        ElseIf cpRecherche <> "" And cpExistant = cpRecherche _
        And villeRecherche <> "" And villeExistante = villeRecherche _
        And NS_NomsProches(nomRecherche, nomExistant) Then
            probable = True
        ElseIf NS_NomsProches(nomRecherche, nomExistant) Then
            probable = True
        End If

        If certain Then
            nbCertains = nbCertains + 1
            txtCertains = txtCertains & NS_LigneStructure(d, "DOUBLON CERTAIN") & vbCrLf
        ElseIf probable Then
            nbProbables = nbProbables + 1
            txtProbables = txtProbables & NS_LigneStructure(d, "À VÉRIFIER") & vbCrLf
        End If
    Next d

    If nbCertains = 0 And nbProbables = 0 Then
        resultat = "Aucun doublon de structure détecté."
    Else
        resultat = _
            "Résultat du contrôle :" & vbCrLf & _
            "Doublon(s) certain(s) : " & nbCertains & vbCrLf & _
            "Correspondance(s) à vérifier : " & nbProbables & vbCrLf & vbCrLf

        If txtCertains <> "" Then resultat = resultat & txtCertains
        If txtProbables <> "" Then
            If txtCertains <> "" Then resultat = resultat & vbCrLf
            resultat = resultat & txtProbables
        End If

        resultat = resultat & vbCrLf & _
            "La création restera bloquée tant qu'un doublon certain n'aura pas été traité."
    End If

    frm.Controls("txtNSDoublons").Text = resultat
    frm.Controls("txtNSDoublons").SelStart = 0
End Sub

Private Sub NS_EnregistrerNouvelleStructure()
    Dim verrou As String, verrouAcquis As Boolean
    Dim frm As Object
    Dim nomStructure As String, adresse1 As String, adresse2 As String
    Dim codePostal As String, ville As String, telephone As String
    Dim rep As VbMsgBoxResult
    Dim excelApp As Object, wb As Object, ws As Object
    Dim colID As Long, nouvelleLigne As Long
    Dim nouvelID As String

    On Error GoTo GestionErreur

    Set frm = frmNouvelleStructure

    nomStructure = Trim$(frm.Controls("txtNSNom").Text)
    adresse1 = Trim$(frm.Controls("txtNSAdresse1").Text)
    adresse2 = Trim$(frm.Controls("txtNSAdresse2").Text)
    codePostal = Trim$(frm.Controls("txtNSCodePostal").Text)
    ville = Trim$(frm.Controls("txtNSVille").Text)
    telephone = Trim$(frm.Controls("txtNSTelephone").Text)

    If nomStructure = "" Then
        MsgBox "Le nom de la structure est obligatoire.", vbExclamation, "Nouvelle structure"
        frm.Controls("txtNSNom").SetFocus
        If verrouAcquis Then modFichiers.RelacherVerrou verrou
    Exit Sub
    End If

    NS_VerifierDoublonsFormulaire

    If NS_A_DoublonCertainMemoire( _
        nomStructure, adresse1, codePostal, ville, telephone) Then

        MsgBox _
            "Création impossible : une structure identique existe déjà.", _
            vbExclamation, _
            "Doublon de structure"
        If verrouAcquis Then modFichiers.RelacherVerrou verrou
    Exit Sub
    End If

    If NS_A_DoublonProbableMemoire( _
        nomStructure, adresse1, codePostal, ville, telephone) Then

        rep = MsgBox( _
            "Une structure proche existe déjà dans la base." & vbCrLf & vbCrLf & _
            "Confirmez-vous qu'il s'agit réellement d'une nouvelle structure ?", _
            vbExclamation + vbYesNo, _
            "Structure proche détectée")

        If rep <> vbYes Then Exit Sub
    End If

    rep = MsgBox( _
        "Créer la structure suivante ?" & vbCrLf & vbCrLf & _
        nomStructure & _
        IIf(adresse1 <> "", vbCrLf & adresse1, "") & _
        IIf(adresse2 <> "", vbCrLf & adresse2, "") & _
        IIf(Trim$(codePostal & " " & ville) <> "", _
            vbCrLf & Trim$(codePostal & " " & ville), "") & _
        IIf(telephone <> "", vbCrLf & "Tél. : " & telephone, ""), _
        vbQuestion + vbYesNo, _
        "Confirmer la création")

    If rep <> vbYes Then Exit Sub

    verrou = CreateObject("Scripting.FileSystemObject").GetBaseName(CHEMIN_BASE_CORRESPONDANTS)
    If Not modFichiers.AcquerirVerrou(verrou) Then Err.Raise vbObjectError + 726, , "Base correspondants occupee."
    verrouAcquis = True
    modFichiers.SauvegardeHorodatee CHEMIN_BASE_CORRESPONDANTS
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    excelApp.AutomationSecurity = 3
    excelApp.EnableEvents = False

    Set wb = excelApp.Workbooks.Open( _
        CHEMIN_BASE_CORRESPONDANTS, _
        0, _
        False)

    If wb.ReadOnly Then
        MsgBox _
            "La base Excel est en lecture seule." & vbCrLf & _
            "Fermez-la dans Excel puis recommencez.", _
            vbExclamation, _
            "Base correspondants"
        GoTo SortieSansSauver
    End If

    Set ws = wb.Worksheets("Structures")

    If NS_DoublonCertainDansFeuille( _
        ws, nomStructure, adresse1, codePostal, ville, telephone) Then

        MsgBox _
            "L'enregistrement est annulé : un doublon certain est maintenant présent dans la base.", _
            vbExclamation, _
            "Doublon détecté"
        GoTo SortieSansSauver
    End If

    colID = NS_TrouverColonneEntete(ws, "ID_Structure")
    If colID = 0 Then
        Err.Raise vbObjectError + 602, , "Colonne ID_Structure absente."
    End If

    nouvelID = NS_ProchainIDStructure(ws, colID)
    nouvelleLigne = ws.Cells(ws.Rows.Count, colID).End(NS_XL_UP).Row + 1

    NS_EcrireValeur ws, nouvelleLigne, "ID_Structure", nouvelID
    NS_EcrireValeur ws, nouvelleLigne, "NomStructure", nomStructure
    NS_EcrireValeur ws, nouvelleLigne, "Adresse1", adresse1
    NS_EcrireValeur ws, nouvelleLigne, "Adresse2", adresse2
    NS_EcrireValeur ws, nouvelleLigne, "CodePostal", codePostal
    NS_EcrireValeur ws, nouvelleLigne, "Ville", ville
    NS_EcrireValeur ws, nouvelleLigne, "Telephone", telephone
    NS_EcrireValeur ws, nouvelleLigne, "Actif", "Oui"
    NS_EcrireValeur ws, nouvelleLigne, "AValider", "Non"
    NS_EcrireValeur ws, nouvelleLigne, "CleNormalisee", NS_CleNormalisee(nomStructure)
    NS_EcrireValeur ws, nouvelleLigne, "SourceIDs", "FORMULAIRE_WORD"
    NS_EcrireValeur ws, nouvelleLigne, "Commentaire", _
        "Créé via formulaire Word le " & Format$(Date, "dd/mm/yyyy")

    wb.Save
    If Not wb.Saved Then Err.Raise vbObjectError + 727, , "Sauvegarde des correspondants non confirmee."

    gNSDernierIDStructure = nouvelID
    gNSDernierNomStructure = nomStructure

    MsgBox _
        "Structure enregistrée." & vbCrLf & vbCrLf & _
        nouvelID & " — " & nomStructure, _
        vbInformation, _
        "Nouvelle structure"

    wb.Close SaveChanges:=False
    Set wb = Nothing
    excelApp.Quit
    Set excelApp = Nothing

    Unload frmNouvelleStructure
    If verrouAcquis Then modFichiers.RelacherVerrou verrou
    Exit Sub

SortieSansSauver:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not excelApp Is Nothing Then excelApp.Quit
    Set ws = Nothing
    Set wb = Nothing
    Set excelApp = Nothing
    On Error GoTo 0
    If verrouAcquis Then modFichiers.RelacherVerrou verrou
    Exit Sub

GestionErreur:
    Dim numeroErreur As Long
    Dim descriptionErreur As String

    numeroErreur = Err.Number
    descriptionErreur = Err.description

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not excelApp Is Nothing Then excelApp.Quit
    Set ws = Nothing
    Set wb = Nothing
    Set excelApp = Nothing
    On Error GoTo 0

    MsgBox _
        "La structure n'a pas été enregistrée." & vbCrLf & _
        CStr(numeroErreur) & " - " & descriptionErreur, _
        vbExclamation, _
        "Nouvelle structure"
    If verrouAcquis Then modFichiers.RelacherVerrou verrou
End Sub

Private Function NS_A_DoublonCertainMemoire( _
    ByVal nom As String, ByVal adresse1 As String, _
    ByVal codePostal As String, ByVal ville As String, _
    ByVal telephone As String) As Boolean

    Dim d As Object
    Dim n As String, tel As String

    n = NS_NormaliserTexte(nom)
    tel = NS_NormaliserTelephone(telephone)

    For Each d In mNSStructures
        If NS_NormaliserTexte(CStr(d("Nom"))) = n Then
            NS_A_DoublonCertainMemoire = True
            Exit Function
        End If
        If tel <> "" And NS_NormaliserTelephone(CStr(d("Telephone"))) = tel Then
            NS_A_DoublonCertainMemoire = True
            Exit Function
        End If
    Next d
End Function

Private Function NS_A_DoublonProbableMemoire( _
    ByVal nom As String, ByVal adresse1 As String, _
    ByVal codePostal As String, ByVal ville As String, _
    ByVal telephone As String) As Boolean

    Dim d As Object
    Dim n As String, adr As String, cp As String, v As String

    n = NS_NormaliserTexte(nom)
    adr = NS_NormaliserTexte(adresse1)
    cp = NS_NormaliserCP(codePostal)
    v = NS_NormaliserTexte(ville)

    For Each d In mNSStructures
        If adr <> "" And NS_NormaliserTexte(CStr(d("Adresse1"))) = adr _
        And v <> "" And NS_NormaliserTexte(CStr(d("Ville"))) = v Then
            NS_A_DoublonProbableMemoire = True
            Exit Function
        End If

        If NS_NomsProches(n, NS_NormaliserTexte(CStr(d("Nom")))) Then
            NS_A_DoublonProbableMemoire = True
            Exit Function
        End If

        If cp <> "" And NS_NormaliserCP(CStr(d("CodePostal"))) = cp _
        And v <> "" And NS_NormaliserTexte(CStr(d("Ville"))) = v _
        And NS_NomsProches(n, NS_NormaliserTexte(CStr(d("Nom")))) Then
            NS_A_DoublonProbableMemoire = True
            Exit Function
        End If
    Next d
End Function

Private Function NS_DoublonCertainDansFeuille( _
    ByVal ws As Object, _
    ByVal nom As String, ByVal adresse1 As String, _
    ByVal codePostal As String, ByVal ville As String, _
    ByVal telephone As String) As Boolean

    Dim colID As Long, colNom As Long, colTel As Long, colActif As Long
    Dim derniereLigne As Long, i As Long
    Dim n As String, tel As String, actif As String

    colID = NS_TrouverColonneEntete(ws, "ID_Structure")
    colNom = NS_TrouverColonneEntete(ws, "NomStructure")
    colTel = NS_TrouverColonneEntete(ws, "Telephone")
    colActif = NS_TrouverColonneEntete(ws, "Actif")

    If colID = 0 Or colNom = 0 Then Exit Function

    n = NS_NormaliserTexte(nom)
    tel = NS_NormaliserTelephone(telephone)
    derniereLigne = ws.Cells(ws.Rows.Count, colID).End(NS_XL_UP).Row

    For i = 2 To derniereLigne
        If colActif > 0 Then
            actif = NS_LireCellule(ws, i, colActif)
            If Not NS_EstOui(actif) Then GoTo LigneSuivanteLive
        End If

        If NS_NormaliserTexte(NS_LireCellule(ws, i, colNom)) = n Then
            NS_DoublonCertainDansFeuille = True
            Exit Function
        End If

        If tel <> "" And colTel > 0 Then
            If NS_NormaliserTelephone(NS_LireCellule(ws, i, colTel)) = tel Then
                NS_DoublonCertainDansFeuille = True
                Exit Function
            End If
        End If

LigneSuivanteLive:
    Next i
End Function

Private Function NS_ProchainIDStructure(ByVal ws As Object, ByVal colID As Long) As String
    Dim derniereLigne As Long, i As Long
    Dim valeur As String, partie As String
    Dim n As Long, nMax As Long

    derniereLigne = ws.Cells(ws.Rows.Count, colID).End(NS_XL_UP).Row

    For i = 2 To derniereLigne
        valeur = UCase$(Trim$(NS_LireCellule(ws, i, colID)))
        If Left$(valeur, 4) = "STR_" Then
            partie = Mid$(valeur, 5)
            If IsNumeric(partie) Then
                n = CLng(partie)
                If n > nMax Then nMax = n
            End If
        End If
    Next i

    NS_ProchainIDStructure = "STR_" & Format$(nMax + 1, "0000")
End Function

Private Sub NS_EcrireValeur( _
    ByVal ws As Object, ByVal ligne As Long, _
    ByVal entete As String, ByVal valeur As String)

    Dim col As Long
    col = NS_TrouverColonneEntete(ws, entete)
    If col = 0 Then Err.Raise vbObjectError + 603, , "Colonne Structures absente : " & entete
    modFichiers.EcrireCelluleTexte ws.Cells(ligne, col), valeur
End Sub

Private Function NS_CleNormalisee(ByVal texte As String) As String
    Dim t As String
    t = UCase$(NS_NormaliserTexte(texte))
    t = Replace(t, " ", "_")
    Do While InStr(t, "__") > 0
        t = Replace(t, "__", "_")
    Loop
    NS_CleNormalisee = t
End Function

Private Function NS_LigneStructure(ByVal d As Object, ByVal niveau As String) As String
    Dim texte As String

    texte = "[" & niveau & "] " & CStr(d("ID")) & " — " & CStr(d("Nom"))

    If Trim$(CStr(d("Adresse1"))) <> "" Then
        texte = texte & " — " & CStr(d("Adresse1"))
    End If

    If Trim$(CStr(d("CodePostal")) & " " & CStr(d("Ville"))) <> "" Then
        texte = texte & " — " & Trim$(CStr(d("CodePostal")) & " " & CStr(d("Ville")))
    End If

    If Trim$(CStr(d("Telephone"))) <> "" Then
        texte = texte & " — " & CStr(d("Telephone"))
    End If

    NS_LigneStructure = texte
End Function

Private Function NS_NomsProches(ByVal nom1 As String, ByVal nom2 As String) As Boolean
    Dim compact1 As String, compact2 As String

    compact1 = Replace(nom1, " ", "")
    compact2 = Replace(nom2, " ", "")

    If Len(compact1) < 5 Or Len(compact2) < 5 Then Exit Function

    If InStr(1, compact1, compact2, vbTextCompare) > 0 _
    Or InStr(1, compact2, compact1, vbTextCompare) > 0 Then
        NS_NomsProches = True
        Exit Function
    End If

    If Len(compact1) >= 8 And Len(compact2) >= 8 Then
        If Left$(compact1, 8) = Left$(compact2, 8) Then
            NS_NomsProches = True
        End If
    End If
End Function

Private Function NS_TrouverColonneEntete(ByVal ws As Object, ByVal entete As String) As Long
    Dim derniereColonne As Long, c As Long
    Dim valeur As String

    derniereColonne = ws.Cells(1, ws.Columns.Count).End(-4159).Column

    For c = 1 To derniereColonne
        valeur = Trim$(CStr(ws.Cells(1, c).Value))
        If StrComp(valeur, entete, vbTextCompare) = 0 Then
            NS_TrouverColonneEntete = c
            Exit Function
        End If
    Next c
End Function

Private Function NS_LireCellule(ByVal ws As Object, ByVal ligne As Long, ByVal colonne As Long) As String
    Dim v As Variant

    If colonne <= 0 Then Exit Function
    v = ws.Cells(ligne, colonne).Value

    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function

    NS_LireCellule = CStr(v)
End Function

Private Function NS_EstOui(ByVal valeur As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(valeur))
    NS_EstOui = (t = "oui" Or t = "true" Or t = "vrai" Or t = "1")
End Function

Private Function NS_NormaliserTexte(ByVal texte As String) As String
    texte = LCase$(Trim$(texte))

    texte = Replace(texte, "à", "a")
    texte = Replace(texte, "â", "a")
    texte = Replace(texte, "ä", "a")
    texte = Replace(texte, "á", "a")
    texte = Replace(texte, "ç", "c")
    texte = Replace(texte, "é", "e")
    texte = Replace(texte, "è", "e")
    texte = Replace(texte, "ê", "e")
    texte = Replace(texte, "ë", "e")
    texte = Replace(texte, "î", "i")
    texte = Replace(texte, "ï", "i")
    texte = Replace(texte, "ô", "o")
    texte = Replace(texte, "ö", "o")
    texte = Replace(texte, "ù", "u")
    texte = Replace(texte, "û", "u")
    texte = Replace(texte, "ü", "u")
    texte = Replace(texte, "œ", "oe")
    texte = Replace(texte, "æ", "ae")

    texte = Replace(texte, Chr(160), " ")
    texte = Replace(texte, vbTab, " ")
    texte = Replace(texte, vbCr, " ")
    texte = Replace(texte, vbLf, " ")
    texte = Replace(texte, "’", "'")
    texte = Replace(texte, "'", " ")
    texte = Replace(texte, "-", " ")
    texte = Replace(texte, ".", " ")
    texte = Replace(texte, ",", " ")
    texte = Replace(texte, ";", " ")
    texte = Replace(texte, ":", " ")
    texte = Replace(texte, "/", " ")
    texte = Replace(texte, "\", " ")
    texte = Replace(texte, "(", " ")
    texte = Replace(texte, ")", " ")

    Do While InStr(texte, "  ") > 0
        texte = Replace(texte, "  ", " ")
    Loop

    NS_NormaliserTexte = Trim$(texte)
End Function

Private Function NS_NormaliserTelephone(ByVal texte As String) As String
    Dim i As Long, ch As String, resultat As String

    For i = 1 To Len(texte)
        ch = Mid$(texte, i, 1)
        If ch >= "0" And ch <= "9" Then resultat = resultat & ch
    Next i

    NS_NormaliserTelephone = resultat
End Function

Private Function NS_NormaliserCP(ByVal texte As String) As String
    Dim i As Long, ch As String, resultat As String

    For i = 1 To Len(texte)
        ch = Mid$(texte, i, 1)
        If ch >= "0" And ch <= "9" Then resultat = resultat & ch
    Next i

    NS_NormaliserCP = resultat
End Function
