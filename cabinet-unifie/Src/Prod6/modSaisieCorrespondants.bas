Attribute VB_Name = "modSaisieCorrespondants"
Option Explicit

'===============================================================================
' MODULE : modSaisieCorrespondants
' VERSION : STRUCT-9A — enregistrement sécurisé
'
' Cette version remplace intégralement STRUCT-5C.
'
' Fonctions conservées de STRUCT-7A :
' - entrée "< Nouvelle structure... >" dans la liste Structure ;
' - ouverture de frmNouvelleStructure ;
' - rechargement automatique de la feuille Structures après création ;
' - sélection automatique de la structure nouvellement créée ;
' - retour à la structure précédente si la création est annulée.
'
' Fonctions déjà validées conservées :
' - formulaire dynamique Nouveau correspondant ;
' - lecture Structures / Listes / Generalistes / Specialistes ;
' - activation des types d'examens pour les spécialistes ;
' - contrôle des doublons correspondants ;
' - déclenchement par Tab ;
' - contrôle des doublons correspondants ;
'
' Fonctions conservées de STRUCT-8A :
' - bouton Aperçu de la fiche ;
' - alimentation temporaire de Saisie_Correspondants ;
' - calcul par Excel des champs automatiques ;
' - affichage du résultat calculé ;
' - fermeture du classeur SANS SAUVEGARDER.
'
' Dépendances :
' - CHEMIN_BASE_CORRESPONDANTS
' - frmNouveauCorrespondant
' - frmNouvelleStructure
' - modSaisieStructures (STRUCT-6E)
' - gNSDernierIDStructure
' - gNSDernierNomStructure
' - clsNCTextBox
' - clsNCCombo
' - clsNCBouton
'===============================================================================

Private mHandlersBoutons As Collection
Private mHandlersTextes As Collection
Private mHandlersCombos As Collection

Private mStructures As Object
Private mCorrespondants As Collection

Private mDernierIDStructureSelectionne As String
Private mGestionStructureEnCours As Boolean

Private Const XL_UP As Long = -4162
Private Const ID_NOUVELLE_STRUCTURE As String = "__NOUVELLE_STRUCTURE__"

'-------------------------------------------------------------------------------
' MACRO PUBLIQUE DE TEST
'-------------------------------------------------------------------------------
Public Sub NC_OuvrirNouveauCorrespondant()

    Dim frm As Object

    On Error GoTo GestionErreur

    On Error Resume Next
    Unload frmNouveauCorrespondant
    On Error GoTo GestionErreur

    Set frm = frmNouveauCorrespondant

    NC_InitialiserMemoire
    NC_ConstruireInterface frm
    NC_ChargerDonneesExcel frm
    NC_MAJModeCorrespondant frm

    frm.Show

    Exit Sub

GestionErreur:

    MsgBox _
        "Impossible d'ouvrir le formulaire Nouveau correspondant." & vbCrLf & _
        Err.Number & " - " & Err.description, _
        vbExclamation, _
        "Nouveau correspondant"

End Sub

Private Sub NC_InitialiserMemoire()

    Set mHandlersBoutons = New Collection
    Set mHandlersTextes = New Collection
    Set mHandlersCombos = New Collection

    Set mStructures = CreateObject("Scripting.Dictionary")
    Set mCorrespondants = New Collection

    mDernierIDStructureSelectionne = ""
    mGestionStructureEnCours = False

End Sub

'-------------------------------------------------------------------------------
' CONSTRUCTION DYNAMIQUE DU FORMULAIRE
'-------------------------------------------------------------------------------
Private Sub NC_ConstruireInterface(ByVal frm As Object)

    Dim i As Long
    Dim ctl As Object

    On Error Resume Next
    For i = frm.Controls.Count - 1 To 0 Step -1
        frm.Controls.Remove frm.Controls(i).Name
    Next i
    On Error GoTo 0

    With frm
        .Caption = "Nouveau correspondant"
        .Width = 650
        .Height = 650
        .StartUpPosition = 1
    End With

    Set ctl = NC_AjouterLabel( _
        frm, _
        "lblTitreNC", _
        "Nouveau correspondant médical", _
        14, 10, 600, 22, True)
    ctl.Font.Size = 12

    Set ctl = NC_AjouterLabel( _
        frm, _
        "lblEtapeNC", _
        "Sélectionnez une structure existante ou choisissez < Nouvelle structure... >.", _
        14, 34, 600, 30, False)
    ctl.WordWrap = True

    '--------------------------- Colonne gauche -------------------------------

    NC_AjouterLabel frm, "lblType", "Type de correspondant", 14, 76, 135, 16, True
    Set ctl = NC_AjouterCombo( _
        frm, "cmbTypeCorrespondant", 14, 94, 245, 22, "TypeCorrespondant")
    ctl.AddItem "Généraliste"
    ctl.AddItem "Spécialiste"
    ctl.ListIndex = 0

    NC_AjouterLabel frm, "lblSexe", "Sexe", 14, 126, 135, 16, True
    Set ctl = NC_AjouterCombo( _
        frm, "cmbSexe", 14, 144, 115, 22, "Sexe")
    ctl.AddItem "Homme"
    ctl.AddItem "Femme"
    ctl.ListIndex = -1

    NC_AjouterLabel frm, "lblTutoiement", "Relation", 144, 126, 115, 16, True
    Set ctl = NC_AjouterCombo( _
        frm, "cmbTutoiement", 144, 144, 115, 22, "Tutoiement")
    ctl.AddItem "vous"
    ctl.AddItem "tu"
    ctl.ListIndex = 0

    NC_AjouterLabel frm, "lblPrenom", "Prénom", 14, 178, 110, 16, True
    Set ctl = NC_AjouterTexte( _
        frm, "txtPrenom", 14, 196, 245, 22, "Prenom")

    NC_AjouterLabel frm, "lblNom", "Nom", 14, 228, 110, 16, True
    Set ctl = NC_AjouterTexte( _
        frm, "txtNom", 14, 246, 245, 22, "Nom")

    NC_AjouterLabel frm, "lblTelephone", "Téléphone direct / manuel", 14, 278, 190, 16, True
    Set ctl = NC_AjouterTexte( _
        frm, "txtTelephone", 14, 296, 245, 22, "Telephone")

    NC_AjouterLabel frm, "lblAdresse1", "Adresse manuelle (si nécessaire)", 14, 330, 230, 16, True
    Set ctl = NC_AjouterTexte( _
        frm, "txtAdresse1", 14, 348, 245, 22, "Adresse1")

    Set ctl = NC_AjouterTexte( _
        frm, "txtAdresse2", 14, 374, 245, 22, "Adresse2")

    Set ctl = NC_AjouterTexte( _
        frm, "txtCodePostal", 14, 400, 78, 22, "CodePostal")

    Set ctl = NC_AjouterTexte( _
        frm, "txtVille", 98, 400, 161, 22, "Ville")

    '--------------------------- Colonne droite -------------------------------

    NC_AjouterLabel frm, "lblStructure", "Structure", 285, 76, 120, 16, True
    Set ctl = NC_AjouterCombo( _
        frm, "cmbStructure", 285, 94, 330, 22, "Structure")
    ctl.ColumnCount = 3
    ctl.ColumnWidths = "220 pt;95 pt;0 pt"

    Set ctl = NC_AjouterLabel( _
        frm, _
        "lblNbStructures", _
        "Chargement des structures…", _
        285, 120, 330, 16, False)

    NC_AjouterLabel frm, "lblAdresseStructure", "Adresse récupérée automatiquement", 285, 146, 280, 16, True
    Set ctl = NC_AjouterTexte( _
        frm, "txtStructureAdresse", 285, 164, 330, 74, "")
    ctl.MultiLine = True
    ctl.WordWrap = True
    ctl.Locked = True
    ctl.TabStop = False
    ctl.ScrollBars = 2
    ctl.BackColor = &H8000000F

    NC_AjouterLabel frm, "lblTypesExamens", "Type(s) d'examen / avis — spécialistes", 285, 252, 310, 16, True
    Set ctl = frm.Controls.Add("Forms.ListBox.1", "lstTypesExamens", True)
    With ctl
        .Left = 285
        .Top = 270
        .Width = 330
        .Height = 152
        .MultiSelect = 1
        .ColumnCount = 1
        .Enabled = False
    End With

    '--------------------------- Doublons -------------------------------------

    NC_AjouterLabel frm, "lblDoublons", "Contrôle des doublons", 14, 440, 220, 16, True

    Set ctl = NC_AjouterTexte( _
        frm, "txtDoublons", 14, 458, 601, 105, "")
    ctl.MultiLine = True
    ctl.WordWrap = True
    ctl.Locked = True
    ctl.TabStop = False
    ctl.ScrollBars = 2
    ctl.Text = _
        "Saisissez au minimum le nom. Le contrôle se déclenche avec Tab " & _
        "ou avec le bouton Vérifier les doublons."
    ctl.BackColor = &H8000000F

    NC_AjouterBouton frm, "cmdVerifierDoublons", "Vérifier les doublons", 14, 578, 140, 28, "VerifierDoublons"
    NC_AjouterBouton frm, "cmdApercuNC", "Aperçu de la fiche", 170, 578, 140, 28, "Apercu"
    NC_AjouterBouton frm, "cmdEnregistrerNC", "Enregistrer", 326, 578, 140, 28, "Enregistrer"
    NC_AjouterBouton frm, "cmdFermerNC", "Fermer", 482, 578, 130, 28, "Fermer"

    Set ctl = NC_AjouterLabel( _
        frm, _
        "lblPasEnregistrement", _
        "L'enregistrement crée d'abord une sauvegarde Excel et recontrôle les doublons juste avant l'écriture.", _
        14, 612, 600, 24, False)
    ctl.WordWrap = True

End Sub

Private Function NC_AjouterLabel( _
    ByVal frm As Object, _
    ByVal nom As String, _
    ByVal texte As String, _
    ByVal gauche As Single, _
    ByVal haut As Single, _
    ByVal largeur As Single, _
    ByVal hauteur As Single, _
    Optional ByVal gras As Boolean = False) As Object

    Dim ctl As Object

    Set ctl = frm.Controls.Add("Forms.Label.1", nom, True)

    With ctl
        .Caption = texte
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        .Font.Bold = gras
        .Font.Size = 9
    End With

    Set NC_AjouterLabel = ctl

End Function

Private Function NC_AjouterTexte( _
    ByVal frm As Object, _
    ByVal nom As String, _
    ByVal gauche As Single, _
    ByVal haut As Single, _
    ByVal largeur As Single, _
    ByVal hauteur As Single, _
    ByVal role As String) As Object

    Dim ctl As Object
    Dim h As clsNCTextBox

    Set ctl = frm.Controls.Add("Forms.TextBox.1", nom, True)

    With ctl
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        .Font.Size = 9
    End With

    If Trim$(role) <> "" Then

        Set h = New clsNCTextBox
        Set h.ctl = ctl
        h.role = role
        mHandlersTextes.Add h

    End If

    Set NC_AjouterTexte = ctl

End Function

Private Function NC_AjouterCombo( _
    ByVal frm As Object, _
    ByVal nom As String, _
    ByVal gauche As Single, _
    ByVal haut As Single, _
    ByVal largeur As Single, _
    ByVal hauteur As Single, _
    ByVal role As String) As Object

    Dim ctl As Object
    Dim h As clsNCCombo

    Set ctl = frm.Controls.Add("Forms.ComboBox.1", nom, True)

    With ctl
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        .Font.Size = 9
        .Style = 2
    End With

    Set h = New clsNCCombo
    Set h.ctl = ctl
    h.role = role
    mHandlersCombos.Add h

    Set NC_AjouterCombo = ctl

End Function

Private Sub NC_AjouterBouton( _
    ByVal frm As Object, _
    ByVal nom As String, _
    ByVal texte As String, _
    ByVal gauche As Single, _
    ByVal haut As Single, _
    ByVal largeur As Single, _
    ByVal hauteur As Single, _
    ByVal role As String)

    Dim ctl As Object
    Dim h As clsNCBouton

    Set ctl = frm.Controls.Add("Forms.CommandButton.1", nom, True)

    With ctl
        .Caption = texte
        .Left = gauche
        .Top = haut
        .Width = largeur
        .Height = hauteur
        .Font.Size = 9
    End With

    Set h = New clsNCBouton
    Set h.ctl = ctl
    h.role = role
    mHandlersBoutons.Add h

End Sub

'-------------------------------------------------------------------------------
' CHARGEMENT EXCEL
'-------------------------------------------------------------------------------
Private Sub NC_ChargerDonneesExcel(ByVal frm As Object)

    Dim excelApp As Object
    Dim wb As Object

    On Error GoTo GestionErreur

    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    excelApp.AutomationSecurity = 3
    excelApp.EnableEvents = False

    Set wb = excelApp.Workbooks.Open( _
        CHEMIN_BASE_CORRESPONDANTS, _
        0, _
        True)

    NC_ChargerStructuresDepuisClasseur wb, frm
    NC_ChargerTypesDepuisClasseur wb, frm
    NC_ChargerCorrespondantsDepuisClasseur wb

Sortie:

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If

    Set wb = Nothing
    Set excelApp = Nothing

    On Error GoTo 0
    Exit Sub

GestionErreur:

    MsgBox _
        "Impossible de lire la base des correspondants." & vbCrLf & _
        CHEMIN_BASE_CORRESPONDANTS & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.description, _
        vbExclamation, _
        "Nouveau correspondant"

    Resume Sortie

End Sub

Private Sub NC_ChargerStructuresDepuisClasseur( _
    ByVal wb As Object, _
    ByVal frm As Object)

    Dim ws As Object
    Dim cmb As Object

    Dim colID As Long
    Dim colNom As Long
    Dim colAdresse1 As Long
    Dim colAdresse2 As Long
    Dim colCP As Long
    Dim colVille As Long
    Dim colTel As Long
    Dim colActif As Long

    Dim derniereLigne As Long
    Dim i As Long
    Dim indexListe As Long

    Dim idStructure As String
    Dim nomStructure As String
    Dim ville As String
    Dim actif As String
    Dim record As Variant

    Set ws = wb.Worksheets("Structures")
    Set cmb = frm.Controls("cmbStructure")

    colID = NC_TrouverColonneEntete(ws, "ID_Structure")
    colNom = NC_TrouverColonneEntete(ws, "NomStructure")
    colAdresse1 = NC_TrouverColonneEntete(ws, "Adresse1")
    colAdresse2 = NC_TrouverColonneEntete(ws, "Adresse2")
    colCP = NC_TrouverColonneEntete(ws, "CodePostal")
    colVille = NC_TrouverColonneEntete(ws, "Ville")
    colTel = NC_TrouverColonneEntete(ws, "Telephone")
    colActif = NC_TrouverColonneEntete(ws, "Actif")

    If colID = 0 Or colNom = 0 Or colActif = 0 Then

        Err.Raise _
            vbObjectError + 510, _
            , _
            "Colonnes obligatoires absentes de Structures."

    End If

    Set mStructures = CreateObject("Scripting.Dictionary")

    mGestionStructureEnCours = True

    cmb.Clear

    cmb.AddItem "<Sans structure>"
    cmb.List(0, 1) = ""
    cmb.List(0, 2) = ""

    derniereLigne = _
        ws.Cells(ws.Rows.Count, colID).End(XL_UP).Row

    For i = 2 To derniereLigne

        actif = _
            NC_LireCellule(ws, i, colActif)

        If NC_EstOui(actif) Then

            idStructure = _
                Trim$(NC_LireCellule(ws, i, colID))

            nomStructure = _
                Trim$(NC_LireCellule(ws, i, colNom))

            ville = _
                Trim$(NC_LireCellule(ws, i, colVille))

            If idStructure <> "" _
            And nomStructure <> "" Then

                record = Array( _
                    nomStructure, _
                    NC_LireCellule(ws, i, colAdresse1), _
                    NC_LireCellule(ws, i, colAdresse2), _
                    NC_LireCellule(ws, i, colCP), _
                    ville, _
                    NC_LireCellule(ws, i, colTel))

                If Not mStructures.Exists(idStructure) Then
                    mStructures.Add idStructure, record
                End If

                indexListe = cmb.ListCount
                cmb.AddItem nomStructure
                cmb.List(indexListe, 1) = ville
                cmb.List(indexListe, 2) = idStructure

            End If

        End If

    Next i

    'Entrée spéciale toujours placée à la fin.
    indexListe = cmb.ListCount
    cmb.AddItem "< Nouvelle structure... >"
    cmb.List(indexListe, 1) = ""
    cmb.List(indexListe, 2) = ID_NOUVELLE_STRUCTURE

    cmb.ListIndex = 0
    mDernierIDStructureSelectionne = ""

    frm.Controls("lblNbStructures").Caption = _
        CStr(mStructures.Count) & _
        " structure(s) active(s) disponible(s)."

    mGestionStructureEnCours = False

End Sub

Private Sub NC_ChargerTypesDepuisClasseur( _
    ByVal wb As Object, _
    ByVal frm As Object)

    Dim ws As Object
    Dim lst As Object

    Dim colListe As Long
    Dim colValeur As Long
    Dim derniereLigne As Long
    Dim i As Long

    Dim nomListe As String
    Dim valeur As String

    Set ws = wb.Worksheets("Listes")
    Set lst = frm.Controls("lstTypesExamens")

    colListe = NC_TrouverColonneEntete(ws, "Liste")
    colValeur = NC_TrouverColonneEntete(ws, "Valeur")

    If colListe = 0 Or colValeur = 0 Then

        Err.Raise _
            vbObjectError + 511, _
            , _
            "Colonnes obligatoires absentes de Listes."

    End If

    lst.Clear

    derniereLigne = _
        ws.Cells(ws.Rows.Count, colListe).End(XL_UP).Row

    For i = 2 To derniereLigne

        nomListe = _
            Trim$(NC_LireCellule(ws, i, colListe))

        If StrComp( _
            nomListe, _
            "TypeExamen", _
            vbTextCompare) = 0 Then

            valeur = _
                Trim$(NC_LireCellule(ws, i, colValeur))

            If valeur <> "" _
            And StrComp(valeur, "A_COMPLETER", vbTextCompare) <> 0 Then
                lst.AddItem valeur
            End If

        End If

    Next i

End Sub

Private Sub NC_ChargerCorrespondantsDepuisClasseur(ByVal wb As Object)

    Set mCorrespondants = New Collection

    NC_ChargerCorrespondantsFeuille _
        wb.Worksheets("Generalistes"), _
        "Généraliste", _
        "NomAffichage"

    NC_ChargerCorrespondantsFeuille _
        wb.Worksheets("Specialistes"), _
        "Spécialiste", _
        "NomDestinataire"

End Sub

Private Sub NC_ChargerCorrespondantsFeuille( _
    ByVal ws As Object, _
    ByVal typeCorrespondant As String, _
    ByVal enteteAffichage As String)

    Dim colID As Long
    Dim colNom As Long
    Dim colPrenom As Long
    Dim colAffichage As Long
    Dim colStructure As Long
    Dim colAdresse1 As Long
    Dim colVille As Long
    Dim colTelephone As Long
    Dim colIDStructure As Long
    Dim colActif As Long

    Dim derniereLigne As Long
    Dim i As Long

    Dim d As Object
    Dim nom As String
    Dim affichage As String
    Dim actif As String

    colID = NC_TrouverColonneEntete(ws, "ID")
    colNom = NC_TrouverColonneEntete(ws, "Nom")
    colPrenom = NC_TrouverColonneEntete(ws, "PrenomOuInitiale")
    colAffichage = NC_TrouverColonneEntete(ws, enteteAffichage)
    colStructure = NC_TrouverColonneEntete(ws, "Structure")
    colAdresse1 = NC_TrouverColonneEntete(ws, "Adresse1")
    colVille = NC_TrouverColonneEntete(ws, "Ville")
    colTelephone = NC_TrouverColonneEntete(ws, "Telephone")
    colIDStructure = NC_TrouverColonneEntete(ws, "ID_Structure")
    colActif = NC_TrouverColonneEntete(ws, "Actif")

    If colID = 0 Or colNom = 0 Or colAffichage = 0 Then
        Exit Sub
    End If

    derniereLigne = _
        ws.Cells(ws.Rows.Count, colID).End(XL_UP).Row

    For i = 2 To derniereLigne

        If colActif > 0 Then

            actif = _
                NC_LireCellule(ws, i, colActif)

            If Not NC_EstOui(actif) Then
                GoTo LigneSuivante
            End If

        End If

        nom = _
            Trim$(NC_LireCellule(ws, i, colNom))

        affichage = _
            Trim$(NC_LireCellule(ws, i, colAffichage))

        If nom = "" And affichage = "" Then
            GoTo LigneSuivante
        End If

        Set d = CreateObject("Scripting.Dictionary")

        d.Add "Type", typeCorrespondant
        d.Add "ID", NC_LireCellule(ws, i, colID)
        d.Add "Nom", nom
        d.Add "Prenom", NC_LireCellule(ws, i, colPrenom)
        d.Add "Affichage", affichage
        d.Add "Structure", NC_LireCellule(ws, i, colStructure)
        d.Add "Adresse1", NC_LireCellule(ws, i, colAdresse1)
        d.Add "Ville", NC_LireCellule(ws, i, colVille)
        d.Add "Telephone", NC_LireCellule(ws, i, colTelephone)
        d.Add "ID_Structure", NC_LireCellule(ws, i, colIDStructure)

        mCorrespondants.Add d

LigneSuivante:

    Next i

End Sub

'-------------------------------------------------------------------------------
' NOUVELLE STRUCTURE DEPUIS LE FORMULAIRE CORRESPONDANT
'-------------------------------------------------------------------------------
Private Sub NC_GererNouvelleStructure()

    Dim frm As Object
    Dim idAvant As String
    Dim idCree As String

    If mGestionStructureEnCours Then Exit Sub

    Set frm = frmNouveauCorrespondant

    idAvant = mDernierIDStructureSelectionne

    'Le module Structures remet ces variables à vide à chaque ouverture.
    NS_OuvrirNouvelleStructure

    idCree = Trim$(gNSDernierIDStructure)

    If idCree <> "" Then

        NC_RechargerStructuresEtSelectionner _
            frm, _
            idCree

    Else

        'Annulation : revenir à la sélection précédente.
        NC_SelectionnerStructureParID _
            frm, _
            idAvant

    End If

End Sub

Private Sub NC_RechargerStructuresEtSelectionner( _
    ByVal frm As Object, _
    ByVal idASelectionner As String)

    Dim excelApp As Object
    Dim wb As Object

    On Error GoTo GestionErreur

    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    excelApp.AutomationSecurity = 3
    excelApp.EnableEvents = False

    Set wb = excelApp.Workbooks.Open( _
        CHEMIN_BASE_CORRESPONDANTS, _
        0, _
        True)

    NC_ChargerStructuresDepuisClasseur _
        wb, _
        frm

Sortie:

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If

    Set wb = Nothing
    Set excelApp = Nothing

    On Error GoTo 0

    NC_SelectionnerStructureParID _
        frm, _
        idASelectionner

    Exit Sub

GestionErreur:

    MsgBox _
        "La structure a été créée, mais la liste n'a pas pu être rechargée." & _
        vbCrLf & _
        Err.Number & " - " & Err.description, _
        vbExclamation, _
        "Nouveau correspondant"

    Resume Sortie

End Sub

Private Sub NC_SelectionnerStructureParID( _
    ByVal frm As Object, _
    ByVal idStructure As String)

    Dim cmb As Object
    Dim i As Long
    Dim idLigne As String

    Set cmb = frm.Controls("cmbStructure")

    mGestionStructureEnCours = True

    If Trim$(idStructure) = "" Then

        cmb.ListIndex = 0
        mDernierIDStructureSelectionne = ""

    Else

        For i = 0 To cmb.ListCount - 1

            idLigne = _
                Trim$(CStr(cmb.List(i, 2)))

            If StrComp( _
                idLigne, _
                idStructure, _
                vbTextCompare) = 0 Then

                cmb.ListIndex = i
                mDernierIDStructureSelectionne = idStructure

                Exit For

            End If

        Next i

    End If

    mGestionStructureEnCours = False

    NC_AfficherStructureSelectionnee frm

End Sub

'-------------------------------------------------------------------------------
' EVENEMENTS ROUTES PAR LES CLASSES
'-------------------------------------------------------------------------------
Public Sub NC_EvenementBouton(ByVal role As String)

    Select Case role

        Case "VerifierDoublons"
            NC_VerifierDoublonsFormulaire

        Case "Apercu"
            NC_PrevisualiserFicheCorrespondant

        Case "Enregistrer"
            NCE_EnregistrerCorrespondant

        Case "Fermer"
            Unload frmNouveauCorrespondant

    End Select

End Sub

Public Sub NC_EvenementCombo(ByVal role As String)

    Dim frm As Object
    Dim idStructure As String

    Set frm = frmNouveauCorrespondant

    Select Case role

        Case "TypeCorrespondant"

            NC_MAJModeCorrespondant frm

        Case "Structure"

            If mGestionStructureEnCours Then Exit Sub

            idStructure = _
                NC_IDStructureSelectionnee(frm)

            If StrComp( _
                idStructure, _
                ID_NOUVELLE_STRUCTURE, _
                vbTextCompare) = 0 Then

                NC_GererNouvelleStructure
                Exit Sub

            End If

            mDernierIDStructureSelectionne = idStructure

            NC_AfficherStructureSelectionnee frm

            If Trim$(frm.Controls("txtNom").Text) <> "" Then
                NC_VerifierDoublonsFormulaire
            End If

    End Select

End Sub

Public Sub NC_EvenementTexteQuitte(ByVal role As String)

    If role = "Nom" _
    Or role = "Prenom" _
    Or role = "Telephone" Then

        If Trim$(frmNouveauCorrespondant.Controls("txtNom").Text) <> "" Then
            NC_VerifierDoublonsFormulaire
        End If

    End If

End Sub

Private Sub NC_MAJModeCorrespondant(ByVal frm As Object)

    Dim estSpecialiste As Boolean
    Dim lst As Object

    estSpecialiste = _
        (StrComp( _
            Trim$(frm.Controls("cmbTypeCorrespondant").Text), _
            "Spécialiste", _
            vbTextCompare) = 0)

    'Sécurisation : l'événement Change peut se déclencher pendant
    'la construction avant la création de lstTypesExamens.
    On Error Resume Next
    Set lst = frm.Controls("lstTypesExamens")
    On Error GoTo 0

    If lst Is Nothing Then Exit Sub

    lst.Enabled = estSpecialiste

End Sub

Private Sub NC_AfficherStructureSelectionnee(ByVal frm As Object)

    Dim cmb As Object
    Dim idStructure As String
    Dim v As Variant
    Dim texte As String

    Set cmb = frm.Controls("cmbStructure")

    frm.Controls("txtStructureAdresse").Text = ""

    If cmb.ListIndex < 0 Then Exit Sub

    idStructure = _
        Trim$(CStr(cmb.List(cmb.ListIndex, 2)))

    If idStructure = "" Then Exit Sub

    If StrComp( _
        idStructure, _
        ID_NOUVELLE_STRUCTURE, _
        vbTextCompare) = 0 Then
        Exit Sub
    End If

    If Not mStructures.Exists(idStructure) Then Exit Sub

    v = mStructures(idStructure)

    texte = CStr(v(0))

    If Trim$(CStr(v(1))) <> "" Then
        texte = texte & vbCrLf & CStr(v(1))
    End If

    If Trim$(CStr(v(2))) <> "" Then
        texte = texte & vbCrLf & CStr(v(2))
    End If

    If Trim$(CStr(v(3)) & " " & CStr(v(4))) <> "" Then

        texte = _
            texte & vbCrLf & _
            Trim$(CStr(v(3)) & " " & CStr(v(4)))

    End If

    If Trim$(CStr(v(5))) <> "" Then
        texte = texte & vbCrLf & CStr(v(5))
    End If

    frm.Controls("txtStructureAdresse").Text = texte

End Sub

'-------------------------------------------------------------------------------
' CONTROLE DES DOUBLONS CORRESPONDANTS
'-------------------------------------------------------------------------------
Public Sub NC_VerifierDoublonsFormulaire()

    Dim frm As Object
    Dim d As Object

    Dim nomRecherche As String
    Dim prenomRecherche As String
    Dim telephoneRecherche As String
    Dim idStructureRecherche As String

    Dim nomExistant As String
    Dim prenomExistant As String
    Dim affichageExistant As String
    Dim telephoneExistant As String
    Dim idStructureExistant As String

    Dim identite1 As String
    Dim identite2 As String

    Dim score As Long
    Dim certains As String
    Dim probables As String
    Dim nbCertains As Long
    Dim nbProbables As Long
    Dim resultat As String

    Set frm = frmNouveauCorrespondant

    nomRecherche = _
        NC_NormaliserSimple(frm.Controls("txtNom").Text)

    prenomRecherche = _
        NC_NormaliserSimple(frm.Controls("txtPrenom").Text)

    telephoneRecherche = _
        NC_NormaliserTelephone(frm.Controls("txtTelephone").Text)

    idStructureRecherche = _
        NC_IDStructureSelectionnee(frm)

    If StrComp( _
        idStructureRecherche, _
        ID_NOUVELLE_STRUCTURE, _
        vbTextCompare) = 0 Then

        idStructureRecherche = ""

    End If

    If nomRecherche = "" Then

        frm.Controls("txtDoublons").Text = _
            "Saisissez au minimum le nom avant de lancer le contrôle."

        Exit Sub

    End If

    identite1 = _
        Trim$(prenomRecherche & " " & nomRecherche)

    identite2 = _
        Trim$(nomRecherche & " " & prenomRecherche)

    certains = ""
    probables = ""
    nbCertains = 0
    nbProbables = 0

    For Each d In mCorrespondants

        nomExistant = _
            NC_NormaliserSimple(CStr(d("Nom")))

        prenomExistant = _
            NC_NormaliserSimple(CStr(d("Prenom")))

        affichageExistant = _
            NC_NormaliserIdentite(CStr(d("Affichage")))

        telephoneExistant = _
            NC_NormaliserTelephone(CStr(d("Telephone")))

        idStructureExistant = _
            Trim$(CStr(d("ID_Structure")))

        score = 0

        If nomRecherche <> "" _
        And prenomRecherche <> "" _
        And nomExistant = nomRecherche _
        And prenomExistant = prenomRecherche Then

            score = 100

        ElseIf prenomRecherche <> "" _
        And (affichageExistant = identite1 _
             Or affichageExistant = identite2) Then

            score = 100

        ElseIf nomExistant = nomRecherche _
        And telephoneRecherche <> "" _
        And telephoneExistant <> "" _
        And telephoneExistant = telephoneRecherche Then

            score = 100

        ElseIf nomExistant = nomRecherche _
        And idStructureRecherche <> "" _
        And idStructureExistant <> "" _
        And StrComp( _
                idStructureRecherche, _
                idStructureExistant, _
                vbTextCompare) = 0 Then

            score = 100

        ElseIf nomExistant <> "" _
        And nomExistant = nomRecherche Then

            score = 60

        ElseIf Len(nomRecherche) >= 4 _
        And InStr( _
                1, _
                affichageExistant, _
                nomRecherche, _
                vbTextCompare) > 0 Then

            If prenomRecherche = "" _
            Or InStr( _
                    1, _
                    affichageExistant, _
                    prenomRecherche, _
                    vbTextCompare) > 0 Then

                score = 50

            End If

        End If

        If score >= 100 Then

            nbCertains = nbCertains + 1

            certains = _
                certains & _
                NC_LigneCorrespondant(d, "DOUBLON CERTAIN") & _
                vbCrLf

        ElseIf score >= 50 Then

            nbProbables = nbProbables + 1

            probables = _
                probables & _
                NC_LigneCorrespondant(d, "À VÉRIFIER") & _
                vbCrLf

        End If

    Next d

    If nbCertains = 0 And nbProbables = 0 Then

        resultat = _
            "Aucun doublon détecté dans Generalistes et Specialistes " & _
            "pour l'identité actuellement saisie."

    Else

        resultat = _
            "Résultat du contrôle :" & vbCrLf & _
            "Doublon(s) certain(s) : " & nbCertains & vbCrLf & _
            "Correspondance(s) à vérifier : " & nbProbables & _
            vbCrLf & vbCrLf

        If certains <> "" Then
            resultat = resultat & certains
        End If

        If probables <> "" Then

            If certains <> "" Then
                resultat = resultat & vbCrLf
            End If

            resultat = resultat & probables

        End If

        resultat = _
            resultat & vbCrLf & _
            "Aucune nouvelle fiche ne sera enregistrée automatiquement " & _
            "tant que ce contrôle n'aura pas été traité."

    End If

    frm.Controls("txtDoublons").Text = resultat
    frm.Controls("txtDoublons").SelStart = 0

End Sub

Private Function NC_LigneCorrespondant( _
    ByVal d As Object, _
    ByVal niveau As String) As String

    Dim texte As String

    texte = _
        "[" & niveau & "] " & _
        CStr(d("Type")) & " — " & _
        CStr(d("ID")) & " — " & _
        CStr(d("Affichage"))

    If Trim$(CStr(d("Structure"))) <> "" Then
        texte = texte & " — " & CStr(d("Structure"))
    End If

    If Trim$(CStr(d("Ville"))) <> "" Then
        texte = texte & " — " & CStr(d("Ville"))
    End If

    NC_LigneCorrespondant = texte

End Function

Private Function NC_IDStructureSelectionnee(ByVal frm As Object) As String

    Dim cmb As Object

    Set cmb = frm.Controls("cmbStructure")

    NC_IDStructureSelectionnee = ""

    If cmb.ListIndex < 0 Then Exit Function

    NC_IDStructureSelectionnee = _
        Trim$(CStr(cmb.List(cmb.ListIndex, 2)))

End Function


'-------------------------------------------------------------------------------
' APERÇU DE LA FICHE VIA Saisie_Correspondants
'-------------------------------------------------------------------------------
Private Sub NC_PrevisualiserFicheCorrespondant()

    Dim frm As Object

    Dim typeCorrespondant As String
    Dim sexe As String
    Dim prenom As String
    Dim nom As String
    Dim idStructure As String
    Dim adresse1 As String
    Dim adresse2 As String
    Dim codePostal As String
    Dim ville As String
    Dim telephone As String
    Dim relation As String
    Dim typesExamens As String

    Dim excelApp As Object
    Dim wb As Object
    Dim ws As Object

    Dim titre As String
    Dim civilite As String
    Dim nomAffichage As String
    Dim nomDestinataire As String
    Dim nomStructure As String
    Dim adresse1Effective As String
    Dim adresse2Effective As String
    Dim cpEffectif As String
    Dim villeEffective As String
    Dim telephoneEffectif As String
    Dim formuleAppel As String
    Dim formulePolitesse As String
    Dim blocDestinataire As String
    Dim typePrincipal As String
    Dim typesPossibles As String

    Dim apercu As String

    On Error GoTo GestionErreur

    Set frm = frmNouveauCorrespondant

    typeCorrespondant = _
        Trim$(frm.Controls("cmbTypeCorrespondant").Text)

    sexe = _
        Trim$(frm.Controls("cmbSexe").Text)

    prenom = _
        Trim$(frm.Controls("txtPrenom").Text)

    nom = _
        Trim$(frm.Controls("txtNom").Text)

    idStructure = _
        NC_IDStructureSelectionnee(frm)

    adresse1 = _
        Trim$(frm.Controls("txtAdresse1").Text)

    adresse2 = _
        Trim$(frm.Controls("txtAdresse2").Text)

    codePostal = _
        Trim$(frm.Controls("txtCodePostal").Text)

    ville = _
        Trim$(frm.Controls("txtVille").Text)

    telephone = _
        Trim$(frm.Controls("txtTelephone").Text)

    relation = _
        Trim$(frm.Controls("cmbTutoiement").Text)

    typesExamens = _
        NC_TypesExamensSelectionnes(frm)

    If typeCorrespondant = "" Then

        MsgBox _
            "Choisissez Généraliste ou Spécialiste.", _
            vbExclamation, _
            "Aperçu correspondant"

        Exit Sub

    End If

    If sexe = "" Then

        MsgBox _
            "Le sexe du correspondant est nécessaire pour calculer " & _
            "la civilité et la formule d'appel.", _
            vbExclamation, _
            "Aperçu correspondant"

        frm.Controls("cmbSexe").SetFocus
        Exit Sub

    End If

    If prenom = "" Then

        MsgBox _
            "Le prénom est obligatoire pour une nouvelle fiche.", _
            vbExclamation, _
            "Aperçu correspondant"

        frm.Controls("txtPrenom").SetFocus
        Exit Sub

    End If

    If nom = "" Then

        MsgBox _
            "Le nom est obligatoire pour une nouvelle fiche.", _
            vbExclamation, _
            "Aperçu correspondant"

        frm.Controls("txtNom").SetFocus
        Exit Sub

    End If

    If StrComp( _
        idStructure, _
        ID_NOUVELLE_STRUCTURE, _
        vbTextCompare) = 0 Then

        MsgBox _
            "Créez ou sélectionnez d'abord la structure.", _
            vbExclamation, _
            "Aperçu correspondant"

        Exit Sub

    End If

    If idStructure = "" Then

        If adresse1 = "" _
        Or codePostal = "" _
        Or ville = "" Then

            MsgBox _
                "Sans structure sélectionnée, renseignez au minimum " & _
                "l'adresse, le code postal et la ville.", _
                vbExclamation, _
                "Aperçu correspondant"

            Exit Sub

        End If

    End If

    If StrComp( _
        typeCorrespondant, _
        "Spécialiste", _
        vbTextCompare) = 0 Then

        If typesExamens = "" Then

            MsgBox _
                "Pour un spécialiste, sélectionnez au moins un type " & _
                "d'examen ou d'avis.", _
                vbExclamation, _
                "Aperçu correspondant"

            Exit Sub

        End If

    End If

    'Le doublon certain bloque déjà au stade de l'aperçu.
    If NC_ExisteDoublonCertainCourant(frm) Then

        NC_VerifierDoublonsFormulaire

        MsgBox _
            "Aperçu interrompu : un doublon certain existe déjà.", _
            vbExclamation, _
            "Doublon correspondant"

        Exit Sub

    End If

    Set excelApp = CreateObject("Excel.Application")

    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    excelApp.AutomationSecurity = 3
    excelApp.EnableEvents = False

    'Lecture seule volontaire : les cellules peuvent être modifiées
    'en mémoire pour le calcul, mais le fichier ne pourra pas être sauvegardé.
    Set wb = excelApp.Workbooks.Open( _
        CHEMIN_BASE_CORRESPONDANTS, _
        0, _
        True)

    Set ws = wb.Worksheets("Saisie_Correspondants")

    'Champs saisis par le formulaire, définis dans la feuille STRUCT-3.
    modFichiers.EcrireCelluleTexte ws.Range("B4"), typeCorrespondant
    modFichiers.EcrireCelluleTexte ws.Range("B5"), sexe
    modFichiers.EcrireCelluleTexte ws.Range("B6"), prenom
    modFichiers.EcrireCelluleTexte ws.Range("B7"), nom
    modFichiers.EcrireCelluleTexte ws.Range("B8"), idStructure
    modFichiers.EcrireCelluleTexte ws.Range("B9"), adresse1
    modFichiers.EcrireCelluleTexte ws.Range("B10"), adresse2
    modFichiers.EcrireCelluleTexte ws.Range("B11"), codePostal
    modFichiers.EcrireCelluleTexte ws.Range("B12"), ville
    modFichiers.EcrireCelluleTexte ws.Range("B13"), telephone
    modFichiers.EcrireCelluleTexte ws.Range("B14"), relation
    modFichiers.EcrireCelluleTexte ws.Range("B15"), typesExamens
    ws.Range("B16").Value = ""

    modBaseCorrespondants.PROD_ActualiserPlagesStructures wb
    excelApp.CalculateFull

    titre = NC_LireValeurRange(ws.Range("E4"))
    civilite = NC_LireValeurRange(ws.Range("E5"))
    nomAffichage = NC_LireValeurRange(ws.Range("E6"))
    nomDestinataire = NC_LireValeurRange(ws.Range("E7"))
    nomStructure = NC_LireValeurRange(ws.Range("E8"))
    adresse1Effective = NC_LireValeurRange(ws.Range("E9"))
    adresse2Effective = NC_LireValeurRange(ws.Range("E10"))
    cpEffectif = NC_LireValeurRange(ws.Range("E11"))
    villeEffective = NC_LireValeurRange(ws.Range("E12"))
    telephoneEffectif = NC_LireValeurRange(ws.Range("E13"))
    formuleAppel = NC_LireValeurRange(ws.Range("E14"))
    formulePolitesse = NC_LireValeurRange(ws.Range("E15"))
    blocDestinataire = NC_LireValeurRange(ws.Range("E17"))
    typePrincipal = NC_LireValeurRange(ws.Range("E22"))
    typesPossibles = NC_LireValeurRange(ws.Range("E23"))

    apercu = _
        "APERÇU — AUCUN ENREGISTREMENT" & vbCrLf & vbCrLf & _
        "Type : " & typeCorrespondant & vbCrLf & _
        "Titre : " & titre & vbCrLf & _
        "Civilité : " & civilite & vbCrLf & _
        "Nom affiché : " & nomAffichage & vbCrLf

    If nomStructure <> "" Then
        apercu = apercu & _
            "Structure : " & nomStructure & vbCrLf
    Else
        apercu = apercu & _
            "Structure : aucune" & vbCrLf
    End If

    apercu = apercu & _
        "Adresse : " & adresse1Effective & vbCrLf

    If adresse2Effective <> "" Then
        apercu = apercu & _
            "Complément : " & adresse2Effective & vbCrLf
    End If

    apercu = apercu & _
        "Code postal / ville : " & _
        Trim$(cpEffectif & " " & villeEffective) & vbCrLf

    If telephoneEffectif <> "" Then
        apercu = apercu & _
            "Téléphone : " & telephoneEffectif & vbCrLf
    End If

    apercu = apercu & vbCrLf & _
        "Formule d'appel : " & formuleAppel & vbCrLf & _
        "Formule de politesse : " & formulePolitesse & vbCrLf

    If StrComp( _
        typeCorrespondant, _
        "Spécialiste", _
        vbTextCompare) = 0 Then

        apercu = apercu & vbCrLf & _
            "Type principal : " & typePrincipal & vbCrLf & _
            "Types possibles : " & typesPossibles & vbCrLf

    End If

    apercu = apercu & vbCrLf & _
        "Bloc destinataire calculé :" & vbCrLf & _
        "----------------------------------------" & vbCrLf & _
        blocDestinataire & vbCrLf & _
        "----------------------------------------"

    MsgBox _
        apercu, _
        vbInformation, _
        "Aperçu fiche correspondant"

Sortie:

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If

    Set ws = Nothing
    Set wb = Nothing
    Set excelApp = Nothing

    On Error GoTo 0
    Exit Sub

GestionErreur:

    Dim numeroErreur As Long
    Dim descriptionErreur As String

    numeroErreur = Err.Number
    descriptionErreur = Err.description

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If

    Set ws = Nothing
    Set wb = Nothing
    Set excelApp = Nothing

    On Error GoTo 0

    MsgBox _
        "Impossible de calculer l'aperçu." & vbCrLf & _
        CStr(numeroErreur) & " - " & descriptionErreur, _
        vbExclamation, _
        "Aperçu correspondant"

End Sub

Private Function NC_TypesExamensSelectionnes( _
    ByVal frm As Object) As String

    Dim lst As Object
    Dim i As Long
    Dim resultat As String

    Set lst = frm.Controls("lstTypesExamens")

    resultat = ""

    For i = 0 To lst.ListCount - 1

        If lst.Selected(i) Then

            If Trim$(CStr(lst.List(i))) <> "" Then

                If resultat <> "" Then
                    resultat = resultat & ";"
                End If

                resultat = _
                    resultat & Trim$(CStr(lst.List(i)))

            End If

        End If

    Next i

    NC_TypesExamensSelectionnes = resultat

End Function

Private Function NC_ExisteDoublonCertainCourant( _
    ByVal frm As Object) As Boolean

    Dim d As Object

    Dim nomRecherche As String
    Dim prenomRecherche As String
    Dim telephoneRecherche As String
    Dim idStructureRecherche As String

    Dim nomExistant As String
    Dim prenomExistant As String
    Dim affichageExistant As String
    Dim telephoneExistant As String
    Dim idStructureExistant As String

    Dim identite1 As String
    Dim identite2 As String

    NC_ExisteDoublonCertainCourant = False

    nomRecherche = _
        NC_NormaliserSimple(frm.Controls("txtNom").Text)

    prenomRecherche = _
        NC_NormaliserSimple(frm.Controls("txtPrenom").Text)

    telephoneRecherche = _
        NC_NormaliserTelephone(frm.Controls("txtTelephone").Text)

    idStructureRecherche = _
        NC_IDStructureSelectionnee(frm)

    If StrComp( _
        idStructureRecherche, _
        ID_NOUVELLE_STRUCTURE, _
        vbTextCompare) = 0 Then

        idStructureRecherche = ""

    End If

    If nomRecherche = "" Then Exit Function

    identite1 = _
        Trim$(prenomRecherche & " " & nomRecherche)

    identite2 = _
        Trim$(nomRecherche & " " & prenomRecherche)

    For Each d In mCorrespondants

        nomExistant = _
            NC_NormaliserSimple(CStr(d("Nom")))

        prenomExistant = _
            NC_NormaliserSimple(CStr(d("Prenom")))

        affichageExistant = _
            NC_NormaliserIdentite(CStr(d("Affichage")))

        telephoneExistant = _
            NC_NormaliserTelephone(CStr(d("Telephone")))

        idStructureExistant = _
            Trim$(CStr(d("ID_Structure")))

        If nomRecherche <> "" _
        And prenomRecherche <> "" _
        And nomExistant = nomRecherche _
        And prenomExistant = prenomRecherche Then

            NC_ExisteDoublonCertainCourant = True
            Exit Function

        End If

        If prenomRecherche <> "" _
        And (affichageExistant = identite1 _
             Or affichageExistant = identite2) Then

            NC_ExisteDoublonCertainCourant = True
            Exit Function

        End If

        If nomExistant = nomRecherche _
        And telephoneRecherche <> "" _
        And telephoneExistant <> "" _
        And telephoneExistant = telephoneRecherche Then

            NC_ExisteDoublonCertainCourant = True
            Exit Function

        End If

        If nomExistant = nomRecherche _
        And idStructureRecherche <> "" _
        And idStructureExistant <> "" _
        And StrComp( _
                idStructureRecherche, _
                idStructureExistant, _
                vbTextCompare) = 0 Then

            NC_ExisteDoublonCertainCourant = True
            Exit Function

        End If

    Next d

End Function

Private Function NC_LireValeurRange(ByVal rng As Object) As String

    Dim v As Variant

    NC_LireValeurRange = ""

    v = rng.Value

    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function

    NC_LireValeurRange = CStr(v)

End Function

'-------------------------------------------------------------------------------
' UTILITAIRES EXCEL
'-------------------------------------------------------------------------------
Private Function NC_TrouverColonneEntete( _
    ByVal ws As Object, _
    ByVal entete As String) As Long

    Dim derniereColonne As Long
    Dim c As Long
    Dim valeur As String

    NC_TrouverColonneEntete = 0

    derniereColonne = _
        ws.Cells(1, ws.Columns.Count).End(-4159).Column

    For c = 1 To derniereColonne

        valeur = _
            Trim$(CStr(ws.Cells(1, c).Value))

        If StrComp( _
            valeur, _
            entete, _
            vbTextCompare) = 0 Then

            NC_TrouverColonneEntete = c
            Exit Function

        End If

    Next c

End Function

Private Function NC_LireCellule( _
    ByVal ws As Object, _
    ByVal ligne As Long, _
    ByVal colonne As Long) As String

    Dim v As Variant

    NC_LireCellule = ""

    If colonne <= 0 Then Exit Function

    v = ws.Cells(ligne, colonne).Value

    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function

    NC_LireCellule = CStr(v)

End Function

Private Function NC_EstOui(ByVal valeur As String) As Boolean

    Dim t As String

    t = LCase$(Trim$(valeur))

    NC_EstOui = _
        (t = "oui" _
         Or t = "true" _
         Or t = "vrai" _
         Or t = "1")

End Function

'-------------------------------------------------------------------------------
' NORMALISATION POUR DOUBLONS
'-------------------------------------------------------------------------------
Private Function NC_NormaliserSimple(ByVal texte As String) As String

    texte = LCase$(Trim$(texte))

    texte = Replace(texte, "à", "a")
    texte = Replace(texte, "â", "a")
    texte = Replace(texte, "ä", "a")
    texte = Replace(texte, "á", "a")
    texte = Replace(texte, "ã", "a")

    texte = Replace(texte, "ç", "c")

    texte = Replace(texte, "é", "e")
    texte = Replace(texte, "è", "e")
    texte = Replace(texte, "ê", "e")
    texte = Replace(texte, "ë", "e")

    texte = Replace(texte, "î", "i")
    texte = Replace(texte, "ï", "i")
    texte = Replace(texte, "í", "i")

    texte = Replace(texte, "ô", "o")
    texte = Replace(texte, "ö", "o")
    texte = Replace(texte, "ó", "o")

    texte = Replace(texte, "ù", "u")
    texte = Replace(texte, "û", "u")
    texte = Replace(texte, "ü", "u")
    texte = Replace(texte, "ú", "u")

    texte = Replace(texte, "ÿ", "y")

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

    NC_NormaliserSimple = Trim$(texte)

End Function

Private Function NC_NormaliserIdentite(ByVal texte As String) As String

    Dim t As String

    t = " " & NC_NormaliserSimple(texte) & " "

    t = Replace(t, " monsieur ", " ")
    t = Replace(t, " madame ", " ")
    t = Replace(t, " docteur ", " ")
    t = Replace(t, " dr ", " ")

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    NC_NormaliserIdentite = Trim$(t)

End Function

Private Function NC_NormaliserTelephone(ByVal texte As String) As String

    Dim i As Long
    Dim ch As String
    Dim resultat As String

    resultat = ""

    For i = 1 To Len(texte)

        ch = Mid$(texte, i, 1)

        If ch >= "0" And ch <= "9" Then
            resultat = resultat & ch
        End If

    Next i

    NC_NormaliserTelephone = resultat

End Function
