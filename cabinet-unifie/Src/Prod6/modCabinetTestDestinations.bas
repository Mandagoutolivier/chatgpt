Attribute VB_Name = "modCabinetTestDestinations"
Option Explicit

'===============================================================================
' MODULE : modCabinetTestDestinations
' VERSION : CABINET TEST v1
'
' Objet :
' - rechercher dans Excel un destinataire à partir de CLE_DESTINATION ;
' - reconstruire un record compatible avec ChampDestination ;
' - distinguer les demandes résolues des demandes à compléter.
'
' Dépendances :
' - CHEMIN_BASE_CORRESPONDANTS
' - CT_ConstruireCollectionDemandes
' - ChampDestination
'===============================================================================

Public Function CT_DestinationParCle( _
    ByVal cleRecherche As String) As String

    Const XL_UP As Long = -4162

    Dim excelApp As Object
    Dim classeur As Object
    Dim feuille As Object

    Dim colIDLigne As Long
    Dim colTypeExamen As Long
    Dim colNom As Long
    Dim colStructure As Long
    Dim colCodePostal As Long
    Dim colVille As Long
    Dim colBloc As Long
    Dim colFormuleAppel As Long
    Dim colPriorite As Long
    Dim colActif As Long
    Dim colCle As Long
    Dim colFormulePolitesse As Long
    Dim colTutoiement As Long

    Dim derniereLigne As Long
    Dim meilleureLigne As Long
    Dim identiteCible As String, identiteLigne As String, colAValider As Long
    Dim meilleurePriorite As Long
    Dim prioriteLigne As Long
    Dim i As Long

    Dim cleLigne As String
    Dim actifLigne As String

    CT_DestinationParCle = ""

    cleRecherche = UCase$(Trim$(cleRecherche))

    If cleRecherche = "" Then Exit Function
    If cleRecherche = "A_COMPLETER" Then Exit Function

    On Error GoTo GestionErreur

    Set excelApp = CreateObject("Excel.Application")

    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    excelApp.AutomationSecurity = 3
    excelApp.EnableEvents = False

    Set classeur = excelApp.Workbooks.Open( _
        CHEMIN_BASE_CORRESPONDANTS, _
        0, _
        True)

    Set feuille = classeur.Worksheets("Specialistes_ParType")

    colIDLigne = CT_TrouverColonneEntete(feuille, "ID_Ligne")
    colTypeExamen = CT_TrouverColonneEntete(feuille, "TypeExamen")
    colAValider = CT_TrouverColonneEntete(feuille, "AValider")
    colNom = CT_TrouverColonneEntete(feuille, "NomDestinataire")
    colStructure = CT_TrouverColonneEntete(feuille, "Structure")
    colCodePostal = CT_TrouverColonneEntete(feuille, "CodePostal")
    colVille = CT_TrouverColonneEntete(feuille, "Ville")
    colBloc = CT_TrouverColonneEntete(feuille, "BlocDestinataireComplet")
    colFormuleAppel = CT_TrouverColonneEntete(feuille, "FormuleAppel")
    colPriorite = CT_TrouverColonneEntete(feuille, "Priorite")
    colActif = CT_TrouverColonneEntete(feuille, "Actif")
    colCle = CT_TrouverColonneEntete(feuille, "CleRegroupement")
    colFormulePolitesse = CT_TrouverColonneEntete( _
        feuille, _
        "FormulePolitesse")
    colTutoiement = CT_TrouverColonneEntete( _
        feuille, _
        "TutoiementVouvoiement")

    If colCle = 0 _
    Or colNom = 0 _
    Or colBloc = 0 _
    Or colActif = 0 Then

        MsgBox _
            "Une colonne indispensable est absente de la feuille " & _
            "Specialistes_ParType." & vbCrLf & vbCrLf & _
            "Colonnes requises : CleRegroupement, NomDestinataire, " & _
            "BlocDestinataireComplet et Actif.", _
            vbExclamation, _
            "Cabinet Test"

        GoTo Sortie

    End If

    derniereLigne = feuille.Cells( _
        feuille.Rows.Count, _
        colCle).End(XL_UP).Row

    meilleureLigne = 0
    meilleurePriorite = 2147483647

    For i = 2 To derniereLigne

        cleLigne = UCase$(Trim$( _
            CT_LireValeurCellule( _
                feuille, _
                i, _
                colCle)))

        actifLigne = CT_LireValeurCellule( _
            feuille, _
            i, _
            colActif)

        If StrComp( _
            cleLigne, _
            cleRecherche, _
            vbTextCompare) = 0 _
        And CT_EstValeurActive(actifLigne) Then

            If colAValider > 0 Then
                If CT_EstValeurActive(CT_LireValeurCellule(feuille, i, colAValider)) Then GoTo LigneSuivanteCle
            End If
            identiteLigne = CT_LireValeurCellule(feuille, i, colNom) & "|" & CT_LireValeurCellule(feuille, i, colBloc)
            If Len(identiteCible) > 0 And StrComp(identiteCible, identiteLigne, vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 966, , "Cle de destination ambigue dans la base : " & cleRecherche
            End If
            identiteCible = identiteLigne
            prioriteLigne = CT_LirePriorite( _
                CT_LireValeurCellule( _
                    feuille, _
                    i, _
                    colPriorite))

            If meilleureLigne = 0 _
            Or prioriteLigne < meilleurePriorite Then

                meilleureLigne = i
                meilleurePriorite = prioriteLigne

            End If

        End If

LigneSuivanteCle:
    Next i

    If meilleureLigne = 0 Then GoTo Sortie

    CT_DestinationParCle = CT_ConstruireRecordDestination( _
        CT_LireValeurCellule( _
            feuille, meilleureLigne, colIDLigne), _
        CT_LireValeurCellule( _
            feuille, meilleureLigne, colTypeExamen), _
        CT_LireValeurCellule( _
            feuille, meilleureLigne, colNom), _
        CT_LireValeurCellule( _
            feuille, meilleureLigne, colStructure), _
        CT_LireValeurCellule( _
            feuille, meilleureLigne, colBloc), _
        CT_LireValeurCellule( _
            feuille, meilleureLigne, colFormuleAppel), _
        CT_LireValeurCellule( _
            feuille, meilleureLigne, colTutoiement), _
        CStr(meilleurePriorite), _
        CT_LireValeurCellule( _
            feuille, meilleureLigne, colCodePostal), _
        CT_LireValeurCellule( _
            feuille, meilleureLigne, colVille), _
        CT_LireValeurCellule( _
            feuille, meilleureLigne, colCle), _
        CT_LireValeurCellule( _
            feuille, meilleureLigne, colFormulePolitesse))

Sortie:

    On Error Resume Next

    If Not classeur Is Nothing Then
        classeur.Close False
    End If

    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If

    Set feuille = Nothing
    Set classeur = Nothing
    Set excelApp = Nothing

    On Error GoTo 0
    Exit Function

GestionErreur:

    MsgBox _
        "Erreur pendant la recherche de la clé :" & vbCrLf & _
        cleRecherche & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.description, _
        vbExclamation, _
        "Cabinet Test"

    Resume Sortie

End Function

Private Function CT_TrouverColonneEntete( _
    ByVal feuille As Object, _
    ByVal enteteRecherchee As String) As Long

    Const XL_TO_LEFT As Long = -4159

    Dim derniereColonne As Long
    Dim i As Long
    Dim valeurEntete As String

    CT_TrouverColonneEntete = 0

    derniereColonne = feuille.Cells( _
        1, _
        feuille.Columns.Count).End(XL_TO_LEFT).Column

    For i = 1 To derniereColonne

        valeurEntete = Trim$(CStr( _
            feuille.Cells(1, i).Value))

        If StrComp( _
            valeurEntete, _
            enteteRecherchee, _
            vbTextCompare) = 0 Then

            CT_TrouverColonneEntete = i
            Exit Function

        End If

    Next i

End Function

Private Function CT_LireValeurCellule( _
    ByVal feuille As Object, _
    ByVal numeroLigne As Long, _
    ByVal numeroColonne As Long) As String

    Dim valeur As Variant

    CT_LireValeurCellule = ""

    If numeroColonne <= 0 Then Exit Function

    valeur = feuille.Cells( _
        numeroLigne, _
        numeroColonne).Value

    If IsError(valeur) Then Exit Function
    If IsNull(valeur) Then Exit Function
    If IsEmpty(valeur) Then Exit Function

    CT_LireValeurCellule = Trim$(CStr(valeur))

End Function

Private Function CT_EstValeurActive( _
    ByVal valeur As String) As Boolean

    valeur = UCase$(Trim$(valeur))

    CT_EstValeurActive = _
        (valeur = "OUI" _
        Or valeur = "TRUE" _
        Or valeur = "VRAI" _
        Or valeur = "1" _
        Or valeur = "ACTIF")

End Function

Private Function CT_LirePriorite( _
    ByVal valeur As String) As Long

    If IsNumeric(valeur) Then

        CT_LirePriorite = CLng(valeur)

    Else

        CT_LirePriorite = 999999

    End If

End Function

Private Function CT_ConstruireRecordDestination( _
    ByVal idLigne As String, _
    ByVal typeExamen As String, _
    ByVal nomDestinataire As String, _
    ByVal structure As String, _
    ByVal blocDestinataire As String, _
    ByVal formuleAppel As String, _
    ByVal tutoiementVouvoiement As String, _
    ByVal priorite As String, _
    ByVal codePostal As String, _
    ByVal ville As String, _
    ByVal cleRegroupement As String, _
    ByVal formulePolitesse As String) As String

    Const SEP As String = "§§"

    CT_ConstruireRecordDestination = _
        idLigne & SEP & _
        typeExamen & SEP & _
        nomDestinataire & SEP & _
        structure & SEP & _
        blocDestinataire & SEP & _
        formuleAppel & SEP & _
        tutoiementVouvoiement & SEP & _
        priorite & SEP & _
        codePostal & SEP & _
        ville & SEP & _
        cleRegroupement & SEP & _
        formulePolitesse

End Function

Public Sub CT_TesterResolutionDestinations()

    Dim demandes As Collection
    Dim demande As Object

    Dim cleDestination As String
    Dim recordDestination As String
    Dim message As String

    Dim nombreResolues As Long
    Dim nombreACompleter As Long
    Dim i As Long

    Set demandes = CT_ConstruireCollectionDemandes()

    If demandes.Count = 0 Then

        MsgBox _
            "Aucune demande structurée n'est mémorisée.", _
            vbExclamation, _
            "Cabinet Test"

        Exit Sub

    End If

    message = _
        "Résolution des destinations :" & vbCrLf & vbCrLf

    nombreResolues = 0
    nombreACompleter = 0

    For i = 1 To demandes.Count

        Set demande = demandes(i)

        cleDestination = _
            CStr(demande("CleDestination"))

        recordDestination = _
            CT_DestinationParCle(cleDestination)

        message = message & _
            "========================================" & vbCrLf & _
            "DEMANDE " & i & vbCrLf & _
            "Clé : " & cleDestination & vbCrLf

        If Trim$(recordDestination) = "" Then

            nombreACompleter = nombreACompleter + 1

            message = message & _
                "Destination : À RENSEIGNER" & _
                vbCrLf & vbCrLf

        Else

            nombreResolues = nombreResolues + 1

            message = message & _
                "Destination : " & _
                ChampDestination(recordDestination, 2) & vbCrLf & _
                "Structure : " & _
                ChampDestination(recordDestination, 3) & vbCrLf & _
                "Formule d'appel : " & _
                ChampDestination(recordDestination, 5) & vbCrLf & _
                "Formule de politesse : " & _
                ChampDestination(recordDestination, 11) & vbCrLf & _
                "Bloc destinataire :" & vbCrLf & _
                ChampDestination(recordDestination, 4) & _
                vbCrLf & vbCrLf

        End If

    Next i

    message = message & _
        "========================================" & vbCrLf & _
        "Demandes prêtes : " & nombreResolues & vbCrLf & _
        "Destinations à renseigner : " & nombreACompleter

    If Len(message) > 6000 Then
        message = Left$(message, 6000) & _
                  vbCrLf & vbCrLf & _
                  "[Affichage tronqué]"
    End If

    MsgBox _
        message, _
        vbInformation, _
        "Destinations — Cabinet Test"

End Sub

Public Function CT_ListeDestinationsPourSelection() As Collection

    Const XL_UP As Long = -4162

    Dim resultat As Collection
    Dim ordreCles As Collection
    Dim dict As Object
    Dim destination As Object

    Dim excelApp As Object
    Dim classeur As Object
    Dim feuille As Object

    Dim colNom As Long
    Dim colType As Long
    Dim colStructure As Long
    Dim colVille As Long
    Dim colCle As Long
    Dim colActif As Long

    Dim derniereLigne As Long
    Dim i As Long

    Dim cle As String
    Dim nom As String
    Dim typeExamen As String
    Dim structure As String
    Dim ville As String
    Dim actif As String
    Dim cleBoucle As Variant

    Set resultat = New Collection
    Set ordreCles = New Collection

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    On Error GoTo GestionErreur

    Set excelApp = CreateObject("Excel.Application")

    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    excelApp.AutomationSecurity = 3
    excelApp.EnableEvents = False

    Set classeur = excelApp.Workbooks.Open( _
        CHEMIN_BASE_CORRESPONDANTS, _
        0, _
        True)

    Set feuille = _
        classeur.Worksheets("Specialistes_ParType")

    colNom = CT_TrouverColonneEntete( _
        feuille, _
        "NomDestinataire")

    colType = CT_TrouverColonneEntete( _
        feuille, _
        "TypeExamen")

    colStructure = CT_TrouverColonneEntete( _
        feuille, _
        "Structure")

    colVille = CT_TrouverColonneEntete( _
        feuille, _
        "Ville")

    colCle = CT_TrouverColonneEntete( _
        feuille, _
        "CleRegroupement")

    colActif = CT_TrouverColonneEntete( _
        feuille, _
        "Actif")

    If colNom = 0 _
    Or colType = 0 _
    Or colCle = 0 _
    Or colActif = 0 Then

        MsgBox _
            "Les colonnes nécessaires sont absentes de " & _
            "Specialistes_ParType.", _
            vbExclamation, _
            "Cabinet Test"

        GoTo Sortie

    End If

    derniereLigne = feuille.Cells( _
        feuille.Rows.Count, _
        colCle).End(XL_UP).Row

    For i = 2 To derniereLigne

        actif = CT_LireValeurCellule( _
            feuille, _
            i, _
            colActif)

        cle = UCase$(Trim$( _
            CT_LireValeurCellule( _
                feuille, _
                i, _
                colCle)))

        If CT_EstValeurActive(actif) _
        And cle <> "" _
        And cle <> "A_COMPLETER" Then

            nom = CT_LireValeurCellule( _
                feuille, _
                i, _
                colNom)

            typeExamen = CT_LireValeurCellule( _
                feuille, _
                i, _
                colType)

            structure = CT_LireValeurCellule( _
                feuille, _
                i, _
                colStructure)

            ville = CT_LireValeurCellule( _
                feuille, _
                i, _
                colVille)

            structure = _
                CT_PremiereLignePourListe(structure)

            If Not dict.Exists(cle) Then

                Set destination = _
                    CreateObject("Scripting.Dictionary")

                destination.Add "Cle", cle
                destination.Add "Nom", nom
                destination.Add "Type", typeExamen
                destination.Add "Structure", structure
                destination.Add "Ville", ville

                dict.Add cle, destination
                ordreCles.Add cle

            Else

                Set destination = dict(cle)

                'Une même clé peut correspondre à plusieurs types
                'd'examens, par exemple LOSHKAJIAN.
                If Trim$(typeExamen) <> "" Then

                    If InStr( _
                        1, _
                        "|" & destination("Type") & "|", _
                        "|" & typeExamen & "|", _
                        vbTextCompare) = 0 Then

                        If Trim$(destination("Type")) = "" Then

                            destination("Type") = typeExamen

                        Else

                            destination("Type") = _
                                destination("Type") & _
                                " / " & _
                                typeExamen

                        End If

                    End If

                End If

                If Trim$(destination("Nom")) = "" Then
                    destination("Nom") = nom
                End If

                If Trim$(destination("Structure")) = "" Then
                    destination("Structure") = structure
                End If

                If Trim$(destination("Ville")) = "" Then
                    destination("Ville") = ville
                End If

            End If

        End If

    Next i

    For Each cleBoucle In ordreCles

        Set destination = dict(CStr(cleBoucle))
        resultat.Add destination

    Next cleBoucle

Sortie:

    On Error Resume Next

    If Not classeur Is Nothing Then
        classeur.Close False
    End If

    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If

    Set feuille = Nothing
    Set classeur = Nothing
    Set excelApp = Nothing

    On Error GoTo 0

    Set CT_ListeDestinationsPourSelection = resultat

    Exit Function

GestionErreur:

    MsgBox _
        "Erreur lors du chargement des destinataires :" & _
        vbCrLf & _
        Err.Number & " - " & Err.description, _
        vbExclamation, _
        "Cabinet Test"

    Resume Sortie

End Function


Private Function CT_PremiereLignePourListe( _
    ByVal texte As String) As String

    Dim morceaux() As String

    texte = Replace(texte, vbCrLf, vbLf)
    texte = Replace(texte, vbCr, vbLf)

    morceaux = Split(texte, vbLf)

    If UBound(morceaux) >= 0 Then

        CT_PremiereLignePourListe = _
            Trim$(morceaux(0))

    Else

        CT_PremiereLignePourListe = _
            Trim$(texte)

    End If

End Function


Public Sub CT_TesterFenetreDestinations()

    gCleDestinationChoisieCabinetTest = ""

    frmCTDestination.Show

    If Trim$(gCleDestinationChoisieCabinetTest) = "" Then

        MsgBox _
            "Aucun destinataire sélectionné.", _
            vbInformation, _
            "Cabinet Test"

    Else

        MsgBox _
            "Clé sélectionnée :" & vbCrLf & _
            gCleDestinationChoisieCabinetTest, _
            vbInformation, _
            "Cabinet Test"

    End If

End Sub
