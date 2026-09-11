Attribute VB_Name = "modBaseCorrespondants"
Option Explicit

Public Function PROD_FichierExiste(ByVal chemin As String) As Boolean

    PROD_FichierExiste = (Len(Dir$(chemin)) > 0)

End Function

Private Function FeuilleExiste(ByVal wb As Object, ByVal nomFeuille As String) As Boolean

    Dim ws As Object

    FeuilleExiste = False

    For Each ws In wb.Worksheets
        If LCase$(ws.Name) = LCase$(nomFeuille) Then
            FeuilleExiste = True
            Exit Function
        End If
    Next ws

End Function

Public Sub TesterBaseCorrespondantsExcel()

    Dim xl As Object
    Dim wb As Object
    Dim message As String

    If Not PROD_FichierExiste(CHEMIN_BASE_CORRESPONDANTS) Then
        MsgBox "Base correspondants introuvable :" & vbCrLf & _
               CHEMIN_BASE_CORRESPONDANTS, vbExclamation
        Exit Sub
    End If

    On Error GoTo ErreurOuverture

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    xl.EnableEvents = False

    Set wb = xl.Workbooks.Open(CHEMIN_BASE_CORRESPONDANTS, UpdateLinks:=0, ReadOnly:=True)

    message = "Base correspondants ouverte correctement." & vbCrLf & vbCrLf

    If FeuilleExiste(wb, FEUILLE_BASE_GENERALISTES) Then
        message = message & "Feuille Generalistes : OK" & vbCrLf
    Else
        message = message & "Feuille Generalistes : ABSENTE" & vbCrLf
    End If

    If FeuilleExiste(wb, FEUILLE_BASE_SPECIALISTES) Then
        message = message & "Feuille Specialistes : OK" & vbCrLf
    Else
        message = message & "Feuille Specialistes : ABSENTE" & vbCrLf
    End If

    If FeuilleExiste(wb, FEUILLE_BASE_SPECIALISTES_PARTYPE) Then
        message = message & "Feuille Specialistes_ParType : OK" & vbCrLf
    Else
        message = message & "Feuille Specialistes_ParType : ABSENTE" & vbCrLf
    End If

    wb.Close SaveChanges:=False
    xl.Quit

    Set wb = Nothing
    Set xl = Nothing

    MsgBox message, vbInformation

    Exit Sub

ErreurOuverture:

    On Error Resume Next

    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not xl Is Nothing Then xl.Quit

    Set wb = Nothing
    Set xl = Nothing

    MsgBox "Erreur lors de l'ouverture de la base correspondants :" & vbCrLf & _
           Err.Number & " - " & Err.description, vbExclamation

End Sub

Private Function BaseColonneParEntete(ByVal ws As Object, ByVal nomEntete As String) As Long

    Dim c As Long
    Dim derniereColonne As Long
    Dim valeur As String

    BaseColonneParEntete = 0

    derniereColonne = ws.Cells(1, ws.Columns.Count).End(-4159).Column 'xlToLeft

    For c = 1 To derniereColonne

        valeur = Trim$(CStr(ws.Cells(1, c).Value))

        If LCase$(valeur) = LCase$(nomEntete) Then
            BaseColonneParEntete = c
            Exit Function
        End If

    Next c

End Function

Private Function BaseValeurCellule(ByVal ws As Object, _
                                   ByVal ligne As Long, _
                                   ByVal colonne As Long) As String

    If colonne <= 0 Then
        BaseValeurCellule = ""
    ElseIf IsError(ws.Cells(ligne, colonne).Value) Then
        BaseValeurCellule = ""
    ElseIf IsNull(ws.Cells(ligne, colonne).Value) Then
        BaseValeurCellule = ""
    Else
        BaseValeurCellule = Trim$(CStr(ws.Cells(ligne, colonne).Value))
    End If

End Function

Private Function BaseEstActif(ByVal valeurActif As String) As Boolean

    valeurActif = LCase$(Trim$(valeurActif))

    If valeurActif = "" Then
        BaseEstActif = True
    ElseIf valeurActif = "oui" Then
        BaseEstActif = True
    ElseIf valeurActif = "o" Then
        BaseEstActif = True
    ElseIf valeurActif = "1" Then
        BaseEstActif = True
    ElseIf valeurActif = "true" Then
        BaseEstActif = True
    Else
        BaseEstActif = False
    End If

End Function

Public Function RechercherDestinationsParTypeExamen(ByVal typeExamen As String) As String

    Dim xl As Object
    Dim wb As Object
    Dim ws As Object

    Dim colID As Long
    Dim colType As Long
    Dim colNom As Long
    Dim colStructure As Long
    Dim colCodePostal As Long
    Dim colVille As Long
    Dim colBloc As Long
    Dim colFormule As Long
    Dim colPriorite As Long
    Dim colActif As Long
    Dim colAValider As Long

    Dim derniereLigne As Long
    Dim r As Long
    Dim n As Long

    Dim typeLigne As String
    Dim nom As String
    Dim structure As String
    Dim cp As String
    Dim ville As String
    Dim bloc As String
    Dim formule As String
    Dim priorite As String
    Dim actif As String
    Dim aValider As String
    Dim idLigne As String

    Dim resultat As String

    RechercherDestinationsParTypeExamen = ""

    If Not PROD_FichierExiste(CHEMIN_BASE_CORRESPONDANTS) Then
        MsgBox "Base correspondants introuvable :" & vbCrLf & _
               CHEMIN_BASE_CORRESPONDANTS, vbExclamation
        Exit Function
    End If

    On Error GoTo ErreurRecherche

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    xl.EnableEvents = False

    Set wb = xl.Workbooks.Open(CHEMIN_BASE_CORRESPONDANTS, UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Worksheets(FEUILLE_BASE_SPECIALISTES_PARTYPE)

    colID = BaseColonneParEntete(ws, "ID_Ligne")
    colType = BaseColonneParEntete(ws, "TypeExamen")
    colNom = BaseColonneParEntete(ws, "NomDestinataire")
    colStructure = BaseColonneParEntete(ws, "Structure")
    colCodePostal = BaseColonneParEntete(ws, "CodePostal")
    colVille = BaseColonneParEntete(ws, "Ville")
    colBloc = BaseColonneParEntete(ws, "BlocDestinataireComplet")
    colFormule = BaseColonneParEntete(ws, "FormuleAppel")
    colPriorite = BaseColonneParEntete(ws, "Priorite")
    colActif = BaseColonneParEntete(ws, "Actif")
    colAValider = BaseColonneParEntete(ws, "AValider")

    If colType = 0 Or colNom = 0 Or colBloc = 0 Then
        MsgBox "Colonnes indispensables absentes dans Specialistes_ParType." & vbCrLf & _
               "Colonnes nécessaires : TypeExamen, NomDestinataire, BlocDestinataireComplet.", _
               vbExclamation
        GoTo Sortie
    End If

    derniereLigne = ws.Cells(ws.Rows.Count, colType).End(-4162).Row 'xlUp

    For r = 2 To derniereLigne

        typeLigne = UCase$(BaseValeurCellule(ws, r, colType))
        actif = BaseValeurCellule(ws, r, colActif)

        If typeLigne = UCase$(Trim$(typeExamen)) And BaseEstActif(actif) Then

            n = n + 1

            idLigne = BaseValeurCellule(ws, r, colID)
            nom = BaseValeurCellule(ws, r, colNom)
            structure = BaseValeurCellule(ws, r, colStructure)
            cp = BaseValeurCellule(ws, r, colCodePostal)
            ville = BaseValeurCellule(ws, r, colVille)
            bloc = BaseValeurCellule(ws, r, colBloc)
            formule = BaseValeurCellule(ws, r, colFormule)
            priorite = BaseValeurCellule(ws, r, colPriorite)
            aValider = BaseValeurCellule(ws, r, colAValider)

            resultat = resultat & n & ". " & nom

            If structure <> "" And LCase$(structure) <> LCase$(nom) Then
                resultat = resultat & " - " & structure
            End If

            If cp <> "" Or ville <> "" Then
                resultat = resultat & " - " & cp & " " & ville
            End If

            If priorite <> "" Then
                resultat = resultat & " - priorité " & priorite
            End If

            If LCase$(aValider) = "oui" Then
                resultat = resultat & " - À VALIDER"
            End If

            If idLigne <> "" Then
                resultat = resultat & " [" & idLigne & "]"
            End If

            resultat = resultat & vbCrLf
            resultat = resultat & "Formule d'appel : " & formule & vbCrLf
            resultat = resultat & "Bloc destinataire :" & vbCrLf & bloc & vbCrLf
            resultat = resultat & String(40, "-") & vbCrLf

        End If

    Next r

    RechercherDestinationsParTypeExamen = resultat

Sortie:

    On Error Resume Next

    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not xl Is Nothing Then xl.Quit

    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing

    Exit Function

ErreurRecherche:

    MsgBox "Erreur lors de la recherche dans la base correspondants :" & vbCrLf & _
           Err.Number & " - " & Err.description, vbExclamation

    Resume Sortie

End Function

Public Sub TesterRechercheDestinationsExamens()

    Dim message As String

    message = ""
    message = message & "TEST_EFFORT" & vbCrLf
    message = message & RechercherDestinationsParTypeExamen("TEST_EFFORT") & vbCrLf

    message = message & "ECHODOPPLER_VAISSEAUX_DU_COU" & vbCrLf
    message = message & RechercherDestinationsParTypeExamen("ECHODOPPLER_VAISSEAUX_DU_COU") & vbCrLf

    message = message & "COROSCANNER" & vbCrLf
    message = message & RechercherDestinationsParTypeExamen("COROSCANNER") & vbCrLf

    If Len(message) > 3500 Then
        message = Left$(message, 3500) & vbCrLf & vbCrLf & "[Message tronqué]"
    End If

    MsgBox message, vbInformation, "Test recherche destinations examens"

End Sub

Private Function PrioriteNumeriqueBase(ByVal priorite As String) As Long

    priorite = Trim$(priorite)

    If IsNumeric(priorite) Then
        PrioriteNumeriqueBase = CLng(priorite)
    Else
        PrioriteNumeriqueBase = 9999
    End If

End Function

Private Function CreerRecordDestination( _
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

    CreerRecordDestination = _
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

Public Function ChampDestination(ByVal recordDestination As String, _
                                 ByVal indexChamp As Long) As String

    Dim champs() As String

    Const SEP As String = "§§"

    champs = Split(recordDestination, SEP)

    If indexChamp < LBound(champs) Or indexChamp > UBound(champs) Then
        ChampDestination = ""
    Else
        ChampDestination = champs(indexChamp)
    End If

End Function

Public Function DestinationParDefautPourTypeExamen(ByVal typeExamen As String) As String

    Dim xl As Object
    Dim wb As Object
    Dim ws As Object

    Dim colID As Long
    Dim colType As Long
    Dim colNom As Long
    Dim colStructure As Long
    Dim colCodePostal As Long
    Dim colVille As Long
    Dim colBloc As Long
    Dim colFormule As Long
    Dim colTuto As Long
    Dim colPriorite As Long
    Dim colActif As Long, colAValider As Long
    Dim colCleRegroupement As Long
    Dim colFormulePolitesse As Long

    Dim derniereLigne As Long
    Dim r As Long
    Dim meilleureLigne As Long
    Dim meilleurePriorite As Long
    Dim prioriteLigne As Long

    Dim typeLigne As String
    Dim actif As String
    Dim priorite As String

    Dim idLigne As String
    Dim nom As String
    Dim structure As String
    Dim cp As String
    Dim ville As String
    Dim bloc As String
    Dim formule As String
    Dim tuto As String
    Dim cleRegroupement As String
    Dim formulePolitesse As String


    DestinationParDefautPourTypeExamen = ""

    If Not PROD_FichierExiste(CHEMIN_BASE_CORRESPONDANTS) Then
        MsgBox "Base correspondants introuvable :" & vbCrLf & _
               CHEMIN_BASE_CORRESPONDANTS, vbExclamation
        Exit Function
    End If

    On Error GoTo ErreurRecherche

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    xl.EnableEvents = False

    Set wb = xl.Workbooks.Open(CHEMIN_BASE_CORRESPONDANTS, UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Worksheets(FEUILLE_BASE_SPECIALISTES_PARTYPE)

    colID = BaseColonneParEntete(ws, "ID_Ligne")
    colType = BaseColonneParEntete(ws, "TypeExamen")
    colNom = BaseColonneParEntete(ws, "NomDestinataire")
    colStructure = BaseColonneParEntete(ws, "Structure")
    colCodePostal = BaseColonneParEntete(ws, "CodePostal")
    colVille = BaseColonneParEntete(ws, "Ville")
    colBloc = BaseColonneParEntete(ws, "BlocDestinataireComplet")
    colFormule = BaseColonneParEntete(ws, "FormuleAppel")
    colTuto = BaseColonneParEntete(ws, "TutoiementVouvoiement")
    colPriorite = BaseColonneParEntete(ws, "Priorite")
    colActif = BaseColonneParEntete(ws, "Actif")
    colAValider = BaseColonneParEntete(ws, "AValider")
    colCleRegroupement = BaseColonneParEntete(ws, "CleRegroupement")
    colFormulePolitesse = BaseColonneParEntete(ws, "FormulePolitesse")

    If colType = 0 Or colNom = 0 Or colBloc = 0 Then
        MsgBox "Colonnes indispensables absentes dans Specialistes_ParType." & vbCrLf & _
               "Colonnes nécessaires : TypeExamen, NomDestinataire, BlocDestinataireComplet.", _
               vbExclamation
        GoTo Sortie
    End If

    derniereLigne = ws.Cells(ws.Rows.Count, colType).End(-4162).Row 'xlUp

    meilleureLigne = 0
    meilleurePriorite = 999999

    For r = 2 To derniereLigne

        typeLigne = UCase$(BaseValeurCellule(ws, r, colType))
        actif = BaseValeurCellule(ws, r, colActif)

        If typeLigne = UCase$(Trim$(typeExamen)) And BaseEstActif(actif) Then

            If colAValider > 0 Then
                If BaseEstActif(BaseValeurCellule(ws, r, colAValider)) Then GoTo LigneSuivanteDefaut
            End If
            priorite = BaseValeurCellule(ws, r, colPriorite)
            prioriteLigne = PrioriteNumeriqueBase(priorite)

            If meilleureLigne = 0 Or prioriteLigne < meilleurePriorite Then
                meilleureLigne = r
                meilleurePriorite = prioriteLigne
            End If

        End If

LigneSuivanteDefaut:
    Next r

    If meilleureLigne = 0 Then GoTo Sortie

    idLigne = BaseValeurCellule(ws, meilleureLigne, colID)
nom = BaseValeurCellule(ws, meilleureLigne, colNom)
structure = BaseValeurCellule(ws, meilleureLigne, colStructure)
cp = BaseValeurCellule(ws, meilleureLigne, colCodePostal)
ville = BaseValeurCellule(ws, meilleureLigne, colVille)
bloc = BaseValeurCellule(ws, meilleureLigne, colBloc)
formule = BaseValeurCellule(ws, meilleureLigne, colFormule)
tuto = BaseValeurCellule(ws, meilleureLigne, colTuto)
priorite = BaseValeurCellule(ws, meilleureLigne, colPriorite)

cleRegroupement = BaseValeurCellule( _
    ws, _
    meilleureLigne, _
    colCleRegroupement)

formulePolitesse = BaseValeurCellule( _
    ws, _
    meilleureLigne, _
    colFormulePolitesse)

DestinationParDefautPourTypeExamen = CreerRecordDestination( _
    idLigne, _
    UCase$(Trim$(typeExamen)), _
    nom, _
    structure, _
    bloc, _
    formule, _
    tuto, _
    priorite, _
    cp, _
    ville, _
    cleRegroupement, _
    formulePolitesse)

Sortie:

    On Error Resume Next

    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not xl Is Nothing Then xl.Quit

    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing

    Exit Function

ErreurRecherche:

    MsgBox "Erreur lors de la sélection de la destination par défaut :" & vbCrLf & _
           Err.Number & " - " & Err.description, vbExclamation

    Resume Sortie

End Function

Public Function DecrireDestinationParDefaut(ByVal recordDestination As String) As String

    If recordDestination = "" Then
        DecrireDestinationParDefaut = "Aucune destination trouvée."
        Exit Function
    End If

    DecrireDestinationParDefaut = _
        "ID : " & ChampDestination(recordDestination, 0) & vbCrLf & _
        "Type examen : " & ChampDestination(recordDestination, 1) & vbCrLf & _
        "Nom : " & ChampDestination(recordDestination, 2) & vbCrLf & _
        "Structure : " & ChampDestination(recordDestination, 3) & vbCrLf & _
        "Code postal / ville : " & ChampDestination(recordDestination, 8) & " " & ChampDestination(recordDestination, 9) & vbCrLf & _
        "Formule d'appel : " & ChampDestination(recordDestination, 5) & vbCrLf & _
        "Tutoiement/vouvoiement : " & ChampDestination(recordDestination, 6) & vbCrLf & _
        "Priorité : " & ChampDestination(recordDestination, 7) & vbCrLf & _
        "Bloc destinataire :" & vbCrLf & _
        ChampDestination(recordDestination, 4)

End Function

Public Sub TesterDestinationParDefautExamens()

    Dim message As String
    Dim record As String

    message = ""

    record = DestinationParDefautPourTypeExamen("TEST_EFFORT")
    message = message & "TEST_EFFORT" & vbCrLf
    message = message & DecrireDestinationParDefaut(record) & vbCrLf & _
              String(40, "-") & vbCrLf

    record = DestinationParDefautPourTypeExamen("SCINTIGRAPHIE_MYOCARDIQUE")
    message = message & "SCINTIGRAPHIE_MYOCARDIQUE" & vbCrLf
    message = message & DecrireDestinationParDefaut(record) & vbCrLf & _
              String(40, "-") & vbCrLf

    record = DestinationParDefautPourTypeExamen("ECHODOPPLER_VAISSEAUX_DU_COU")
    message = message & "ECHODOPPLER_VAISSEAUX_DU_COU" & vbCrLf
    message = message & DecrireDestinationParDefaut(record) & vbCrLf & _
              String(40, "-") & vbCrLf

    record = DestinationParDefautPourTypeExamen("ECHODOPPLER_MEMBRES_INFERIEURS")
    message = message & "ECHODOPPLER_MEMBRES_INFERIEURS" & vbCrLf
    message = message & DecrireDestinationParDefaut(record) & vbCrLf & _
              String(40, "-") & vbCrLf

    record = DestinationParDefautPourTypeExamen("COROSCANNER")
    message = message & "COROSCANNER" & vbCrLf
    message = message & DecrireDestinationParDefaut(record) & vbCrLf & _
              String(40, "-") & vbCrLf

    record = DestinationParDefautPourTypeExamen("SCORE_CALCIQUE")
    message = message & "SCORE_CALCIQUE" & vbCrLf
    message = message & DecrireDestinationParDefaut(record)

    If Len(message) > 6000 Then
        message = Left$(message, 6000) & vbCrLf & vbCrLf & "[Message tronqué]"
    End If

    MsgBox message, vbInformation, "Destinations par défaut"

End Sub

Public Sub TesterFormulesLoshkajian()

    Dim recordDestination As String
    Dim message As String

    recordDestination = DestinationParDefautPourTypeExamen( _
        "ECHODOPPLER_VAISSEAUX_DU_COU")

    If recordDestination = "" Then
        MsgBox "Destination Loshkajian introuvable.", vbExclamation
        Exit Sub
    End If

    message = _
        "Destinataire :" & vbCrLf & _
        ChampDestination(recordDestination, 2) & vbCrLf & vbCrLf & _
        "Formule d'appel :" & vbCrLf & _
        ChampDestination(recordDestination, 5) & vbCrLf & vbCrLf & _
        "Formule de politesse :" & vbCrLf & _
        ChampDestination(recordDestination, 11) & vbCrLf & vbCrLf & _
        "Clé de regroupement :" & vbCrLf & _
        ChampDestination(recordDestination, 10)

    MsgBox message, vbInformation, "Formules Loshkajian"

End Sub


' Les formules historiques s arretaient a la ligne 499 : inclure les nouvelles structures.
Public Sub PROD_ActualiserPlagesStructures(ByVal wb As Object)
    Dim ws As Object, cellule As Object, derniere As Long, formule As String
    Dim re As Object, occurrences As Object, i As Long, occurrence As Object, remplacement As String
    Set ws = wb.Worksheets("Structures")
    derniere = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    If derniere < 2 Then derniere = 2
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "Structures!\$[A-Z]+\$2:\$[A-Z]+\$[0-9]+"
    For Each cellule In wb.Worksheets("Saisie_Correspondants").Range("E8:E13")
        If cellule.HasFormula Then
            formule = CStr(cellule.Formula)
            Set occurrences = re.Execute(formule)
            For i = occurrences.Count - 1 To 0 Step -1
                Set occurrence = occurrences(i)
                remplacement = Left$(occurrence.Value, InStrRev(occurrence.Value, "$")) & CStr(derniere)
                formule = Left$(formule, occurrence.FirstIndex) & remplacement & Mid$(formule, occurrence.FirstIndex + occurrence.Length + 1)
            Next i
            cellule.Formula = formule
        End If
    Next cellule
End Sub
