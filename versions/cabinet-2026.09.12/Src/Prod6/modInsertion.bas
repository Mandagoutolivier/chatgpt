Attribute VB_Name = "modInsertion"
Option Explicit

Public Function RemplacerCorpsOriginalParTexte(ByVal texteCorrigeAnonymise As String) As Boolean

    Dim texteRestaure As String
    Dim texteInsertion As String
    Dim debutInsertion As Long
    Dim finInsertion As Long
    Dim rngInsere As Range

    RemplacerCorpsOriginalParTexte = False

    If gPlageOriginale Is Nothing Then
        MsgBox "Aucune plage originale n'est mémorisée.", vbExclamation
        Exit Function
    End If

    If gPatient.NomComplet = "" Then
        MsgBox "Aucun patient n'est mémorisé.", vbExclamation
        Exit Function
    End If

    If texteCorrigeAnonymise = "" Then
        MsgBox "Le texte corrigé est vide.", vbExclamation
        Exit Function
    End If

    texteRestaure = RestaurerPatientDansTexte(texteCorrigeAnonymise)

    'On conserve la version restaurée en mémoire.
    gTexteCorrige = texteRestaure

    'Version réellement insérée dans Word.
    'Elle doit toujours se terminer par un retour paragraphe,
    'sinon la formule de politesse se colle au dernier paragraphe.
    texteInsertion = GarantirRetourParagrapheFinal(texteRestaure)

    debutInsertion = gPlageOriginale.Start

    gPlageOriginale.Text = texteInsertion

    finInsertion = debutInsertion + Len(texteInsertion)

    Set rngInsere = gPlageOriginale.Document.Range( _
        Start:=debutInsertion, _
        End:=finInsertion)

    rngInsere.Document.Bookmarks.Add "CORPS", rngInsere
    Set gPlageOriginale = rngInsere.Duplicate
    ConvertirMarkdownGrasDansRange rngInsere
    MettreIdentitePatientEnGrasDansRange rngInsere

    RemplacerCorpsOriginalParTexte = True
End Function

Public Function GarantirRetourParagrapheFinal(ByVal texte As String) As String

    texte = TrimFinRetoursParagraphes(texte)

    'Deux retours paragraphe :
    '1. fin du dernier paragraphe du corps
    '2. ligne vide avant la formule de politesse
    GarantirRetourParagrapheFinal = texte & vbCr & vbCr

End Function

Private Function TrimFinRetoursParagraphes(ByVal texte As String) As String

    Do While Len(texte) > 0

        If Right$(texte, 1) = vbCr _
        Or Right$(texte, 1) = vbLf _
        Or Right$(texte, 1) = Chr(11) _
        Or Right$(texte, 1) = Chr(12) _
        Or Right$(texte, 1) = " " _
        Or Right$(texte, 1) = Chr(160) Then

            texte = Left$(texte, Len(texte) - 1)

        Else
            Exit Do
        End If

    Loop

    TrimFinRetoursParagraphes = texte

End Function

Public Sub ConvertirMarkdownGrasDansRange(ByVal rngZone As Range)

    Dim doc As Document
    Dim limiteFin As Long
    Dim positionRecherche As Long
    Dim rngOuverture As Range
    Dim rngFermeture As Range
    Dim rngGras As Range
    Dim contenuDebut As Long
    Dim contenuFin As Long

    Set doc = rngZone.Document

    positionRecherche = rngZone.Start
    limiteFin = rngZone.End

    Do While positionRecherche < limiteFin

        Set rngOuverture = doc.Range(positionRecherche, limiteFin)

        With rngOuverture.Find
            .ClearFormatting
            .Text = "**"
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = False
        End With

        If Not rngOuverture.Find.Execute Then Exit Do

        Set rngFermeture = doc.Range(rngOuverture.End, limiteFin)

        With rngFermeture.Find
            .ClearFormatting
            .Text = "**"
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = False
        End With

        If Not rngFermeture.Find.Execute Then Exit Do

        contenuDebut = rngOuverture.End
        contenuFin = rngFermeture.Start

        If contenuFin > contenuDebut Then
            Set rngGras = doc.Range(contenuDebut, contenuFin)
            rngGras.Font.Bold = True
        End If

        'Suppression du marqueur fermant puis du marqueur ouvrant.
        rngFermeture.Text = ""
        rngOuverture.Text = ""

        'Chaque paire de marqueurs supprimée retire 4 caractères.
        limiteFin = limiteFin - 4

        'On reprend la recherche juste après le texte mis en gras.
        positionRecherche = contenuFin - 2

    Loop

End Sub

Public Sub MettreIdentitePatientEnGrasDansRange(ByVal rngZone As Range)

    Dim doc As Document
    Dim rngRecherche As Range
    Dim rngIdentite As Range
    Dim identiteComplete As String
    Dim identiteSansCivilite As String
    Dim debutIdentite As Long
    Dim finZone As Long

    If gPatient.NomComplet = "" Then Exit Sub
    If gPatient.civilite = "" Then Exit Sub

    Set doc = rngZone.Document

    identiteComplete = gPatient.NomComplet
    identiteSansCivilite = Trim$(Replace(identiteComplete, gPatient.civilite, "", 1, 1, vbTextCompare))

    If identiteSansCivilite = "" Then Exit Sub

    finZone = rngZone.End

    Set rngRecherche = doc.Range(rngZone.Start, rngZone.End)

    With rngRecherche.Find
        .ClearFormatting
        .Text = identiteComplete
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWildcards = False
    End With

    Do While rngRecherche.Find.Execute

        'Début du gras après "Monsieur " ou "Madame "
        debutIdentite = rngRecherche.Start + Len(gPatient.civilite) + 1

        If debutIdentite < rngRecherche.End Then
            Set rngIdentite = doc.Range( _
                Start:=debutIdentite, _
                End:=rngRecherche.End)

            rngIdentite.Font.Bold = True
        End If

        'Poursuivre la recherche après l'occurrence trouvée.
        Set rngRecherche = doc.Range(rngRecherche.End, finZone)

        With rngRecherche.Find
            .ClearFormatting
            .Text = identiteComplete
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = False
            .MatchWildcards = False
        End With

    Loop

End Sub

Public Function ExtraireBlocBalise(ByVal texteComplet As String, _
                                   ByVal baliseDebut As String, _
                                   ByVal baliseFin As String) As String

    Dim posDebut As Long
    Dim posFin As Long
    Dim posContenuDebut As Long
    Dim bloc As String

    ExtraireBlocBalise = ""

    posDebut = InStr(1, texteComplet, baliseDebut, vbTextCompare)

    If posDebut = 0 Then Exit Function

    posContenuDebut = posDebut + Len(baliseDebut)

    posFin = InStr(posContenuDebut, texteComplet, baliseFin, vbTextCompare)

    If posFin = 0 Then Exit Function

    If posFin <= posContenuDebut Then Exit Function

    bloc = Mid$(texteComplet, posContenuDebut, posFin - posContenuDebut)

    bloc = NettoyerBlocExtrait(bloc)

    ExtraireBlocBalise = bloc

End Function

Public Function NettoyerBlocExtrait(ByVal texte As String) As String

    'Retire les espaces et retours inutiles au début et à la fin du bloc,
    'sans toucher au contenu interne du courrier.

    Do While Len(texte) > 0

        If Left$(texte, 1) = vbCr _
        Or Left$(texte, 1) = vbLf _
        Or Left$(texte, 1) = Chr(11) _
        Or Left$(texte, 1) = Chr(12) _
        Or Left$(texte, 1) = " " _
        Or Left$(texte, 1) = Chr(160) Then

            texte = Mid$(texte, 2)

        Else
            Exit Do
        End If

    Loop

    Do While Len(texte) > 0

        If Right$(texte, 1) = vbCr _
        Or Right$(texte, 1) = vbLf _
        Or Right$(texte, 1) = Chr(11) _
        Or Right$(texte, 1) = Chr(12) _
        Or Right$(texte, 1) = " " _
        Or Right$(texte, 1) = Chr(160) Then

            texte = Left$(texte, Len(texte) - 1)

        Else
            Exit Do
        End If

    Loop

    NettoyerBlocExtrait = texte

End Function
