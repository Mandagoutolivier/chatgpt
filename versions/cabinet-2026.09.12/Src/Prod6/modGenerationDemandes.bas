Attribute VB_Name = "modGenerationDemandes"
Option Explicit

Public Function CreerCopieDocumentActifPourDemande() As Document

    Set CreerCopieDocumentActifPourDemande = _
        CreerCopieDocumentSourcePourDemande(ActiveDocument)

End Function

Public Sub TesterCreationCopieDemande()

    Dim docCopie As Document

    Set docCopie = CreerCopieDocumentActifPourDemande()

    docCopie.Activate

    MsgBox "Copie du courrier créée." & vbCrLf & _
           "Vérifiez que la mise en page générale est conservée." & vbCrLf & vbCrLf & _
           "Le document original n'a pas été modifié.", _
           vbInformation

End Sub

Public Function LocaliserBlocDestinataireAvantDate(ByVal doc As Document, _
                                                   ByRef rngDestinataire As Range, _
                                                   ByRef rngDate As Range) As Boolean

    Dim i As Long
    Dim idxDate As Long
    Dim idxFinDest As Long
    Dim idxDebutDest As Long
    Dim textePara As String

    LocaliserBlocDestinataireAvantDate = False

    idxDate = 0
    idxFinDest = 0
    idxDebutDest = 0

    'Recherche de la ligne de date Beaumont ... le ... année.
    For i = 1 To doc.Paragraphs.Count

        textePara = NettoyerTexteParagraphe(doc.Paragraphs(i).Range.Text)

        If EstParagrapheDateCourrier(textePara) Then
            idxDate = i
            Exit For
        End If

    Next i

    If idxDate = 0 Then
        MsgBox "Ligne de date Beaumont ... le ... introuvable.", vbExclamation
        Exit Function
    End If

    Set rngDate = doc.Paragraphs(idxDate).Range

    'On remonte depuis la date jusqu'au dernier paragraphe utile.
    idxFinDest = idxDate - 1

    Do While idxFinDest >= 1

        textePara = NettoyerTexteParagraphe(doc.Paragraphs(idxFinDest).Range.Text)

        If textePara = "" _
        Or EstMentionHorsDestinataireAvantDate(textePara) Then

            idxFinDest = idxFinDest - 1

        Else
            Exit Do
        End If

    Loop

    If idxFinDest < 1 Then
        MsgBox "Fin du bloc destinataire introuvable.", vbExclamation
        Exit Function
    End If

    'On remonte jusqu'à la ligne vide précédant le bloc destinataire.
    idxDebutDest = idxFinDest

    Do While idxDebutDest > 1

        textePara = NettoyerTexteParagraphe(doc.Paragraphs(idxDebutDest - 1).Range.Text)

        If textePara = "" Then
            Exit Do
        End If

        idxDebutDest = idxDebutDest - 1

    Loop

    Set rngDestinataire = doc.Range( _
        Start:=doc.Paragraphs(idxDebutDest).Range.Start, _
        End:=doc.Paragraphs(idxFinDest).Range.End)

    LocaliserBlocDestinataireAvantDate = True

End Function

Private Function EstMentionHorsDestinataireAvantDate(ByVal texte As String) As Boolean

    Dim t As String

    t = LCase$(NettoyerTexteParagraphe(texte))

    EstMentionHorsDestinataireAvantDate = False

    If t = "" Then Exit Function

    If InStr(1, t, "lettre dictée", vbTextCompare) > 0 Then
        EstMentionHorsDestinataireAvantDate = True
        Exit Function
    End If

    If InStr(1, t, "lettre dictee", vbTextCompare) > 0 Then
        EstMentionHorsDestinataireAvantDate = True
        Exit Function
    End If

    If InStr(1, t, "remise en main propre", vbTextCompare) > 0 Then
        EstMentionHorsDestinataireAvantDate = True
        Exit Function
    End If

End Function

Public Sub TesterLocaliserBlocDestinataire()

    Dim rngDest As Range
    Dim rngDate As Range
    Dim msg As String

    If Not LocaliserBlocDestinataireAvantDate(ActiveDocument, rngDest, rngDate) Then
        Exit Sub
    End If

    msg = "Bloc destinataire détecté :" & vbCrLf & vbCrLf & _
          rngDest.Text & vbCrLf & _
          String(40, "-") & vbCrLf & _
          "Date détectée :" & vbCrLf & _
          rngDate.Text

    MsgBox msg, vbInformation, "Test bloc destinataire"

End Sub

Private Function GarantirRetourParagrapheTexte(ByVal texte As String) As String

    texte = Replace(texte, vbCrLf, vbCr)
    texte = Replace(texte, vbLf, vbCr)

    Do While Len(texte) > 0
        If Right$(texte, 1) = vbCr _
        Or Right$(texte, 1) = " " _
        Or Right$(texte, 1) = Chr(160) Then
            texte = Left$(texte, Len(texte) - 1)
        Else
            Exit Do
        End If
    Loop

    GarantirRetourParagrapheTexte = texte & vbCr

End Function

Public Sub RemplacerBlocDestinataireDansDocument(ByVal doc As Document, _
                                                 ByVal nouveauBlocDestinataire As String)

    Dim rngDest As Range
    Dim rngDate As Range
    Dim debutRemplacement As Long
    Dim finRemplacement As Long
    Dim rngNouveau As Range

    If Not LocaliserBlocDestinataireAvantDate(doc, rngDest, rngDate) Then
        Exit Sub
    End If

    debutRemplacement = rngDest.Start

    rngDest.Text = GarantirRetourParagrapheTexte(nouveauBlocDestinataire)

    finRemplacement = debutRemplacement + Len(GarantirRetourParagrapheTexte(nouveauBlocDestinataire))

    Set rngNouveau = doc.Range(Start:=debutRemplacement, End:=finRemplacement)

    'On applique au nouveau bloc le style du texte environnant.
    rngNouveau.Font.Name = "Times New Roman"
    rngNouveau.Font.Size = 12

End Sub

Public Sub TesterRemplacementBlocDestinataireDansCopie()

    Dim docCopie As Document
    Dim nouveauDestinataire As String

    nouveauDestinataire = _
        "Docteur Loshkajian" & vbCr & _
        "[Adresse à compléter]" & vbCr & _
        "[Code postal] [Ville]"

    Set docCopie = CreerCopieDocumentActifPourDemande()

    docCopie.Activate

    RemplacerBlocDestinataireDansDocument docCopie, nouveauDestinataire

    MsgBox "Copie créée avec remplacement du bloc destinataire." & vbCrLf & _
           "Vérifiez que le destinataire a été remplacé et que la date reste bien séparée.", _
           vbInformation

End Sub

Public Function LocaliserFormuleAppelApresDate(ByVal doc As Document, _
                                               ByRef rngAppel As Range) As Boolean

    Dim i As Long
    Dim idxDate As Long
    Dim textePara As String

    LocaliserFormuleAppelApresDate = False
    idxDate = 0

    'Recherche de la ligne de date Beaumont ... le ... année.
    For i = 1 To doc.Paragraphs.Count

        textePara = NettoyerTexteParagraphe(doc.Paragraphs(i).Range.Text)

        If EstParagrapheDateCourrier(textePara) Then
            idxDate = i
            Exit For
        End If

    Next i

    If idxDate = 0 Then
        MsgBox "Ligne de date introuvable pour localiser la formule d'appel.", vbExclamation
        Exit Function
    End If

    'La formule d'appel est le premier paragraphe non vide après la date.
    For i = idxDate + 1 To doc.Paragraphs.Count

        textePara = NettoyerTexteParagraphe(doc.Paragraphs(i).Range.Text)

        If textePara <> "" Then
            Set rngAppel = doc.Paragraphs(i).Range
            LocaliserFormuleAppelApresDate = True
            Exit Function
        End If

    Next i

    MsgBox "Formule d'appel introuvable après la date.", vbExclamation

End Function

Public Sub RemplacerFormuleAppelDansDocument(ByVal doc As Document, _
                                             ByVal nouvelleFormuleAppel As String)

    Dim rngAppel As Range
    Dim debutRemplacement As Long
    Dim finRemplacement As Long
    Dim rngNouveau As Range
    Dim texteInsertion As String

    If Not LocaliserFormuleAppelApresDate(doc, rngAppel) Then
        Exit Sub
    End If

    debutRemplacement = rngAppel.Start

    texteInsertion = GarantirRetourParagrapheTexte(nouvelleFormuleAppel)

    rngAppel.Text = texteInsertion

    finRemplacement = debutRemplacement + Len(texteInsertion)

    Set rngNouveau = doc.Range(Start:=debutRemplacement, End:=finRemplacement)

    rngNouveau.Font.Name = "Times New Roman"
    rngNouveau.Font.Size = 10

End Sub

Public Sub TesterRemplacementDestinataireEtFormuleDansCopie()

    Dim docCopie As Document
    Dim nouveauDestinataire As String
    Dim nouvelleFormule As String

    nouveauDestinataire = _
        "Docteur Loshkajian" & vbCr & _
        "[Adresse à compléter]" & vbCr & _
        "[Code postal] [Ville]"

    nouvelleFormule = "Cher Confrère,"

    Set docCopie = CreerCopieDocumentActifPourDemande()

    docCopie.Activate

    RemplacerBlocDestinataireDansDocument docCopie, nouveauDestinataire
    RemplacerFormuleAppelDansDocument docCopie, nouvelleFormule

    MsgBox "Copie créée avec remplacement du destinataire et de la formule d'appel." & vbCrLf & _
           "Vérifiez que le courrier original n'a pas été modifié.", _
           vbInformation

End Sub

Public Function RemplacerCorpsDansDocument(ByVal doc As Document, _
                                           ByVal nouveauCorps As String) As Boolean

    Dim rngCorps As Range
    Dim rngPremierPatient As Range
    Dim rngNouveau As Range
    Dim texteInsertion As String
    Dim debutInsertion As Long
    Dim finInsertion As Long

    RemplacerCorpsDansDocument = False

    If Trim$(nouveauCorps) = "" Then
        MsgBox "Le nouveau corps de texte est vide.", vbExclamation
        Exit Function
    End If

    If Not LocaliserCorpsCourrier(doc, rngCorps, rngPremierPatient) Then
        MsgBox "Impossible de localiser le corps du courrier dans la copie.", _
               vbExclamation
        Exit Function
    End If

    debutInsertion = rngCorps.Start

    'Cette fonction existante ajoute deux retours de paragraphe
    'pour conserver une ligne vide avant la formule de politesse.
    texteInsertion = GarantirRetourParagrapheFinal(nouveauCorps)

    rngCorps.Text = texteInsertion

    finInsertion = debutInsertion + Len(texteInsertion)

    Set rngNouveau = doc.Range( _
        Start:=debutInsertion, _
        End:=finInsertion)

    rngNouveau.Font.Name = "Times New Roman"
    rngNouveau.Font.Size = 10

    'Conversion éventuelle de **texte** en vrai gras Word.
    ConvertirMarkdownGrasDansRange rngNouveau

    RemplacerCorpsDansDocument = True

End Function

Public Sub TesterRemplacementCompletDansCopie()

    Dim docCopie As Document
    Dim nouveauDestinataire As String
    Dim nouvelleFormule As String
    Dim nouveauCorps As String

    nouveauDestinataire = _
        "Docteur Loshkajian" & vbCr & _
        "[Adresse à compléter]" & vbCr & _
        "[Code postal] [Ville]"

    nouvelleFormule = "Cher Confrère,"

    nouveauCorps = _
        "Je vous prie de bien vouloir réaliser chez le patient concerné un " & _
        "**écho-Doppler des vaisseaux du cou et des membres inférieurs**." & vbCr & vbCr & _
        "Cet examen est demandé afin de compléter l'évaluation de son terrain vasculaire." & vbCr & vbCr & _
        "Je vous remercie par avance pour votre prise en charge."

    Set docCopie = CreerCopieDocumentActifPourDemande()

    docCopie.Activate

    RemplacerBlocDestinataireDansDocument _
        docCopie, _
        nouveauDestinataire

    RemplacerFormuleAppelDansDocument _
        docCopie, _
        nouvelleFormule

    If Not RemplacerCorpsDansDocument(docCopie, nouveauCorps) Then
        Exit Sub
    End If

    MsgBox "Copie créée avec remplacement du destinataire," & vbCrLf & _
           "de la formule d'appel et du corps du courrier." & vbCrLf & vbCrLf & _
           "Vérifiez que la formule de politesse et la signature sont conservées.", _
           vbInformation

End Sub

Private Sub AppliquerRetrait8CmAuParagraphe(ByVal p As Paragraph)

    If p Is Nothing Then Exit Sub

    With p.Range.ParagraphFormat
        .LeftIndent = CentimetersToPoints(8)
        .RightIndent = 0
        .FirstLineIndent = 0
        .Alignment = wdAlignParagraphLeft
    End With

End Sub

Private Function TrouverDernierParagrapheSignature( _
    ByVal doc As Document, _
    ByRef pSignature As Paragraph) As Boolean

    Dim rngRecherche As Range
    Dim shp As Shape

    TrouverDernierParagrapheSignature = False

    If doc Is Nothing Then Exit Function

    'Recherche d'abord dans le corps principal du document.
    Set rngRecherche = doc.Content.Duplicate

    With rngRecherche.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "Mandagout"
        .Forward = False
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
    End With

    If rngRecherche.Find.Execute Then
        Set pSignature = rngRecherche.Paragraphs(1)
        TrouverDernierParagrapheSignature = True
        Exit Function
    End If

    'Secours : signature placée dans une zone de texte.
    For Each shp In doc.Shapes

        If shp.TextFrame.HasText <> 0 Then

            Set rngRecherche = shp.TextFrame.TextRange.Duplicate

            With rngRecherche.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "Mandagout"
                .Forward = False
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
            End With

            If rngRecherche.Find.Execute Then
                'Dans une zone de texte, la position visuelle dépend
                'surtout de la zone elle-même. On place son bord gauche
                'à 8 cm de la marge puis on aligne le texte à gauche.
                On Error Resume Next
                shp.RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
                shp.Left = CentimetersToPoints(8)
                On Error GoTo 0

                Set pSignature = rngRecherche.Paragraphs(1)
                With pSignature.Range.ParagraphFormat
                    .LeftIndent = 0
                    .RightIndent = 0
                    .FirstLineIndent = 0
                    .Alignment = wdAlignParagraphLeft
                End With

                TrouverDernierParagrapheSignature = True
                Exit Function
            End If

        End If

    Next shp

End Function

Public Sub RetablirMiseEnPageSignature(ByVal doc As Document)

    Dim pSignature As Paragraph

    If doc Is Nothing Then Exit Sub

    If TrouverDernierParagrapheSignature(doc, pSignature) Then
        'Pour une signature du corps principal, impose réellement 8 cm.
        'Pour une zone de texte, le déplacement de la zone est déjà
        'effectué dans TrouverDernierParagrapheSignature.
        If pSignature.Range.StoryType = wdMainTextStory Then
            AppliquerRetrait8CmAuParagraphe pSignature
        End If
    End If

End Sub

Public Sub TesterRetraitSignature8cm()

    Dim pSignature As Paragraph
    Dim retraitAvant As Single
    Dim retraitApres As Single
    Dim dansTableau As Boolean
    Dim typeZone As Long
    Dim msg As String

    If Not TrouverDernierParagrapheSignature(ActiveDocument, pSignature) Then
        MsgBox "Signature contenant MANDAGOUT introuvable.", vbExclamation
        Exit Sub
    End If

    retraitAvant = pSignature.Range.ParagraphFormat.LeftIndent
    typeZone = pSignature.Range.StoryType
    dansTableau = pSignature.Range.Information(wdWithInTable)

    If typeZone = wdMainTextStory Then
        AppliquerRetrait8CmAuParagraphe pSignature
    End If

    retraitApres = pSignature.Range.ParagraphFormat.LeftIndent

    msg = "Texte détecté : " & NettoyerTexteParagraphe(pSignature.Range.Text) & vbCrLf & vbCrLf & _
          "Retrait avant : " & Format$(retraitAvant / CentimetersToPoints(1), "0.00") & " cm" & vbCrLf & _
          "Retrait après : " & Format$(retraitApres / CentimetersToPoints(1), "0.00") & " cm" & vbCrLf & _
          "Dans un tableau : " & IIf(dansTableau, "Oui", "Non") & vbCrLf & _
          "StoryType : " & CStr(typeZone)

    MsgBox msg, vbInformation, "Diagnostic signature"

End Sub

Public Function CreerCopieDocumentSourcePourDemande( _
    ByVal docSource As Document) As Document

    Dim docNouveau As Document

    If docSource Is Nothing Then
        Set CreerCopieDocumentSourcePourDemande = Nothing
        Exit Function
    End If

    docSource.Range.Copy

    Set docNouveau = Documents.Add

    docNouveau.Range.PasteAndFormat _
        wdFormatOriginalFormatting

    Set CreerCopieDocumentSourcePourDemande = docNouveau

End Function

Public Sub TesterDeuxCopiesDepuisMemeOriginal()

    Dim docSource As Document
    Dim docCopie1 As Document
    Dim docCopie2 As Document

    Dim destinataire1 As String
    Dim destinataire2 As String

    Set docSource = ActiveDocument

    destinataire1 = _
        "DESTINATAIRE DE TEST NUMÉRO 1" & vbCr & _
        "Adresse de test 1" & vbCr & _
        "95000 VILLE TEST 1"

    destinataire2 = _
        "DESTINATAIRE DE TEST NUMÉRO 2" & vbCr & _
        "Adresse de test 2" & vbCr & _
        "60000 VILLE TEST 2"

    Set docCopie1 = _
        CreerCopieDocumentSourcePourDemande(docSource)

    If docCopie1 Is Nothing Then
        MsgBox "La première copie n'a pas pu être créée.", _
               vbExclamation
        Exit Sub
    End If

    RemplacerBlocDestinataireDansDocument _
        docCopie1, _
        destinataire1

    Set docCopie2 = _
        CreerCopieDocumentSourcePourDemande(docSource)

    If docCopie2 Is Nothing Then
        MsgBox "La deuxième copie n'a pas pu être créée.", _
               vbExclamation
        Exit Sub
    End If

    RemplacerBlocDestinataireDansDocument _
        docCopie2, _
        destinataire2

    docCopie2.Activate

    MsgBox _
        "Deux copies ont été créées depuis le même courrier original." & _
        vbCrLf & vbCrLf & _
        "Vérifiez que :" & vbCrLf & _
        "- le premier document contient le destinataire numéro 1 ;" & vbCrLf & _
        "- le second contient le destinataire numéro 2 ;" & vbCrLf & _
        "- les deux documents reprennent le courrier original complet ;" & vbCrLf & _
        "- le courrier original n'a pas été modifié.", _
        vbInformation

End Sub


Public Sub TesterFormulesLoshkajianDansCopie()

    Dim docSource As Document
    Dim docCopie As Document
    Dim recordDestination As String

    Dim blocDestinataire As String
    Dim formuleAppel As String
    Dim formulePolitesse As String

    Set docSource = ActiveDocument

    recordDestination = _
        DestinationParDefautPourTypeExamen( _
            "ECHODOPPLER_VAISSEAUX_DU_COU")

    If recordDestination = "" Then
        MsgBox "Destination Loshkajian introuvable.", _
               vbExclamation
        Exit Sub
    End If

    blocDestinataire = _
        ChampDestination(recordDestination, 4)

    formuleAppel = _
        ChampDestination(recordDestination, 5)

    formulePolitesse = _
        ChampDestination(recordDestination, 11)

    Set docCopie = _
        CreerCopieDocumentSourcePourDemande(docSource)

    If docCopie Is Nothing Then
        MsgBox "Impossible de créer la copie du courrier.", _
               vbExclamation
        Exit Sub
    End If

    RemplacerBlocDestinataireDansDocument _
        docCopie, _
        blocDestinataire

    RemplacerFormuleAppelDansDocument _
        docCopie, _
        formuleAppel

    If Not RemplacerFormulePolitesseDansDocument( _
        docCopie, _
        formulePolitesse) Then

        Exit Sub

    End If

    docCopie.Activate

    MsgBox _
        "Copie créée avec les formules du Dr Loshkajian." & _
        vbCrLf & vbCrLf & _
        "Vérifiez :" & vbCrLf & _
        "- le bloc destinataire ;" & vbCrLf & _
        "- la formule d'appel ;" & vbCrLf & _
        "- la formule de politesse ;" & vbCrLf & _
        "- la conservation des retraits et de la signature.", _
        vbInformation

End Sub

Public Function RemplacerFormulePolitesseDansDocument( _
    ByVal doc As Document, _
    ByVal nouvelleFormulePolitesse As String) As Boolean

    Dim i As Long
    Dim idxSignature As Long
    Dim idxPolitesse As Long

    Dim textePara As String
    Dim texteSignature As String
    Dim texteBloc As String

    Dim positionDebut As Long
    Dim positionFin As Long
    Dim positionSignature As Long

    Dim rngBloc As Range
    Dim rngPolitesse As Range
    Dim rngSignature As Range

    RemplacerFormulePolitesseDansDocument = False

    If doc Is Nothing Then
        MsgBox "Document Word absent.", vbExclamation
        Exit Function
    End If

    nouvelleFormulePolitesse = Trim$(nouvelleFormulePolitesse)

    If nouvelleFormulePolitesse = "" Then
        MsgBox "La formule de politesse provenant de la base est vide.", _
               vbExclamation
        Exit Function
    End If

    '====================================================
    ' V4 : la signature n'est plus conservée/reformatée.
    ' Elle est reconstruite systématiquement avec la
    ' formule de politesse, ce qui évite tout héritage
    ' de mise en forme provenant du courrier original.
    '====================================================

    'HF1 AFFICHAGE : suppression de l'artefact diagnostique avec guillemets.
    texteSignature = "Docteur Olivier Mandagout"

    '====================================================
    ' 1. Repérer l'ancienne signature finale.
    '    Recherche depuis la fin pour ne pas confondre
    '    avec un éventuel en-tête contenant Mandagout.
    '====================================================

    idxSignature = 0

    For i = doc.Paragraphs.Count To 1 Step -1

        textePara = NettoyerTexteParagraphe( _
            doc.Paragraphs(i).Range.Text)

        If InStr(1, textePara, "Mandagout", vbTextCompare) > 0 _
        And (InStr(1, textePara, "Olivier", vbTextCompare) > 0 _
             Or InStr(1, textePara, "Docteur", vbTextCompare) > 0 _
             Or InStr(1, textePara, "Dr ", vbTextCompare) > 0) Then

            idxSignature = i
            Exit For

        End If

    Next i

    If idxSignature = 0 Then
        MsgBox "Signature finale Olivier Mandagout introuvable.", _
               vbExclamation
        Exit Function
    End If

    '====================================================
    ' 2. Repérer la formule de politesse : premier
    '    paragraphe non vide précédant la signature.
    '====================================================

    idxPolitesse = idxSignature - 1

    Do While idxPolitesse >= 1

        textePara = NettoyerTexteParagraphe( _
            doc.Paragraphs(idxPolitesse).Range.Text)

        If textePara <> "" Then Exit Do

        idxPolitesse = idxPolitesse - 1

    Loop

    If idxPolitesse < 1 Then
        MsgBox "Formule de politesse introuvable avant la signature finale.", _
               vbExclamation
        Exit Function
    End If

    '====================================================
    ' 3. Supprimer l'ancien bloc puis reconstruire :
    '       formule de politesse
    '       ligne vide
    '       Docteur Olivier Mandagout
    '====================================================

    positionDebut = doc.Paragraphs(idxPolitesse).Range.Start
    positionFin = doc.Paragraphs(idxSignature).Range.End

    'La signature est précédée de HUIT tabulations réelles.
    'Aucun taquet personnalisé n'est créé : on reproduit exactement
    'le comportement obtenu manuellement dans le document final.
    texteBloc = nouvelleFormulePolitesse & vbCr & _
                vbCr & _
                vbTab & vbTab & vbTab & vbTab & _
                vbTab & vbTab & vbTab & vbTab & _
                texteSignature & vbCr

    Set rngBloc = doc.Range( _
        Start:=positionDebut, _
        End:=positionFin)

    rngBloc.Text = texteBloc

    '====================================================
    ' 4. Mise en forme explicite de la formule.
    '    Elle suit le retrait normal du corps : 3 cm.
    '====================================================

    Set rngPolitesse = doc.Range( _
        Start:=positionDebut, _
        End:=positionDebut + Len(nouvelleFormulePolitesse))

    With rngPolitesse
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Italic = False
    End With

    With rngPolitesse.ParagraphFormat
        .LeftIndent = CentimetersToPoints(3)
        .RightIndent = 0
        .FirstLineIndent = 0
        .Alignment = wdAlignParagraphLeft
    End With

    '====================================================
    ' 5. Mise en forme explicite de la NOUVELLE signature.
    '    Aucun format de l'ancienne signature n'est repris.
    '====================================================

    positionSignature = positionDebut + _
                        Len(nouvelleFormulePolitesse) + 2

    'positionSignature pointe sur la première des huit tabulations.
    Set rngSignature = doc.Range( _
        Start:=positionSignature, _
        End:=positionSignature + 8 + Len(texteSignature))

    With rngSignature
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Italic = False
    End With

    'Aucun taquet de tabulation personnalisé : les huit caractères TAB
    'restent réellement présents dans le texte du document final.
    'On neutralise seulement les retraits de paragraphe pour que leur
    'position dépende exclusivement des tabulations.
    With rngSignature.ParagraphFormat
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
        .Alignment = wdAlignParagraphLeft
    End With

    RemplacerFormulePolitesseDansDocument = True

End Function
