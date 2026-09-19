Attribute VB_Name = "modMiseEnPageFinale"
Option Explicit

'===============================================================================
' MODULE : modMiseEnPageFinale
' VERSION : MISEENPAGE-2E
'
' Corrections :
' - aucun paragraphe vide entre les paragraphes médicaux ;
' - séparation visuelle contrôlée par 6 pt après chaque paragraphe ;
' - interligne simple à l'intérieur des paragraphes ;
' - formule de politesse solidaire de la signature ;
' - distinction stricte entre l'en-tête du cabinet et la signature ;
' - restauration explicite de l'en-tête dans chaque lettre complémentaire.
'
' Ce module ne modifie pas le contenu médical.
'===============================================================================

Private Const MPF_ESPACE_APRES_PARAGRAPHE_PT As Single = 6

Public Function MPF_NormaliserRetoursTexte( _
    ByVal texte As String) As String

    Dim lignes() As String
    Dim resultat As String
    Dim ligne As String
    Dim i As Long
    Dim aCommence As Boolean

    'Uniformiser tous les retours.
    texte = Replace(texte, vbCrLf, vbLf)
    texte = Replace(texte, vbCr, vbLf)

    lignes = Split(texte, vbLf)

    resultat = ""
    aCommence = False

    For i = LBound(lignes) To UBound(lignes)

        ligne = CStr(lignes(i))
        ligne = Trim$(Replace(ligne, Chr$(160), " "))

        'MISEENPAGE-2 :
        'les lignes vides produites par l'API sont supprimées.
        'Deux paragraphes médicaux sont séparés par UN seul paragraphe Word,
        'puis l'espacement visuel est réglé à 6 pt.
        If ligne <> "" Then

            If aCommence Then
                resultat = resultat & vbCr
            End If

            resultat = resultat & ligne
            aCommence = True

        End If

    Next i

    MPF_NormaliserRetoursTexte = resultat

End Function

Public Sub MPF_AppliquerMiseEnPageCourrier( _
    ByVal doc As Document)

    Dim rngCorps As Range
    Dim rngPremierPatient As Range
    Dim p As Paragraph

    If doc Is Nothing Then Exit Sub

    ' Le nouveau papier conserve la presentation relevee sur l annexe.
    ' Ce chemin est strict : une configuration incomplete doit interrompre
    ' la correction, et ne pas etre masquee par le gestionnaire historique.
    If MPF_UtilisePresentationAnnexe(doc) Then
        MPF_AppliquerPresentationAnnexe doc
        Exit Sub
    End If

    On Error GoTo fin

    'Supprimer d'éventuels paragraphes vides résiduels qui auraient été
    'conservés par Word pendant le remplacement du corps.
    MPF_SupprimerParagraphesVidesDansCorps doc

    If LocaliserCorpsCourrier( _
        doc, _
        rngCorps, _
        rngPremierPatient) Then

        For Each p In rngCorps.Paragraphs

            With p.Format
                .SpaceBefore = 0
                .SpaceAfter = MPF_ESPACE_APRES_PARAGRAPHE_PT
                .LineSpacingRule = wdLineSpaceSingle
                .KeepTogether = False
                .KeepWithNext = False
                .WidowControl = True
            End With

        Next p

    End If

    'Réduire spécifiquement l'espace entre le dernier paragraphe médical
    'et la formule de politesse.
    MPF_CompacterAvantPolitesse doc

    MPF_SecuriserSignatureCourrier doc

fin:
End Sub

Private Function MPF_UtilisePresentationAnnexe(ByVal doc As Document) As Boolean
    Dim variable As Variable
    For Each variable In doc.Variables
        If StrComp(variable.Name, "PresentationCorpsU2", vbTextCompare) = 0 Then
            If variable.Value <> "annexe-v1" Then
                Err.Raise vbObjectError + 979, "Presentation du courrier", _
                    "Profil PresentationCorpsU2 inconnu : " & variable.Value
            End If
            MPF_UtilisePresentationAnnexe = True
            Exit Function
        End If
    Next variable
End Function

Private Sub MPF_AppliquerPresentationAnnexe(ByVal doc As Document)
    Dim corps As Range, p As Paragraph, styleCorps As Style

    If Not doc.Bookmarks.Exists("CORPS") Then
        Err.Raise vbObjectError + 979, "Presentation du courrier", _
            "Presentation annexe-v1 : signet CORPS absent."
    End If
    Set corps = doc.Bookmarks("CORPS").Range.Duplicate
    If corps.StoryType <> wdMainTextStory Or corps.Start = corps.End Then
        Err.Raise vbObjectError + 979, "Presentation du courrier", _
            "Presentation annexe-v1 : signet CORPS vide ou hors du courrier principal."
    End If
    If doc.Bookmarks.Exists("PR_DEBUT_DEMANDES") Then
        If corps.End > doc.Bookmarks("PR_DEBUT_DEMANDES").Range.Start Then
            Err.Raise vbObjectError + 979, "Presentation du courrier", _
                "Presentation annexe-v1 : le signet CORPS empiète sur les annexes."
        End If
    End If
    On Error Resume Next
    Set styleCorps = doc.Styles("CabinetCorpsU2")
    On Error GoTo 0
    If styleCorps Is Nothing Then
        Err.Raise vbObjectError + 979, "Presentation du courrier", _
            "Presentation annexe-v1 : style CabinetCorpsU2 absent."
    End If
    If styleCorps.Type <> wdStyleTypeParagraph Then
        Err.Raise vbObjectError + 979, "Presentation du courrier", _
            "Presentation annexe-v1 : CabinetCorpsU2 doit etre un style de paragraphe."
    End If

    ' Valider toute la plage avant modification : aucun paragraphe partage
    ' avec l appel, la politesse ou la signature ne doit etre reformate.
    For Each p In corps.Paragraphs
        If p.Range.Start < corps.Start Or p.Range.End - 1 > corps.End Then
            Err.Raise vbObjectError + 979, "Presentation du courrier", _
                "Presentation annexe-v1 : les limites du signet CORPS doivent suivre les paragraphes."
        End If
    Next p

    For Each p In corps.Paragraphs
        ' Duplicate conserve notamment SpaceBeforeAuto/SpaceAfterAuto,
        ' les retraits, l interligne et les tabulations du modele.
        p.Range.ParagraphFormat = styleCorps.ParagraphFormat.Duplicate
        p.Range.Style = styleCorps
        ' Aucun Font.Reset ni remplacement global de Font : le gras et
        ' l italique explicites du texte medical doivent etre conserves.
    Next p
    ' Ne pas compacter la politesse ni securiser ici les signatures : ces
    ' operations historiques modifieraient les zones hors du signet CORPS.
End Sub

Private Sub MPF_CompacterAvantPolitesse( _
    ByVal doc As Document)

    Dim i As Long
    Dim idxPolitesse As Long
    Dim idxPrecedentNonVide As Long
    Dim texte As String
    Dim j As Long

    On Error GoTo fin

    If doc Is Nothing Then Exit Sub

    idxPolitesse = 0

    'Repérer la première formule de politesse située après le corps médical.
    For i = 1 To doc.Paragraphs.Count

        texte = LCase$(MPF_TexteParagraphe(doc.Paragraphs(i)))

        If MPF_EstFormulePolitesse(texte) Then
            idxPolitesse = i
            Exit For
        End If

    Next i

    If idxPolitesse = 0 Then Exit Sub

    'Supprimer tous les paragraphes vides immédiatement placés
    'avant la formule de politesse.
    For i = idxPolitesse - 1 To 1 Step -1

        If MPF_TexteParagraphe(doc.Paragraphs(i)) = "" Then
            doc.Paragraphs(i).Range.Delete
            idxPolitesse = idxPolitesse - 1
        Else
            Exit For
        End If

    Next i

    'Repérer le dernier paragraphe médical non vide.
    idxPrecedentNonVide = 0

    For j = idxPolitesse - 1 To 1 Step -1

        If MPF_TexteParagraphe(doc.Paragraphs(j)) <> "" Then
            idxPrecedentNonVide = j
            Exit For
        End If

    Next j

    'Aucun espace Word ajouté après le dernier paragraphe médical.
    If idxPrecedentNonVide > 0 Then

        With doc.Paragraphs(idxPrecedentNonVide).Format
            .SpaceAfter = 0
            .KeepWithNext = False
        End With

    End If

    'La formule de politesse commence immédiatement après le retour
    'de paragraphe normal, sans espace supplémentaire.
    With doc.Paragraphs(idxPolitesse).Format
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
    End With

fin:
End Sub

Private Function MPF_EstFormulePolitesse( _
    ByVal texteMinuscule As String) As Boolean

    Dim t As String

    t = Trim$(texteMinuscule)

    MPF_EstFormulePolitesse = False

    If t = "" Then Exit Function

    If InStr(1, t, "merci de votre confiance", vbTextCompare) > 0 Then
        MPF_EstFormulePolitesse = True
        Exit Function
    End If

    If InStr(1, t, "merci de ta confiance", vbTextCompare) > 0 Then
        MPF_EstFormulePolitesse = True
        Exit Function
    End If

    If InStr(1, t, "bien cordialement", vbTextCompare) > 0 Then
        MPF_EstFormulePolitesse = True
        Exit Function
    End If

    If InStr(1, t, "cordialement", vbTextCompare) > 0 Then
        MPF_EstFormulePolitesse = True
        Exit Function
    End If

    If InStr(1, t, "amitiés", vbTextCompare) > 0 _
    Or InStr(1, t, "amities", vbTextCompare) > 0 Then

        MPF_EstFormulePolitesse = True
        Exit Function

    End If

    If InStr(1, t, "bien amicalement", vbTextCompare) > 0 Then
        MPF_EstFormulePolitesse = True
        Exit Function
    End If

    If Left$(t, 9) = "je vous " _
    Or Left$(t, 10) = "je te prie" Then

        MPF_EstFormulePolitesse = True

    End If

End Function

Private Sub MPF_SupprimerParagraphesVidesDansCorps( _
    ByVal doc As Document)

    Dim rngCorps As Range
    Dim rngPremierPatient As Range
    Dim i As Long
    Dim texte As String

    On Error GoTo fin

    If Not LocaliserCorpsCourrier( _
        doc, _
        rngCorps, _
        rngPremierPatient) Then

        Exit Sub

    End If

    For i = rngCorps.Paragraphs.Count To 1 Step -1

        texte = MPF_TexteParagraphe(rngCorps.Paragraphs(i))

        If texte = "" Then
            rngCorps.Paragraphs(i).Range.Delete
        End If

    Next i

fin:
End Sub

Public Sub MPF_GarantirEnteteDepuisSource( _
    ByVal docSource As Document, _
    ByVal docCible As Document)

    Dim rngSource As Range
    Dim rngCible As Range

    On Error GoTo fin

    If docSource Is Nothing Then Exit Sub
    If docCible Is Nothing Then Exit Sub

    'MISEENPAGE-2D :
    'ne plus rechercher l'en-tête avec "Docteur Olivier MANDAGOUT",
    'car cette chaîne existe aussi dans la signature.
    '
    'L'en-tête est défini de façon positionnelle :
    'du début du document jusqu'à la PREMIÈRE ligne contenant le téléphone,
    'située avant la date du courrier.
    '
    'Cela ne peut pas sélectionner la signature de fin de lettre.
    Set rngSource = MPF_PlageEnteteParTelephone(docSource)

    If rngSource Is Nothing Then Exit Sub

    Set rngCible = MPF_PlageEnteteParTelephone(docCible)

    If Not rngCible Is Nothing Then

        'Remplacer uniquement l'en-tête existant.
        rngCible.FormattedText = rngSource.FormattedText

    Else

        'Si l'en-tête n'existe plus dans la copie temporaire,
        'le réinsérer au tout début du document.
        Set rngCible = docCible.Range( _
            Start:=docCible.Content.Start, _
            End:=docCible.Content.Start)

        rngCible.FormattedText = rngSource.FormattedText

    End If

fin:
End Sub

Private Function MPF_PlageEnteteParTelephone( _
    ByVal doc As Document) As Range

    Dim i As Long
    Dim idxDate As Long
    Dim idxTelephone As Long
    Dim texte As String
    Dim r As Range

    Set MPF_PlageEnteteParTelephone = Nothing

    If doc Is Nothing Then Exit Function

    idxDate = 0
    idxTelephone = 0

    'Repérer d'abord la date afin d'exclure tout ce qui se trouve
    'dans le corps ou la signature.
    For i = 1 To doc.Paragraphs.Count

        texte = MPF_NormaliserTexte( _
            MPF_TexteParagraphe(doc.Paragraphs(i)))

        If (InStr(1, texte, _
            "BEAUMONT SUR OISE", vbTextCompare) > 0 _
        Or InStr(1, texte, _
            "BEAUMONT-SUR-OISE", vbTextCompare) > 0) _
        And InStr(1, " " & texte & " ", _
            " LE ", vbTextCompare) > 0 Then

            idxDate = i
            Exit For

        End If

    Next i

    If idxDate = 0 Then
        'Sécurité : on ne copie rien si la structure du courrier
        'ne permet pas de borner l'en-tête sans ambiguïté.
        Exit Function
    End If

    'Chercher le PREMIER téléphone uniquement avant la date.
    For i = 1 To idxDate - 1

        texte = MPF_NormaliserTexte( _
            MPF_TexteParagraphe(doc.Paragraphs(i)))

        If InStr(1, texte, "TEL", vbTextCompare) > 0 _
        Or InStr(1, texte, "TELEPHONE", vbTextCompare) > 0 Then

            idxTelephone = i
            Exit For

        End If

    Next i

    If idxTelephone = 0 Then Exit Function

    'Copie exacte depuis le début du document jusqu'à la fin de la ligne
    'du téléphone. FormattedText conservera toutes les polices et tailles.
    Set r = doc.Range( _
        Start:=doc.Content.Start, _
        End:=doc.Paragraphs(idxTelephone).Range.End)

    Set MPF_PlageEnteteParTelephone = r

End Function

Public Sub MPF_ReappliquerEnteteApresFusion( _
    ByVal docPrincipal As Document, _
    ByVal rngLettreAjoutee As Range)

    Dim rngSource As Range
    Dim rngCible As Range

    On Error GoTo fin

    If docPrincipal Is Nothing Then Exit Sub
    If rngLettreAjoutee Is Nothing Then Exit Sub

    'En-tête de référence : celui du courrier principal, déjà affiché correctement.
    Set rngSource = MPF_PlageEnteteParTelephone(docPrincipal)

    If rngSource Is Nothing Then Exit Sub

    'En-tête cible : uniquement dans la nouvelle lettre qui vient d'être collée.
    Set rngCible = _
        MPF_PlageEnteteDansRangeParTelephone(rngLettreAjoutee)

    If rngCible Is Nothing Then Exit Sub

    'Word peut remapper les styles pendant une fusion entre documents.
    'On réapplique donc APRÈS la fusion les attributs de police caractère par
    'caractère, ce qui transforme la mise en forme en format direct et empêche
    'un style Normal/Arial de remplacer la police du modèle principal.
    MPF_ReappliquerFormatEnteteExact _
        rngSource, _
        rngCible

fin:
End Sub

Private Function MPF_PlageEnteteDansRangeParTelephone( _
    ByVal rngLettre As Range) As Range

    Dim i As Long
    Dim idxDate As Long
    Dim idxTelephone As Long
    Dim texte As String
    Dim r As Range

    Set MPF_PlageEnteteDansRangeParTelephone = Nothing

    If rngLettre Is Nothing Then Exit Function
    If rngLettre.Paragraphs.Count = 0 Then Exit Function

    idxDate = 0
    idxTelephone = 0

    'Repérer la date à l'intérieur de la lettre ajoutée.
    For i = 1 To rngLettre.Paragraphs.Count

        texte = MPF_NormaliserTexte( _
            MPF_TexteParagraphe(rngLettre.Paragraphs(i)))

        If (InStr(1, texte, _
            "BEAUMONT SUR OISE", vbTextCompare) > 0 _
        Or InStr(1, texte, _
            "BEAUMONT-SUR-OISE", vbTextCompare) > 0) _
        And InStr(1, " " & texte & " ", _
            " LE ", vbTextCompare) > 0 Then

            idxDate = i
            Exit For

        End If

    Next i

    If idxDate = 0 Then Exit Function

    'Le téléphone de l'en-tête doit impérativement être avant la date.
    For i = 1 To idxDate - 1

        texte = MPF_NormaliserTexte( _
            MPF_TexteParagraphe(rngLettre.Paragraphs(i)))

        If InStr(1, texte, "TEL", vbTextCompare) > 0 _
        Or InStr(1, texte, "TELEPHONE", vbTextCompare) > 0 Then

            idxTelephone = i
            Exit For

        End If

    Next i

    If idxTelephone = 0 Then Exit Function

    Set r = rngLettre.Document.Range( _
        Start:=rngLettre.Paragraphs(1).Range.Start, _
        End:=rngLettre.Paragraphs(idxTelephone).Range.End)

    Set MPF_PlageEnteteDansRangeParTelephone = r

End Function

Private Sub MPF_ReappliquerFormatEnteteExact( _
    ByVal rngSource As Range, _
    ByVal rngCible As Range)

    Dim i As Long
    Dim n As Long
    Dim p As Long
    Dim np As Long

    On Error GoTo fin

    If rngSource Is Nothing Then Exit Sub
    If rngCible Is Nothing Then Exit Sub

    n = rngSource.Characters.Count

    If rngCible.Characters.Count < n Then
        n = rngCible.Characters.Count
    End If

    'Police et attributs caractère par caractère.
    For i = 1 To n

        With rngCible.Characters(i).Font

            .Name = rngSource.Characters(i).Font.Name
            .NameAscii = rngSource.Characters(i).Font.NameAscii
            .NameFarEast = rngSource.Characters(i).Font.NameFarEast
            .NameOther = rngSource.Characters(i).Font.NameOther

            .Size = rngSource.Characters(i).Font.Size
            .Bold = rngSource.Characters(i).Font.Bold
            .Italic = rngSource.Characters(i).Font.Italic
            .Underline = rngSource.Characters(i).Font.Underline
            .StrikeThrough = rngSource.Characters(i).Font.StrikeThrough
            .SmallCaps = rngSource.Characters(i).Font.SmallCaps
            .AllCaps = rngSource.Characters(i).Font.AllCaps
            .Superscript = rngSource.Characters(i).Font.Superscript
            .Subscript = rngSource.Characters(i).Font.Subscript

        End With

    Next i

    'Mise en forme de paragraphe.
    np = rngSource.Paragraphs.Count

    If rngCible.Paragraphs.Count < np Then
        np = rngCible.Paragraphs.Count
    End If

    For p = 1 To np

        With rngCible.Paragraphs(p).Format

            .Alignment = _
                rngSource.Paragraphs(p).Format.Alignment

            .LeftIndent = _
                rngSource.Paragraphs(p).Format.LeftIndent

            .RightIndent = _
                rngSource.Paragraphs(p).Format.RightIndent

            .FirstLineIndent = _
                rngSource.Paragraphs(p).Format.FirstLineIndent

            .SpaceBefore = _
                rngSource.Paragraphs(p).Format.SpaceBefore

            .SpaceAfter = _
                rngSource.Paragraphs(p).Format.SpaceAfter

            .LineSpacingRule = _
                rngSource.Paragraphs(p).Format.LineSpacingRule

            .LineSpacing = _
                rngSource.Paragraphs(p).Format.LineSpacing

        End With

    Next p

fin:
End Sub

Public Sub MPF_SecuriserToutesSignatures( _
    ByVal doc As Document)

    Dim i As Long
    Dim texte As String

    On Error GoTo fin

    If doc Is Nothing Then Exit Sub

    For i = doc.Paragraphs.Count To 1 Step -1

        texte = MPF_NormaliserTexte( _
            MPF_TexteParagraphe(doc.Paragraphs(i)))

        If MPF_EstSignature(texte) Then

            'Ne jamais confondre l'en-tête du cabinet avec la signature.
            If Not MPF_EstEnteteAIndex(doc, i) Then
                MPF_SecuriserSignatureIndex doc, i
            End If

        End If

    Next i

fin:
End Sub

Private Sub MPF_SecuriserSignatureCourrier( _
    ByVal doc As Document)

    Dim i As Long
    Dim texte As String

    If doc Is Nothing Then Exit Sub

    For i = doc.Paragraphs.Count To 1 Step -1

        texte = MPF_NormaliserTexte( _
            MPF_TexteParagraphe(doc.Paragraphs(i)))

        If MPF_EstSignature(texte) Then

            If Not MPF_EstEnteteAIndex(doc, i) Then

                MPF_SecuriserSignatureIndex doc, i
                Exit Sub

            End If

        End If

    Next i

End Sub

Private Function MPF_EstEnteteAIndex( _
    ByVal doc As Document, _
    ByVal idx As Long) As Boolean

    Dim j As Long
    Dim limite As Long
    Dim texte As String
    Dim zone As String

    MPF_EstEnteteAIndex = False

    If doc Is Nothing Then Exit Function
    If idx < 1 Or idx > doc.Paragraphs.Count Then Exit Function

    texte = MPF_NormaliserTexte( _
        MPF_TexteParagraphe(doc.Paragraphs(idx)))

    If InStr(1, texte, _
        "DOCTEUR OLIVIER MANDAGOUT", _
        vbTextCompare) = 0 Then

        Exit Function

    End If

    'Dans le modèle habituel, tout l'en-tête est dans un seul paragraphe
    'avec des retours manuels.
    If MPF_TexteContientMarqueurEntete(texte) Then

        MPF_EstEnteteAIndex = True
        Exit Function

    End If

    'Sécurité pour les anciens modèles où l'en-tête peut occuper
    'plusieurs paragraphes successifs.
    limite = idx + 5

    If limite > doc.Paragraphs.Count Then
        limite = doc.Paragraphs.Count
    End If

    zone = ""

    For j = idx To limite

        zone = zone & " " & _
            MPF_NormaliserTexte( _
                MPF_TexteParagraphe(doc.Paragraphs(j)))

    Next j

    If MPF_TexteContientMarqueurEntete(zone) Then
        MPF_EstEnteteAIndex = True
    End If

End Function

Private Function MPF_TexteContientMarqueurEntete( _
    ByVal texte As String) As Boolean

    MPF_TexteContientMarqueurEntete = False

    If InStr(1, texte, _
        "ANCIEN INTERNE", _
        vbTextCompare) > 0 Then

        MPF_TexteContientMarqueurEntete = True
        Exit Function

    End If

    If InStr(1, texte, _
        "MALADIES CARDIO-VASCULAIRES", _
        vbTextCompare) > 0 Then

        MPF_TexteContientMarqueurEntete = True
        Exit Function

    End If

    If InStr(1, texte, _
        "RUE DE LA LIBERATION", _
        vbTextCompare) > 0 Then

        MPF_TexteContientMarqueurEntete = True
        Exit Function

    End If

    If InStr(1, texte, _
        "TEL", _
        vbTextCompare) > 0 Then

        MPF_TexteContientMarqueurEntete = True

    End If

End Function

Private Sub MPF_SecuriserSignatureIndex( _
    ByVal doc As Document, _
    ByVal idxSignature As Long)

    Dim idxPrecedentNonVide As Long
    Dim j As Long

    If idxSignature <= 1 Then Exit Sub

    idxPrecedentNonVide = 0

    For j = idxSignature - 1 To 1 Step -1

        If Trim$(MPF_TexteParagraphe( _
            doc.Paragraphs(j))) <> "" Then

            idxPrecedentNonVide = j
            Exit For

        End If

    Next j

    If idxPrecedentNonVide > 0 Then

        'La formule de politesse reste avec la signature.
        For j = idxPrecedentNonVide To idxSignature - 1

            ' Sans politesse, le paragraphe precedent peut etre medical.
            ' Le nouveau profil doit survivre aussi a la passe globale
            ' executee apres l ajout des annexes. Ne proteger ici que le
            ' CORPS principal marque, jamais les paragraphes des annexes.
            If Not MPF_ParagrapheCorpsPresentationAnnexe(doc, doc.Paragraphs(j)) Then
                With doc.Paragraphs(j).Format
                    .KeepWithNext = True
                    .KeepTogether = True
                    .WidowControl = True
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .LineSpacingRule = wdLineSpaceSingle
                End With
            End If

        Next j

    End If

    With doc.Paragraphs(idxSignature).Format
        .KeepTogether = True
        .WidowControl = True
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
    End With

    If idxSignature < doc.Paragraphs.Count Then

        On Error Resume Next

        If doc.Paragraphs( _
            idxSignature + 1).Range.InlineShapes.Count > 0 Then

            doc.Paragraphs(idxSignature).Format.KeepWithNext = True

            doc.Paragraphs( _
                idxSignature + 1).Format.KeepTogether = True

        End If

        On Error GoTo 0

    End If

End Sub

Private Function MPF_ParagrapheCorpsPresentationAnnexe( _
    ByVal doc As Document, ByVal p As Paragraph) As Boolean

    Dim corps As Range
    If Not MPF_UtilisePresentationAnnexe(doc) Then Exit Function
    If Not doc.Bookmarks.Exists("CORPS") Then Exit Function
    Set corps = doc.Bookmarks("CORPS").Range
    MPF_ParagrapheCorpsPresentationAnnexe = _
        corps.StoryType = wdMainTextStory And p.Range.StoryType = wdMainTextStory And _
        p.Range.Start >= corps.Start And p.Range.End - 1 <= corps.End
End Function

Private Function MPF_EstSignature( _
    ByVal texteNormalise As String) As Boolean

    MPF_EstSignature = False

    If InStr(1, texteNormalise, _
        "DOCTEUR OLIVIER MANDAGOUT", _
        vbTextCompare) > 0 Then

        MPF_EstSignature = True
        Exit Function

    End If

    If InStr(1, texteNormalise, _
        "DR OLIVIER MANDAGOUT", _
        vbTextCompare) > 0 Then

        MPF_EstSignature = True

    End If

End Function

Private Function MPF_TexteParagraphe( _
    ByVal p As Paragraph) As String

    Dim s As String

    s = p.Range.Text
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr$(7), "")
    s = Replace(s, Chr$(11), " ")
    s = Replace(s, Chr$(12), "")
    s = Replace(s, Chr$(160), " ")

    MPF_TexteParagraphe = Trim$(s)

End Function

Private Function MPF_NormaliserTexte( _
    ByVal s As String) As String

    s = UCase$(s)

    s = Replace(s, "É", "E")
    s = Replace(s, "È", "E")
    s = Replace(s, "Ê", "E")
    s = Replace(s, "Ë", "E")
    s = Replace(s, "À", "A")
    s = Replace(s, "Â", "A")
    s = Replace(s, "Ä", "A")
    s = Replace(s, "Î", "I")
    s = Replace(s, "Ï", "I")
    s = Replace(s, "Ô", "O")
    s = Replace(s, "Ö", "O")
    s = Replace(s, "Ù", "U")
    s = Replace(s, "Û", "U")
    s = Replace(s, "Ü", "U")
    s = Replace(s, "Ç", "C")
    s = Replace(s, "Œ", "OE")

    MPF_NormaliserTexte = Trim$(s)

End Function
