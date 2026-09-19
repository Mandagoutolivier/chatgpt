Attribute VB_Name = "modRecettePresentationAnnexe"
Option Explicit
' Recette autonome Word : documents fictifs en memoire, aucun NAS ni API.
Private mNombre As Long

Private Sub Exiger(ByVal condition As Boolean, ByVal description As String)
    If Not condition Then Err.Raise vbObjectError + 1194, "Recette presentation", description
    mNombre = mNombre + 1
End Sub

Private Function FormatParagraphe(ByVal format As ParagraphFormat) As String
    Dim valeur As String, tabulation As TabStop
    With format
        valeur = CStr(.Alignment) & "|" & CStr(.LeftIndent) & "|" & CStr(.RightIndent) & _
            "|" & CStr(.FirstLineIndent) & "|" & CStr(.SpaceBefore) & "|" & CStr(.SpaceAfter) & _
            "|" & CStr(.SpaceBeforeAuto) & "|" & CStr(.SpaceAfterAuto) & _
            "|" & CStr(.LineSpacingRule) & "|" & CStr(.LineSpacing) & _
            "|" & CStr(.KeepTogether) & "|" & CStr(.KeepWithNext) & "|" & CStr(.WidowControl) & _
            "|" & CStr(.PageBreakBefore)
        For Each tabulation In .TabStops
            ' Un Range peut enumerer les arrets par defaut, contrairement
            ' au ParagraphFormat du style. Seuls les arrets personnalises
            ' appartiennent au format que cette recette doit comparer.
            If tabulation.CustomTab Then
                valeur = valeur & "|TAB:" & CStr(tabulation.Position) & ":" & _
                    CStr(tabulation.Alignment) & ":" & CStr(tabulation.Leader)
            End If
        Next tabulation
    End With
    FormatParagraphe = valeur
End Function

Private Function FormatZone(ByVal zone As Range) As String
    Dim valeur As String, p As Paragraph
    valeur = zone.Text & "|" & CStr(zone.Font.Name) & "|" & CStr(zone.Font.Size) & _
        "|" & CStr(zone.Font.Bold) & "|" & CStr(zone.Font.Italic) & "|" & CStr(zone.Font.Underline)
    For Each p In zone.Paragraphs
        valeur = valeur & "|" & FormatParagraphe(p.Format)
    Next p
    FormatZone = valeur
End Function

Private Function EtatDocument(ByVal doc As Document) As String
    EtatDocument = FormatZone(doc.Content) & "|HEADER|" & _
        FormatZone(doc.Sections(1).Headers(wdHeaderFooterPrimary).Range) & "|FOOTER|" & _
        FormatZone(doc.Sections(1).Footers(wdHeaderFooterPrimary).Range)
End Function

Private Function Marquer(ByVal doc As Document, ByVal index As Long, ByVal nom As String) As Range
    Dim zone As Range
    Set zone = doc.Paragraphs(index).Range.Duplicate
    doc.Bookmarks.Add nom, zone
    Set Marquer = zone
End Function

Private Function CreerFixture(ByVal avecProfil As Boolean) As Document
    Dim doc As Document, st As Style, corps As Range, zone As Range
    Dim i As Long, paragrapheLong As String
    Dim numero As Long, description As String
    On Error GoTo Echec
    paragrapheLong = "Texte de recette entierement fictif sans donnee clinique reelle. "
    paragrapheLong = paragrapheLong & paragrapheLong & paragrapheLong & paragrapheLong
    Set doc = Documents.Add
    doc.Content.Text = "Docteur Olivier MANDAGOUT" & Chr$(11) & "Tel : FICTIF" & vbCr & _
        "DESTINATAIRE FICTIF" & vbCr & "Beaumont-sur-Oise, le 19 septembre 2026" & vbCr & _
        "Cher Confrere," & vbCr & "Monsieur PATIENT FICTIF. " & paragrapheLong & vbCr & _
        "Examen clinique fictif. " & paragrapheLong & vbCr & _
        "Au total, conclusion fictive. " & paragrapheLong & vbCr & _
        "Amities." & vbCr & "Docteur Olivier MANDAGOUT" & vbCr & _
        "ANNEXE FICTIVE" & vbCr & "Contenu annexe a conserver, avec un format distinct." & vbCr
    With doc.Styles(wdStyleNormal)
        .Font.Name = "Calibri"
        .Font.Size = 11
        .ParagraphFormat.SpaceAfter = 9
        .NextParagraphStyle = wdStyleNormal
    End With
    For i = 1 To doc.Paragraphs.Count
        With doc.Paragraphs(i).Format
            .LeftIndent = 11 + i
            .FirstLineIndent = 3
            .SpaceBefore = 2 + i
            .SpaceAfter = 4 + i
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .KeepTogether = False
            .KeepWithNext = False
        End With
    Next i
    Set zone = Marquer(doc, 1, "FIXE_ENTETE")
    Set zone = Marquer(doc, 2, "FIXE_DESTINATAIRE")
    Set zone = Marquer(doc, 3, "FIXE_DATE")
    Set zone = Marquer(doc, 4, "FIXE_APPEL")
    Set zone = Marquer(doc, 8, "FIXE_POLITESSE")
    Set zone = Marquer(doc, 9, "FIXE_SIGNATURE")
    Set zone = Marquer(doc, 10, "FIXE_TITRE_ANNEXE")
    Set zone = Marquer(doc, 11, "FIXE_CORPS_ANNEXE")
    Set zone = doc.Paragraphs(10).Range.Duplicate
    zone.Collapse wdCollapseStart
    doc.Bookmarks.Add "PR_DEBUT_DEMANDES", zone
    doc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = "EN TETE DE SECTION FICTIF"
    doc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = "PIED DE PAGE FICTIF"
    Set corps = doc.Range(doc.Paragraphs(5).Range.Start, doc.Paragraphs(7).Range.End)
    doc.Bookmarks.Add "CORPS", corps
    doc.Variables.Add "PatientID", "PATIENT-FICTIF-PRESENTATION"
    Set st = doc.Styles.Add("CabinetCorpsU2", wdStyleTypeParagraph)
    st.AutomaticallyUpdate = False
    st.Font.Name = "Times New Roman"
    st.Font.Size = 10
    st.Font.Bold = False
    st.Font.Italic = False
    st.NextParagraphStyle = st
    With st.ParagraphFormat
        .Alignment = wdAlignParagraphLeft
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 85.05
        .SpaceBefore = 5
        .SpaceAfter = 5
        .SpaceBeforeAuto = True
        .SpaceAfterAuto = True
        .LineSpacingRule = wdLineSpaceSingle
        .KeepTogether = False
        .KeepWithNext = False
        .WidowControl = True
        .PageBreakBefore = False
        .TabStops.ClearAll
        .TabStops.Add Position:=110, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    End With
    corps.Style = st
    ' Simuler les surcharges historiques que la finalisation doit remplacer.
    With corps.ParagraphFormat
        .FirstLineIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 6
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .TabStops.ClearAll
    End With
    corps.Font.Bold = False
    corps.Font.Italic = False
    Set zone = doc.Range(corps.Start, corps.Start + 8)
    zone.Font.Bold = True
    doc.Bookmarks.Add "GRAS_FICTIF", zone
    Set zone = doc.Range(doc.Paragraphs(6).Range.Start, doc.Paragraphs(6).Range.Start + 6)
    zone.Font.Italic = True
    zone.Font.Underline = wdUnderlineSingle
    doc.Bookmarks.Add "ITALIQUE_FICTIF", zone
    Set zone = doc.Range(doc.Paragraphs(7).Range.Start + 10, doc.Paragraphs(7).Range.Start + 20)
    doc.Bookmarks.Add "NORMAL_FICTIF", zone
    If avecProfil Then doc.Variables.Add "PresentationCorpsU2", "annexe-v1"
    Set CreerFixture = doc
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "CreerFixturePresentation", description
End Function

Private Sub VerifierProfil()
    Dim doc As Document, zone As Range, p As Paragraph, noms As Variant, nom As Variant
    Dim avant As Object, attendu As String, contenu As String, normal As String, etat As String
    Dim numero As Long, description As String
    On Error GoTo Echec
    Set doc = CreerFixture(True)
    contenu = doc.Content.Text
    normal = FormatParagraphe(doc.Styles(wdStyleNormal).ParagraphFormat) & "|" & _
        CStr(doc.Styles(wdStyleNormal).Font.Name) & "|" & CStr(doc.Styles(wdStyleNormal).Font.Size) & _
        "|" & CStr(doc.Styles(wdStyleNormal).NextParagraphStyle)
    attendu = FormatParagraphe(doc.Styles("CabinetCorpsU2").ParagraphFormat)
    Set avant = CreateObject("Scripting.Dictionary")
    noms = Array("FIXE_ENTETE", "FIXE_DESTINATAIRE", "FIXE_DATE", "FIXE_APPEL", _
        "FIXE_POLITESSE", "FIXE_SIGNATURE", "FIXE_TITRE_ANNEXE", "FIXE_CORPS_ANNEXE")
    For Each nom In noms
        avant(CStr(nom)) = FormatZone(doc.Bookmarks(CStr(nom)).Range)
    Next nom
    avant("HEADER") = FormatZone(doc.Sections(1).Headers(wdHeaderFooterPrimary).Range)
    avant("FOOTER") = FormatZone(doc.Sections(1).Footers(wdHeaderFooterPrimary).Range)
    Exiger FormatParagraphe(doc.Bookmarks("CORPS").Range.Paragraphs(1).Format) <> attendu, _
        "Fixture initiale distincte du profil attendu"
    modMiseEnPageFinale.MPF_AppliquerMiseEnPageCourrier doc
    Exiger doc.Content.Text = contenu, "Texte medical et paragraphes conserves"
    Exiger doc.Bookmarks("CORPS").Range.Paragraphs.Count = 3, "Trois longs paragraphes conserves"
    For Each p In doc.Bookmarks("CORPS").Range.Paragraphs
        Exiger FormatParagraphe(p.Format) = attendu, "Format exact du profil : retraits, espacements auto et tabulations"
        Exiger CStr(p.Range.Style) = "CabinetCorpsU2", "Style reserve au corps"
        Exiger p.Range.Font.Name = "Times New Roman" And p.Range.Font.Size = 10, "Police du profil conservee"
    Next p
    Exiger doc.Bookmarks("GRAS_FICTIF").Range.Font.Bold = True, "Gras explicite conserve"
    Exiger doc.Bookmarks("NORMAL_FICTIF").Range.Font.Bold = False, "Texte courant non mis en gras"
    Exiger doc.Bookmarks("ITALIQUE_FICTIF").Range.Font.Italic = True And _
        doc.Bookmarks("ITALIQUE_FICTIF").Range.Font.Underline = wdUnderlineSingle, "Italique et souligne conserves"
    For Each nom In noms
        Exiger FormatZone(doc.Bookmarks(CStr(nom)).Range) = avant(CStr(nom)), "Zone voisine inchangee : " & CStr(nom)
    Next nom
    Exiger FormatZone(doc.Sections(1).Headers(wdHeaderFooterPrimary).Range) = avant("HEADER"), "En tete de section conserve"
    Exiger FormatZone(doc.Sections(1).Footers(wdHeaderFooterPrimary).Range) = avant("FOOTER"), "Pied de page conserve"
    Exiger normal = FormatParagraphe(doc.Styles(wdStyleNormal).ParagraphFormat) & "|" & _
        CStr(doc.Styles(wdStyleNormal).Font.Name) & "|" & CStr(doc.Styles(wdStyleNormal).Font.Size) & _
        "|" & CStr(doc.Styles(wdStyleNormal).NextParagraphStyle), "Style Normal inchange"
    etat = EtatDocument(doc)
    modMiseEnPageFinale.MPF_AppliquerMiseEnPageCourrier doc
    Exiger EtatDocument(doc) = etat, "Deuxieme application sans changement"
    doc.Close wdDoNotSaveChanges
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "VerifierProfilPresentation", description
End Sub

Private Sub VerifierHistorique()
    Dim doc As Document, numero As Long, description As String
    On Error GoTo Echec
    Set doc = CreerFixture(False)
    doc.Paragraphs(5).Format.SpaceBefore = 17
    doc.Paragraphs(5).Format.SpaceAfter = 21
    modMiseEnPageFinale.MPF_AppliquerMiseEnPageCourrier doc
    With doc.Paragraphs(5).Format
        Exiger .SpaceBefore = 0 And .SpaceAfter = 6, "Sans marqueur : espacement historique conserve"
        Exiger .FirstLineIndent = 0, "Sans marqueur : aucun nouveau retrait impose"
        Exiger .LineSpacingRule = wdLineSpaceSingle, "Sans marqueur : interligne historique"
    End With
    Exiger doc.Paragraphs(7).Format.SpaceAfter = 0, "Sans marqueur : compactage historique avant politesse"
    doc.Close wdDoNotSaveChanges
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "VerifierHistoriquePresentation", description
End Sub

Private Sub VerifierFinDuCycle()
    Dim doc As Document, temoin As Document, zone As Range, p As Paragraph
    Dim attendu As String, contenu As String, annexe As String, annexeTemoin As String
    Dim numero As Long, description As String
    On Error GoTo Echec
    Set doc = CreerFixture(True)
    Set temoin = CreerFixture(False)
    ' Principal sans politesse : la signature suit directement le dernier
    ' paragraphe medical. L annexe conserve sa propre politesse/signature.
    doc.Bookmarks("FIXE_POLITESSE").Range.Delete
    temoin.Bookmarks("FIXE_POLITESSE").Range.Delete
    doc.Paragraphs.Last.Range.Text = "Amities." & vbCr & "Docteur Olivier MANDAGOUT" & vbCr
    temoin.Paragraphs.Last.Range.Text = "Amities." & vbCr & "Docteur Olivier MANDAGOUT" & vbCr
    contenu = doc.Content.Text
    modMiseEnPageFinale.MPF_AppliquerMiseEnPageCourrier doc
    attendu = FormatParagraphe(doc.Styles("CabinetCorpsU2").ParagraphFormat)
    modMiseEnPageFinale.MPF_SecuriserToutesSignatures doc
    modMiseEnPageFinale.MPF_SecuriserToutesSignatures temoin
    Exiger doc.Content.Text = contenu, "Passe finale de signature sans changement du texte"
    For Each p In doc.Bookmarks("CORPS").Range.Paragraphs
        Exiger FormatParagraphe(p.Format) = attendu, "Profil principal conserve apres la passe finale de signature"
    Next p
    Exiger doc.Bookmarks("FIXE_SIGNATURE").Range.Paragraphs(1).Format.KeepTogether = True, _
        "Protection de la signature principale conservee"
    Set zone = doc.Range(doc.Bookmarks("PR_DEBUT_DEMANDES").Range.Start, doc.Content.End)
    annexe = FormatZone(zone)
    Set zone = temoin.Range(temoin.Bookmarks("PR_DEBUT_DEMANDES").Range.Start, temoin.Content.End)
    annexeTemoin = FormatZone(zone)
    Exiger annexe = annexeTemoin, "Passe finale de l annexe identique au comportement historique"
    Exiger doc.Bookmarks("GRAS_FICTIF").Range.Font.Bold = True, "Gras conserve apres la passe finale"
    doc.Close wdDoNotSaveChanges
    temoin.Close wdDoNotSaveChanges
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    If Not temoin Is Nothing Then temoin.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "VerifierFinDuCyclePresentation", description
End Sub

Private Sub VerifierDoubleRetourFinal()
    Dim doc As Document, corps As Range, voisin As Range, p As Paragraph
    Dim texte As String, contenu As String, avantVoisin As String, attendu As String
    Dim debut As Long, fin As Long, numero As Long, description As String
    On Error GoTo Echec
    Set doc = CreerFixture(True)
    Set corps = doc.Bookmarks("CORPS").Range.Duplicate
    debut = corps.Start
    ' Meme preparation que RemplacerCorpsOriginalParTexte : les deux
    ' marques finales appartiennent a CORPS, dont un paragraphe vide.
    texte = modInsertion.GarantirRetourParagrapheFinal(corps.Text)
    corps.Text = texte
    fin = debut + Len(texte)
    doc.Bookmarks.Add "CORPS", doc.Range(debut, fin)
    Set corps = doc.Bookmarks("CORPS").Range.Duplicate
    Set voisin = doc.Range(fin, fin).Paragraphs(1).Range.Duplicate
    Exiger Right$(corps.Text, 2) = vbCr & vbCr, "Double retour final de production dans CORPS"
    Exiger corps.Paragraphs.Count = 4, "Trois paragraphes medicaux et un paragraphe vide final"
    Exiger corps.Paragraphs.Last.Range.Text = vbCr, "Dernier paragraphe du corps vide"
    Exiger voisin.Text = "Amities." & vbCr And voisin.Start = fin, "Politesse voisine distincte du corps"
    avantVoisin = FormatZone(voisin)
    contenu = doc.Content.Text
    attendu = FormatParagraphe(doc.Styles("CabinetCorpsU2").ParagraphFormat)
    modMiseEnPageFinale.MPF_AppliquerMiseEnPageCourrier doc
    Exiger doc.Content.Text = contenu, "Double retour final conserve sans suppression de texte"
    Exiger doc.Bookmarks("CORPS").Range.Start = debut And _
        doc.Bookmarks("CORPS").Range.End = fin, "Bornes CORPS conservees avec double retour"
    For Each p In doc.Bookmarks("CORPS").Range.Paragraphs
        Exiger FormatParagraphe(p.Format) = attendu, "Profil applique au corps et au paragraphe vide final"
    Next p
    Exiger FormatZone(voisin) = avantVoisin, "Format de la politesse apres double retour inchange"
    Exiger FormatParagraphe(voisin.Paragraphs(1).Format) <> attendu, "Paragraphe voisin non inclus dans le profil"
    modMiseEnPageFinale.MPF_SecuriserToutesSignatures doc
    For Each p In doc.Bookmarks("CORPS").Range.Paragraphs
        Exiger FormatParagraphe(p.Format) = attendu, "Corps avec double retour conserve apres passe finale"
    Next p
    doc.Close wdDoNotSaveChanges
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "VerifierDoubleRetourFinalPresentation", description
End Sub

Private Sub VerifierRefus(ByVal cas As String, ByVal message As String)
    Dim doc As Document, zone As Range, avant As String
    Dim refus As Long, details As String, numero As Long, description As String
    On Error GoTo Echec
    Set doc = CreerFixture(True)
    Select Case cas
        Case "sans-corps": doc.Bookmarks("CORPS").Delete
        Case "sans-style": doc.Styles("CabinetCorpsU2").Delete
        Case "profil-inconnu": doc.Variables("PresentationCorpsU2").Value = "annexe-inconnue"
        Case "annexe-incluse"
            Set zone = doc.Range(doc.Paragraphs(5).Range.Start, doc.Paragraphs(11).Range.End)
            doc.Bookmarks.Add "CORPS", zone
        Case "paragraphe-partage"
            Set zone = doc.Range(doc.Paragraphs(5).Range.Start + 1, doc.Paragraphs(7).Range.End)
            doc.Bookmarks.Add "CORPS", zone
        Case "corps-vide"
            Set zone = doc.Paragraphs(5).Range.Duplicate
            zone.Collapse wdCollapseStart
            doc.Bookmarks.Add "CORPS", zone
    End Select
    avant = EtatDocument(doc)
    On Error Resume Next
    modMiseEnPageFinale.MPF_AppliquerMiseEnPageCourrier doc
    refus = Err.Number: details = Err.Description: Err.Clear
    On Error GoTo Echec
    Exiger refus <> 0 And InStr(1, details, message, vbTextCompare) > 0, "Refus explicite : " & cas
    Exiger EtatDocument(doc) = avant, "Refus avant modification : " & cas
    doc.Close wdDoNotSaveChanges
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "VerifierRefusPresentation", description
End Sub

Public Function ExecuterPresentationAnnexe() As String
    Dim nombre As Long, numero As Long, description As String
    On Error GoTo Echec
    mNombre = 0
    nombre = Documents.Count
    VerifierProfil
    VerifierHistorique
    VerifierFinDuCycle
    VerifierDoubleRetourFinal
    VerifierRefus "sans-corps", "CORPS absent"
    VerifierRefus "sans-style", "CabinetCorpsU2 absent"
    VerifierRefus "profil-inconnu", "inconnu"
    VerifierRefus "annexe-incluse", "annexes"
    VerifierRefus "paragraphe-partage", "limites"
    VerifierRefus "corps-vide", "vide"
    Exiger Documents.Count = nombre, "Tous les documents fictifs fermes"
    ExecuterPresentationAnnexe = "{""reussis"":" & CStr(mNombre) & ",""echec"":false}"
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    ExecuterPresentationAnnexe = "{""reussis"":" & CStr(mNombre) & ",""echec"":true,""numero"":" & _
        CStr(numero) & ",""description"":" & modServiceNas.JsonValeur(description) & "}"
End Function
