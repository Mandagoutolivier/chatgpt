Attribute VB_Name = "modRecetteModeleCourrier"
Option Explicit
' Tests locaux du contrat du papier a lettres, sans service NAS ni API.
Private mNombre As Long

Private Sub Exiger(ByVal condition As Boolean, ByVal description As String)
    If Not condition Then Err.Raise vbObjectError + 1193, "Recette modele courrier", description
    mNombre = mNombre + 1
End Sub

Private Sub Configurer(ByVal racine As String, ByVal nom As String)
    modFichiers.EcrireTexteUTF8 racine & "\Config\config.ini", _
        "[COURRIER]" & vbCrLf & "Modele=" & nom & vbCrLf & "AppelAuto=0" & vbCrLf & "PolitesseAuto=0"
    modConfig.DefinirRacine racine
End Sub

Private Function MarquerParagraphe(ByVal doc As Document, ByVal index As Long, ByVal nom As String) As Range
    Dim zone As Range
    Set zone = doc.Paragraphs(index).Range.Duplicate
    zone.MoveEnd wdCharacter, -1
    doc.Bookmarks.Add nom, zone
    Set MarquerParagraphe = zone
End Function

Private Sub CreerFixture(ByVal chemin As String, ByVal avecDestinataire As Boolean, ByVal avecAppel As Boolean)
    Dim doc As Document, zone As Range, numero As Long, description As String
    On Error GoTo Echec
    Set doc = Documents.Add
    doc.Content.Text = "EN TETE FICTIF" & vbCr & vbCr & "DATE FICTIVE" & vbCr & vbCr & vbCr & "SIGNATURE FICTIVE" & vbCr
    With doc.Paragraphs(1).Format
        .Alignment = wdAlignParagraphCenter
        .LeftIndent = -12
        .SpaceBefore = 3
        .SpaceAfter = 7
    End With
    doc.Paragraphs(1).Range.Font.Bold = True
    doc.Paragraphs(6).Format.LeftIndent = 110
    doc.Paragraphs(6).Range.Font.Italic = True
    doc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = "PIED DE PAGE FICTIF"
    If avecDestinataire Then Set zone = MarquerParagraphe(doc, 2, "CORRESPONDANT")
    If avecAppel Then Set zone = MarquerParagraphe(doc, 4, "FORMULE_APPEL")
    doc.SaveAs2 chemin, wdFormatXMLTemplate
    doc.Close wdDoNotSaveChanges
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "CreerFixture", description
End Sub

Private Sub CreerFixtureModerne(ByVal chemin As String, ByVal complet As Boolean)
    Dim doc As Document, zone As Range, noms As Variant, i As Long, numero As Long, description As String
    On Error GoTo Echec
    noms = Array("EXPEDITEUR", "DESTINATAIRE", "DATELIEU", "CONCERNE", "APPEL", "CORPS", "SIGNATURE")
    Set doc = Documents.Add
    doc.Content.Text = "EXP" & vbCr & "DEST" & vbCr & "DATE" & vbCr & "PATIENT" & vbCr & "APPEL" & vbCr & "CORPS" & vbCr & "SIGNATURE" & vbCr
    For i = 0 To UBound(noms)
        If complet Or i > 0 Then Set zone = MarquerParagraphe(doc, i + 1, CStr(noms(i)))
    Next i
    doc.SaveAs2 chemin, wdFormatXMLTemplate
    doc.Close wdDoNotSaveChanges
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "CreerFixtureModerne", description
End Sub

Private Sub VerifierModerne(ByVal racine As String)
    Dim doc As Document, nombre As Long, numero As Long, description As String
    On Error GoTo Echec
    Configurer racine, "moderne"
    nombre = Documents.Count
    Set doc = modCourrier.CreerCourrierRapide()
    Exiger doc.Bookmarks.Exists("EXPEDITEUR") And doc.Bookmarks.Exists("CONCERNE") And _
        doc.Bookmarks.Exists("SIGNATURE") And doc.Bookmarks.Exists("DATELIEU"), "Contrat moderne complet accepte"
    Exiger Documents.Count = nombre + 1, "Un seul courrier moderne cree"
    doc.Close wdDoNotSaveChanges
    Exiger Documents.Count = nombre, "Courrier moderne de test ferme"
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "VerifierModerne", description
End Sub

Private Function FormatZone(ByVal zone As Range) As String
    Dim texte As String, p As Paragraph, arret As TabStop
    texte = CStr(zone.Font.Name) & "|" & CStr(zone.Font.Size) & "|" & CStr(zone.Font.Bold) & "|" & _
        CStr(zone.Font.Italic) & "|" & CStr(zone.Font.Underline) & "|" & CStr(zone.Font.Color) & "|" & _
        CStr(zone.Font.Hidden) & "|" & CStr(zone.Font.Spacing) & "|" & CStr(zone.Font.Position)
    For Each p In zone.Paragraphs
        With p.Format
            texte = texte & "|" & CStr(.Alignment) & "|" & CStr(.LeftIndent) & "|" & CStr(.RightIndent) & _
                "|" & CStr(.FirstLineIndent) & "|" & CStr(.SpaceBefore) & "|" & CStr(.SpaceAfter) & _
                "|" & CStr(.LineSpacingRule) & "|" & CStr(.LineSpacing) & "|" & CStr(.KeepWithNext) & _
                "|" & CStr(.KeepTogether) & "|" & CStr(.PageBreakBefore)
            For Each arret In .TabStops
                texte = texte & "|TAB:" & CStr(arret.Position) & ":" & CStr(arret.Alignment) & ":" & CStr(arret.Leader)
            Next arret
        End With
    Next p
    FormatZone = texte
End Function

Private Function ZonesStatiques(ByVal doc As Document, ByVal dates As Object) As Collection
    Dim resultat As New Collection, p As Paragraph, zone As Range, item As Object, story As Range, suite As Range, index As Long
    ' Paragraphes non vides hors des deux signets de saisie. Les champs de date
    ' sont testes separement car ils doivent etre recalcules et figes.
    For Each p In doc.Paragraphs
        index = index + 1
        Set zone = p.Range.Duplicate
        If Len(Trim$(Replace(zone.Text, vbCr, ""))) > 0 And Not dates.Exists(CStr(index)) Then
            If Not Chevauche(zone, doc.Bookmarks("CORRESPONDANT").Range) And _
               Not Chevauche(zone, doc.Bookmarks("FORMULE_APPEL").Range) Then
                Set item = CreateObject("Scripting.Dictionary")
                item("texte") = zone.Text: item("format") = FormatZone(zone)
                resultat.Add item
            End If
        End If
    Next p
    For Each story In doc.StoryRanges
        If story.StoryType <> wdMainTextStory Then
            Set suite = story
            Do While Not suite Is Nothing
                Set item = CreateObject("Scripting.Dictionary")
                item("texte") = suite.Text: item("format") = FormatZone(suite)
                resultat.Add item
                Set suite = suite.NextStoryRange
            Loop
        End If
    Next story
    Set ZonesStatiques = resultat
End Function

Private Function Chevauche(ByVal paragraphe As Range, ByVal signet As Range) As Boolean
    Chevauche = (paragraphe.StoryType = signet.StoryType And signet.Start >= paragraphe.Start And signet.Start < paragraphe.End)
End Function

Private Function MemeZone(ByVal a As Range, ByVal b As Range) As Boolean
    MemeZone = (a.StoryType = b.StoryType And a.Start = b.Start And a.End = b.End)
End Function

Private Sub VerifierModele(ByVal racine As String, ByVal nom As String, ByVal chemin As String)
    Dim original As Document, doc As Document, avant As Collection, apres As Collection, i As Long
    Dim styleNormal As String, formatDest As String, formatAppel As String, formatCorps As String
    Dim contenu As String, nombreDocuments As Long, securite As Long, numero As Long, description As String
    Dim dates As Object, p As Paragraph, index As Long, nombreDates As Long, dateInfo As Object, cle As Variant, patient As Object
    On Error GoTo Echec
    Configurer racine, nom
    nombreDocuments = Documents.Count
    securite = Application.AutomationSecurity
    Application.AutomationSecurity = 3
    Set original = Documents.Add(Template:=chemin)
    Application.AutomationSecurity = securite
    Exiger original.Bookmarks.Exists("CORRESPONDANT") And original.Bookmarks.Exists("FORMULE_APPEL"), "Deux signets historiques disponibles"
    Set dates = CreateObject("Scripting.Dictionary")
    For Each p In original.Paragraphs
        index = index + 1
        If p.Range.Fields.Count > 0 Then
            p.Range.Fields.Update
            Set dateInfo = CreateObject("Scripting.Dictionary")
            dateInfo("texte") = p.Range.Text
            dateInfo("format") = FormatZone(p.Range)
            Set dates(CStr(index)) = dateInfo
        End If
    Next p
    nombreDates = original.Fields.Count
    Set avant = ZonesStatiques(original, dates)
    formatDest = FormatZone(original.Bookmarks("CORRESPONDANT").Range)
    formatAppel = FormatZone(original.Bookmarks("FORMULE_APPEL").Range)
    styleNormal = CStr(original.Styles(wdStyleNormal).NextParagraphStyle)
    original.Close wdDoNotSaveChanges: Set original = Nothing
    Set patient = CreateObject("Scripting.Dictionary")
    patient("ID") = "PATIENT-FICTIF-MODELE": patient("Nom") = "FICTIF": patient("Prenom") = "MODELE"
    patient("DDN") = "01/01/1980": patient("Sexe") = "M"
    Set doc = modCourrier.CreerCourrierRapidePour(patient)
    Exiger doc.Variables("PatientID").Value = "PATIENT-FICTIF-MODELE", "PatientID conserve sans signet CONCERNE"
    Exiger Documents.Count = nombreDocuments + 1, "Un seul document cree"
    Exiger doc.Bookmarks.Exists("DESTINATAIRE") And doc.Bookmarks.Exists("APPEL") And doc.Bookmarks.Exists("CORPS"), "Trois signets applicatifs disponibles"
    Exiger Not doc.Bookmarks.Exists("EXPEDITEUR") And Not doc.Bookmarks.Exists("DATELIEU") And _
        Not doc.Bookmarks.Exists("CONCERNE") And Not doc.Bookmarks.Exists("SIGNATURE"), "Aucun faux signet autour des zones statiques"
    Set apres = ZonesStatiques(doc, dates)
    If nombreDates > 0 Then Exiger doc.Fields.Count = 0, "Champs de date figes"
    For Each cle In dates.Keys
        Set p = doc.Paragraphs(CLng(cle))
        Exiger p.Range.Text = dates(cle)("texte"), "Date identique au calcul Word pour le nouveau courrier"
        Exiger FormatZone(p.Range) = dates(cle)("format"), "Mise en forme de la date conservee"
    Next cle
    Exiger avant.Count = apres.Count, "Nombre de zones statiques conserve"
    For i = 1 To avant.Count
        Exiger avant(i)("texte") = apres(i)("texte"), "Texte statique conserve"
        Exiger avant(i)("format") = apres(i)("format"), "Mise en forme statique conservee"
    Next i
    Exiger formatDest = FormatZone(doc.Bookmarks("DESTINATAIRE").Range), "Format du destinataire conserve"
    Exiger formatAppel = FormatZone(doc.Bookmarks("APPEL").Range), "Format de la formule d appel conserve"
    Exiger CStr(doc.Styles(wdStyleNormal).NextParagraphStyle) = styleNormal, "Style Normal non modifie"
    Exiger CStr(doc.Bookmarks("CORPS").Range.Style) = "CabinetCorpsU2", _
        "Style reserve au corps (" & nom & ") : " & CStr(doc.Bookmarks("CORPS").Range.Style)
    Exiger CStr(doc.Styles("CabinetCorpsU2").NextParagraphStyle) = "CabinetCorpsU2", "Paragraphes de corps chaines entre eux"
    Exiger MemeZone(doc.Bookmarks("DESTINATAIRE").Range, doc.Bookmarks("CORRESPONDANT").Range), "Alias destinataire conserve apres remplissage"
    Exiger MemeZone(doc.Bookmarks("APPEL").Range, doc.Bookmarks("FORMULE_APPEL").Range), "Alias appel conserve apres remplissage"
    modCourrier.RemplirSignet doc, "DESTINATAIRE", "DESTINATAIRE FICTIF"
    modCourrier.RemplirSignet doc, "APPEL", "FORMULE FICTIVE"
    Exiger doc.Bookmarks("CORRESPONDANT").Range.Text = "DESTINATAIRE FICTIF", "Alias Dragon destinataire suit le texte"
    Exiger doc.Bookmarks("FORMULE_APPEL").Range.Text = "FORMULE FICTIVE", "Alias Dragon appel suit le texte"
    contenu = doc.Content.Text
    formatCorps = FormatZone(doc.Bookmarks("CORPS").Range)
    modCourrier.NormaliserModele doc
    modCourrier.PreparerStyleCorps doc
    Exiger doc.Content.Text = contenu And FormatZone(doc.Bookmarks("CORPS").Range) = formatCorps, "Normalisation idempotente"
    doc.Close wdDoNotSaveChanges: Set doc = Nothing
    Exiger Documents.Count = nombreDocuments, "Document de test ferme"
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    Application.AutomationSecurity = securite
    If Not original Is Nothing Then original.Close wdDoNotSaveChanges
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "VerifierModele", description
End Sub

Private Sub VerifierPatientInvalide(ByVal racine As String)
    Dim doc As Document, patient As Object, nombre As Long, numero As Long
    Configurer racine, "historique"
    Set patient = CreateObject("Scripting.Dictionary")
    patient("Nom") = "FICTIF SANS IDENTIFIANT"
    nombre = Documents.Count
    On Error Resume Next
    Set doc = modCourrier.CreerCourrierRapidePour(patient)
    numero = Err.Number: Err.Clear
    On Error GoTo 0
    Exiger numero <> 0, "Identifiant patient absent refuse avant creation"
    Exiger doc Is Nothing And Documents.Count = nombre, "Aucun document cree pour un patient sans identifiant"
End Sub

Private Sub VerifierRefus(ByVal racine As String, ByVal nom As String, ByVal erreurAttendue As String)
    Dim doc As Document, nombre As Long, numero As Long, description As String
    Configurer racine, nom
    nombre = Documents.Count
    On Error Resume Next
    Set doc = modCourrier.CreerCourrierRapide()
    numero = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    Exiger numero <> 0 And InStr(1, description, erreurAttendue, vbTextCompare) > 0, "Refus explicite : " & nom
    Exiger doc Is Nothing And Documents.Count = nombre, "Aucun document orphelin : " & nom
End Sub

Public Function ExecuterModeleCourrier(ByVal dossier As String, Optional ByVal modeleReel As String = "") As String
    Dim racine As String, fso As Object, numero As Long, description As String, reelVerifie As Boolean
    On Error GoTo Echec
    mNombre = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(dossier) < 3 Or Mid$(dossier, 2, 2) <> ":\" Then Err.Raise vbObjectError + 1193, , "Dossier de tests local absolu requis."
    If fso.GetDrive(fso.GetDriveName(dossier)).DriveType <> 2 Then Err.Raise vbObjectError + 1193, , "Disque local requis."
    racine = dossier & "\modele-courrier-" & modFichiers.IdUnique()
    fso.CreateFolder racine
    fso.CreateFolder racine & "\Config"
    fso.CreateFolder racine & "\Modeles"
    modConfig.DefinirRacine racine
    CreerFixture racine & "\Modeles\historique.dotx", True, True
    CreerFixture racine & "\Modeles\sans-destinataire.dotx", False, True
    CreerFixture racine & "\Modeles\sans-appel.dotx", True, False
    CreerFixtureModerne racine & "\Modeles\moderne.dotx", True
    CreerFixtureModerne racine & "\Modeles\moderne-incomplet.dotx", False
    ' Une ressource LETTRE TYPE presente ne doit pas masquer un modele nomme absent.
    fso.CopyFile racine & "\Modeles\historique.dotx", racine & "\Modeles\LETTRE TYPE.dotx", False
    VerifierModele racine, "historique", racine & "\Modeles\historique.dotx"
    VerifierPatientInvalide racine
    VerifierRefus racine, "sans-destinataire", "DESTINATAIRE"
    VerifierRefus racine, "sans-appel", "APPEL"
    VerifierRefus racine, "modele-absent", "Modele principal introuvable"
    VerifierModerne racine
    VerifierRefus racine, "moderne-incomplet", "EXPEDITEUR"
    If Len(modeleReel) > 0 Then
        If Len(modeleReel) < 3 Or Mid$(modeleReel, 2, 2) <> ":\" Then Err.Raise vbObjectError + 1193, , "Modele reel local absolu requis."
        If fso.GetDrive(fso.GetDriveName(modeleReel)).DriveType <> 2 Then Err.Raise vbObjectError + 1193, , "Modele reel sur disque local requis."
        fso.CopyFile modeleReel, racine & "\Modeles\modele-reel." & fso.GetExtensionName(modeleReel), False
        VerifierModele racine, "modele-reel", racine & "\Modeles\modele-reel." & fso.GetExtensionName(modeleReel)
        reelVerifie = True
    End If
    ExecuterModeleCourrier = "{""reussis"":" & CStr(mNombre) & ",""echec"":false,""modele_reel_verifie"":" & IIf(reelVerifie, "true", "false") & "}"
Sortie:
    modConfig.DefinirRacine ""
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    ExecuterModeleCourrier = "{""reussis"":" & CStr(mNombre) & ",""echec"":true,""numero"":" & CStr(numero) & ",""description"":" & modServiceNas.JsonValeur(description) & "}"
    Resume Sortie
End Function
