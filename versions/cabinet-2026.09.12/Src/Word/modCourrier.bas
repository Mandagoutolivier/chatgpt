Attribute VB_Name = "modCourrier"
Option Explicit
' =====================================================================
' modCourrier - Creation des courriers depuis les modeles, remplissage
' de l'en-tete SANS RESSAISIE (patient + correspondant + config), et
' acces au corps du courrier via le signet CORPS.
' Contrat principal : DESTINATAIRE et APPEL (ou leurs alias Dragon), CORPS.
' Le modele historique conserve son en-tete et sa signature statiques ;
' Les autres modeles exigent EXPEDITEUR, DATELIEU, CONCERNE et SIGNATURE.
' =====================================================================

' --- Commande principale (Ctrl+Alt+N / commande vocale) --------------
Public Sub NouveauCourrier()
    modPowerMicUnifie.Unifie_A_NouvelleLettre
End Sub
' --- file des patients arrives (deposee par le secretariat) ---------------
' Renvoie le drapeau d'arrivee a prendre (dict, "_Chemin" = fichier), ou
' Nothing. Un seul patient arrive du jour -> pris sans question ; plusieurs
' -> courte liste (heure, nom) ; les arrivees d'un autre jour sont ignorees.

' L'arrivee prise en charge est deplacee dans Arrives\Pris (historique du jour)

' Courrier pour un patient connu mais SANS medecin traitant en base :
' en-tete complet sauf le bloc adresse, a dicter ; patient rattache.
Public Function CreerCourrierRapidePour(ByVal pat As Object) As Document
    On Error GoTo Echec
    Dim doc As Document
    If pat Is Nothing Then Err.Raise vbObjectError + 1172, "modCourrier", "Patient absent du courrier."
    If Not pat.Exists("ID") Then Err.Raise vbObjectError + 1172, "modCourrier", "Identifiant patient absent."
    If Len(Trim$(CStr(pat("ID")))) = 0 Then Err.Raise vbObjectError + 1172, "modCourrier", "Identifiant patient vide."
    Set doc = CreerCourrierRapide()
    doc.Variables("PatientID") = pat("ID")
    If doc.Bookmarks.Exists("CONCERNE") Then
        RemplirSignet doc, "CONCERNE", "Concerne : " & Trim$(modTexte.CiviliteCourte(modTexte.SexePatient(pat)) & " " & _
            pat("Prenom") & " " & pat("Nom")) & ", " & modTexte.NeLe(modTexte.SexePatient(pat)) & " " & modTexte.DdnPatient(pat)
    End If
    Set CreerCourrierRapidePour = doc
    Exit Function
Echec:
    Dim numero As Long, description As String
    numero = Err.Number: description = Err.Description
    FermerCourrierIncomplet doc
    Err.Raise numero, "modCourrier.CreerCourrierRapidePour", description
End Function

' Courrier vide pret a la dictee : en-tete expediteur, date, appel et
' politesse par defaut, signature ; bloc destinataire VIDE (curseur dedans).
Public Function CreerCourrierRapide() As Document
    On Error GoTo Echec
    Dim doc As Document, rng As Range
    Set doc = CreerDepuisModele("LETTRE TYPE")
    RemplirEnTeteSansDestinataire doc
    PreparerStyleCorps doc
    FigerChampsDate doc
    doc.Variables("TypeCourrier") = "consultation"
    ' curseur dans le bloc destinataire : le medecin dicte le medecin traitant
    If doc.Bookmarks.Exists("DESTINATAIRE") Then
        Set rng = doc.Bookmarks("DESTINATAIRE").Range
        rng.Collapse wdCollapseStart
        rng.Select
    Else
        PlacerCurseurCorps doc
    End If
    Set CreerCourrierRapide = doc
    Exit Function
Echec:
    Dim numero As Long, description As String
    numero = Err.Number: description = Err.Description
    FermerCourrierIncomplet doc
    Err.Raise numero, "modCourrier.CreerCourrierRapide", description
End Function

Private Sub RemplirEnTeteSansDestinataire(ByVal doc As Document)
    Dim expediteur As String, signature As String
    expediteur = modConfig.Config("MEDECIN", "Titre", "Docteur") & " " & _
                 modConfig.Config("MEDECIN", "Prenom") & " " & modConfig.Config("MEDECIN", "Nom") & vbCr & _
                 modConfig.Config("MEDECIN", "Specialite") & vbCr & _
                 modConfig.Config("MEDECIN", "AdresseLigne1") & vbCr & _
                 modConfig.Config("MEDECIN", "AdresseLigne2") & vbCr & _
                 "Tel : " & modConfig.Config("MEDECIN", "Telephone")
    signature = Replace(modConfig.Config("MEDECIN", "Signature", ""), "|", vbCr)
    If Len(signature) = 0 Then
        signature = modConfig.Config("MEDECIN", "Titre", "Docteur") & " " & _
                    modConfig.Config("MEDECIN", "Prenom") & " " & modConfig.Config("MEDECIN", "Nom")
    End If
    If doc.Bookmarks.Exists("EXPEDITEUR") Then RemplirSignet doc, "EXPEDITEUR", expediteur
    ' un espace : un signet totalement vide disparait a la premiere frappe
    RemplirSignet doc, "DESTINATAIRE", " "
    MettreEnFormeDestinataire doc
    If doc.Bookmarks.Exists("DATELIEU") Then RemplirSignet doc, "DATELIEU", modConfig.Config("GENERAL", "Ville") & ", le " & Format$(Date, "d mmmm yyyy")
    If doc.Bookmarks.Exists("CONCERNE") Then RemplirSignet doc, "CONCERNE", ""
    RemplirSignet doc, "APPEL", IIf(AppelAuto(), AppelParDefaut(False), " ")
    If PolitesseAuto() Then RemplirSignet doc, "POLITESSE", PolitesseParDefaut(False)
    If doc.Bookmarks.Exists("SIGNATURE") Then RemplirSignet doc, "SIGNATURE", signature
End Sub

' Le destinataire a ete dicte : on tente de le reconnaitre dans la base
' (nom present dans le bloc adresse) pour le rattacher au document
' (formules, lettres de demande, drapeau). Renvoie l'ID trouve ou "".
Public Function ReconnaitreDestinataire(ByVal doc As Document) As String
    Dim id As String, cor As Object
    id = Trim$(modIntegrationUnifie.VariableDoc(doc, "CorrespondantID"))
    If Len(id) > 0 Then
        Set cor = modBase.CorrespondantParID(id)
        ReconnaitreDestinataire = CStr(cor("ID"))
    End If
End Function
' --- Creation sans interface (testable, reutilisee par les derivees) --
Public Function CreerCourrierPour(ByVal pat As Object, ByVal cor As Object, _
                                  Optional ByVal typeCourrier As String = "consultation") As Document
    On Error GoTo Echec
    Dim doc As Document
    Set doc = CreerDepuisModele("LETTRE TYPE")
    RemplirEnTete doc, pat, cor
    PreparerStyleCorps doc
    ' Les modeles modernes conservent leur reglage explicite du bloc adresse.
    MettreEnFormeDestinataire doc
    FigerChampsDate doc
    doc.Variables("PatientID") = pat("ID")
    doc.Variables("CorrespondantID") = cor("ID")
    doc.Variables("TypeCourrier") = typeCourrier
    PlacerCurseurCorps doc
    Set CreerCourrierPour = doc
    Exit Function
Echec:
    Dim numero As Long, description As String
    numero = Err.Number: description = Err.Description
    FermerCourrierIncomplet doc
    Err.Raise numero, "modCourrier.CreerCourrierPour", description
End Function

' nomModele SANS extension : essaie .dotx, .dotm, .dot (le modele reel du
' cabinet peut etre dans n'importe lequel de ces formats)
Private Function CreerDepuisModele(ByVal nomModele As String) As Document
    On Error GoTo Echec
    Dim base As String, ext As Variant, chemin As String, doc As Document
    ' Le choix configure est explicite : aucun repli sur un modele d annexes.
    If nomModele = "LETTRE TYPE" Then nomModele = modConfig.Config("COURRIER", "Modele", "LETTRE TYPE")
    base = modConfig.chemin("Modeles") & "\" & nomModele
    For Each ext In Array(".dotm", ".dotx", ".dot")
        chemin = base & ext
        If modFichiers.FichierExiste(chemin) Then
            Set doc = AjouterModeleSansMacros(chemin)
            NormaliserModele doc
            Set CreerDepuisModele = doc
            Exit Function
        End If
    Next ext
    Err.Raise vbObjectError + 300, "modCourrier", _
        "Modele principal introuvable : " & base & " (.dotm/.dotx/.dot)"
Echec:
    Dim numero As Long, description As String
    numero = Err.Number: description = Err.Description
    FermerCourrierIncomplet doc
    Err.Raise numero, "modCourrier.CreerDepuisModele", description
End Function

' La fonction qui cree le document en reste responsable jusqu a son retour.
Private Sub FermerCourrierIncomplet(ByVal doc As Document)
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
End Sub

Private Sub ExigerSignet(ByVal doc As Document, ByVal nom As String)
    If Not doc.Bookmarks.Exists(nom) Then Err.Raise vbObjectError + 1170, "modCourrier", "Signet obligatoire absent : " & nom
End Sub

Private Function EnteteStatique(ByVal doc As Document) As Boolean
    ' Contrat du modele de dictee historique a deux signets.
    EnteteStatique = doc.Bookmarks.Exists("CORRESPONDANT") And doc.Bookmarks.Exists("FORMULE_APPEL") And _
        Not doc.Bookmarks.Exists("EXPEDITEUR") And Not doc.Bookmarks.Exists("DATELIEU") And _
        Not doc.Bookmarks.Exists("CONCERNE") And Not doc.Bookmarks.Exists("SIGNATURE")
End Function

' Le modele de lettre du medecin (lettretypeclaude.dotm) ne porte que deux
' signets : CORRESPONDANT et FORMULE_APPEL (ceux que Dragon connait). On y
' ajoute les signets attendus par le logiciel sans toucher a la mise en
' page : DESTINATAIRE et APPEL (memes zones), CORPS (paragraphe vide apres
' l'appel), POLITESSE (formule par defaut avant la signature, si
' [COURRIER] PolitesseAuto=1). Idempotent.
Public Sub NormaliserModele(ByVal doc As Document)
    Dim rng As Range, pAppel As Paragraph, pSuiv As Paragraph, nom As Variant
    Alias doc, "DESTINATAIRE", "CORRESPONDANT"
    Alias doc, "APPEL", "FORMULE_APPEL"
    ExigerSignet doc, "DESTINATAIRE"
    ExigerSignet doc, "APPEL"
    If Not EnteteStatique(doc) Then
        For Each nom In Array("EXPEDITEUR", "DATELIEU", "CONCERNE", "SIGNATURE")
            ExigerSignet doc, CStr(nom)
        Next nom
    End If
    If Not doc.Bookmarks.Exists("CORPS") And doc.Bookmarks.Exists("APPEL") Then
        Set pAppel = doc.Bookmarks("APPEL").Range.Paragraphs(1)
        Set pSuiv = pAppel.Next
        ' il faut un paragraphe VIDE entre l'appel et la signature
        If pSuiv Is Nothing Then
            pAppel.Range.InsertParagraphAfter
            Set pSuiv = pAppel.Next
        ElseIf Len(Trim$(Replace(pSuiv.Range.Text, vbCr, ""))) > 0 Then
            pAppel.Range.InsertParagraphAfter
            Set pSuiv = pAppel.Next
        End If
        ' un espace dans le paragraphe : un signet sur du vide ne s'etend
        ' pas quand on dicte dedans
        Set rng = pSuiv.Range
        rng.MoveEnd wdCharacter, -1
        rng.Text = " "
        Set rng = pSuiv.Range
        rng.MoveEnd wdCharacter, -1
        doc.Bookmarks.Add "CORPS", rng
        ' mise en forme du corps : celle des courriers du cabinet
        If Not EnteteStatique(doc) Then
            With pSuiv.Format
                .LeftIndent = 0
                .FirstLineIndent = CentimetersToPoints(modConfig.ConfigNum("COURRIER", "AlineaCm", 0))
                .Alignment = wdAlignParagraphJustify
            End With
            AppliquerEspacement doc, "CORPS"
        End If
    End If
    If Not doc.Bookmarks.Exists("POLITESSE") And doc.Bookmarks.Exists("CORPS") Then
        If PolitesseAuto() Then
            Set pSuiv = doc.Bookmarks("CORPS").Range.Paragraphs(1)
            pSuiv.Range.InsertParagraphAfter
            Set pSuiv = pSuiv.Next
            Set rng = pSuiv.Range
            rng.MoveEnd wdCharacter, -1
            rng.Text = PolitesseParDefaut(False)
            doc.Bookmarks.Add "POLITESSE", rng
            AppliquerEspacement doc, "POLITESSE"
        End If
    End If
    ExigerSignet doc, "CORPS"
End Sub

' Pose le signet 'nouveau' sur la zone du signet 'existant' s'il manque
Private Sub Alias(ByVal doc As Document, ByVal nouveau As String, ByVal existant As String)
    If doc.Bookmarks.Exists(nouveau) Then Exit Sub
    If Not doc.Bookmarks.Exists(existant) Then Exit Sub
    doc.Bookmarks.Add nouveau, doc.Bookmarks(existant).Range
End Sub

Public Sub RemplirEnTete(ByVal doc As Document, ByVal pat As Object, ByVal cor As Object)
    Dim expediteur As String, destinataire As String, concerne As String
    Dim appel As String, signature As String

    expediteur = modConfig.Config("MEDECIN", "Titre", "Docteur") & " " & _
                 modConfig.Config("MEDECIN", "Prenom") & " " & modConfig.Config("MEDECIN", "Nom") & vbCr & _
                 modConfig.Config("MEDECIN", "Specialite") & vbCr & _
                 modConfig.Config("MEDECIN", "AdresseLigne1") & vbCr & _
                 modConfig.Config("MEDECIN", "AdresseLigne2") & vbCr & _
                 "Tel : " & modConfig.Config("MEDECIN", "Telephone")

    ' bloc destinataire pre-compose (classeur des specialistes : structures,
    ' services...) sinon composition classique depuis la fiche
    If cor.Exists("BlocDestinataire") Then destinataire = Trim$(cor("BlocDestinataire"))
    If Len(destinataire) = 0 Then
        destinataire = Trim$(cor("Titre") & " " & cor("Prenom") & " " & cor("Nom"))
        If Len(cor("Specialite")) > 0 Then destinataire = destinataire & vbCr & cor("Specialite")
        If Len(cor("Adresse1")) > 0 Then destinataire = destinataire & vbCr & cor("Adresse1")
        If Len(cor("Adresse2")) > 0 Then destinataire = destinataire & vbCr & cor("Adresse2")
        destinataire = destinataire & vbCr & cor("CP") & " " & cor("Ville")
    End If

    concerne = "Concerne : " & Trim$(modTexte.CiviliteCourte(modTexte.SexePatient(pat)) & " " & _
               pat("Prenom") & " " & pat("Nom")) & ", " & modTexte.NeLe(modTexte.SexePatient(pat)) & " " & modTexte.DdnPatient(pat)

    ' [COURRIER] AppelAuto=0 (defaut) : la formule d'appel n'est PAS
    ' pre-remplie, le medecin la dicte au signet APPEL (bouton B / Ctrl+Alt+Maj+B)
    If AppelAuto() Then
        appel = cor("FormuleAppel")
        If Len(appel) = 0 Then appel = AppelParDefaut(EstTutoye(cor), cor)
    Else
        appel = " "
    End If

    signature = Replace(modConfig.Config("MEDECIN", "Signature", ""), "|", vbCr)
    If Len(signature) = 0 Then
        signature = modConfig.Config("MEDECIN", "Titre", "Docteur") & " " & _
                    modConfig.Config("MEDECIN", "Prenom") & " " & modConfig.Config("MEDECIN", "Nom")
    End If

    Dim politesse As String
    If PolitesseAuto() Then
        politesse = cor("FormulePolitesse")
        If Len(politesse) = 0 Then politesse = PolitesseParDefaut(EstTutoye(cor))
    End If

    If doc.Bookmarks.Exists("EXPEDITEUR") Then RemplirSignet doc, "EXPEDITEUR", expediteur
    ' l'adresse est UN SEUL paragraphe : tous les separateurs de ligne
    ' (vbCrLf, vbCr, vbLf d'une cellule Excel saisie en Alt+Entree) sont
    ' convertis en sauts de ligne manuels, sinon Word cree des paragraphes
    ' espaces de 12 pt et le bloc s'aere.
    RemplirSignet doc, "DESTINATAIRE", EnSautsDeLigne(destinataire)
    MettreEnFormeDestinataire doc
    If doc.Bookmarks.Exists("DATELIEU") Then RemplirSignet doc, "DATELIEU", modConfig.Config("GENERAL", "Ville") & ", le " & Format$(Date, "d mmmm yyyy")
    If doc.Bookmarks.Exists("CONCERNE") Then RemplirSignet doc, "CONCERNE", concerne
    RemplirSignet doc, "APPEL", appel
    If Len(politesse) > 0 Then RemplirSignet doc, "POLITESSE", politesse
    If doc.Bookmarks.Exists("SIGNATURE") Then RemplirSignet doc, "SIGNATURE", signature
End Sub

' --- registre du correspondant ----------------------------------------
' Tutoiement : champ Tutoiement / TutoiementVouvoiement de la fiche
' ("tu"), sinon deduit d'une formule d'appel familiere (Cher Ami, Mon Cher...)
Public Function EstTutoye(ByVal cor As Object) As Boolean
    Dim v As String, appel As String
    If cor Is Nothing Then Exit Function
    If cor.Exists("Tutoiement") Then v = LCase$(Trim$(cor("Tutoiement")))
    If Len(v) = 0 And cor.Exists("TutoiementVouvoiement") Then v = LCase$(Trim$(cor("TutoiementVouvoiement")))
    If v = "tu" Then EstTutoye = True: Exit Function
    If v = "vous" Then Exit Function
    If cor.Exists("FormuleAppel") Then appel = modTexte.Plier(cor("FormuleAppel"))
    EstTutoye = (InStr(appel, "ami") > 0 Or InStr(appel, "mon cher") > 0 Or InStr(appel, "ma chere") > 0)
End Function

' Defauts du cabinet (decisions du 04/09/2026) : "Cher Ami" pour les
' correspondants proches (tutoyes), "Cher Confrère" sinon ;
' "Bien cordialement" au tutoiement, "Bien confraternellement" au vouvoiement.
Public Function AppelParDefaut(ByVal tutoiement As Boolean, Optional ByVal cor As Object = Nothing) As String
    Dim femme As Boolean
    If Not cor Is Nothing Then femme = (modTexte.SexePatient(cor) = "F")
    If tutoiement Then
        If femme Then
            AppelParDefaut = modConfig.Config("COURRIER", "AppelProcheF", "Chère Amie,")
        Else
            AppelParDefaut = modConfig.Config("COURRIER", "AppelProche", "Cher Ami,")
        End If
    ElseIf femme Then
        AppelParDefaut = modConfig.Config("COURRIER", "AppelDefautF", "Chère Consœur,")
    Else
        AppelParDefaut = modConfig.Config("COURRIER", "AppelDefaut", "Cher Confrère,")
    End If
End Function

' [COURRIER] AppelAuto / PolitesseAuto (0 par defaut depuis le 08/09/2026 :
' le medecin dicte lui-meme l'appel et la politesse)
Public Function AppelAuto() As Boolean
    AppelAuto = (modConfig.ConfigNum("COURRIER", "AppelAuto", 0) = 1)
End Function

Public Function PolitesseAuto() As Boolean
    PolitesseAuto = (modConfig.ConfigNum("COURRIER", "PolitesseAuto", 0) = 1)
End Function

Public Function PolitesseParDefaut(ByVal tutoiement As Boolean) As String
    If tutoiement Then
        PolitesseParDefaut = modConfig.Config("COURRIER", "PolitesseProche", "Bien cordialement.")
    Else
        PolitesseParDefaut = modConfig.Config("COURRIER", "PolitesseDefaut", "Bien confraternellement.")
    End If
End Function

' --- identite du patient dans le corps (Ctrl+Alt+P / voix) -----------
' Le medecin ne dicte JAMAIS l'identite : elle est inseree depuis la base.
Public Sub InsererPatient()
    modPowerMicUnifie.Unifie_C_InsererPatient
End Sub
' Ex : "Monsieur Jean FABREGUE, 91 ans"
Public Function TexteIdentitePatient(ByVal pat As Object) As String
    Dim age As String
    age = CalculerAge(modTexte.DdnPatient(pat))
    TexteIdentitePatient = Trim$(modTexte.Civilite(modTexte.SexePatient(pat)) & " " & UCase$(pat("Nom")) & " " & pat("Prenom")) & _
                           IIf(Len(age) > 0, ", " & age & " ans, ", ", ")
End Function

Public Function CalculerAge(ByVal ddn As String) As String
    Dim naissance As Date, age As Long
    If Not modTexte.DateFrValide(ddn) Then Exit Function
    naissance = modTexte.DateFr(ddn)
    If naissance > Date Then Exit Function
    age = DateDiff("yyyy", naissance, Date)
    If Format$(Date, "mmdd") < Format$(naissance, "mmdd") Then age = age - 1
    CalculerAge = CStr(age)
End Function
' Le style de corps est propre au courrier. Ne jamais modifier Normal ni un
' style partage avec l en-tete, le destinataire ou la signature du modele.
Public Sub PreparerStyleCorps(ByVal doc As Document)
    Dim rng As Range, st As Style, modele As ParagraphFormat, police As Font
    ExigerSignet doc, "CORPS"
    Set rng = doc.Bookmarks("CORPS").Range
    Set modele = rng.Paragraphs(1).Format.Duplicate
    Set police = rng.Font.Duplicate
    On Error Resume Next
    Set st = doc.Styles("CabinetCorpsU2")
    On Error GoTo 0
    If st Is Nothing Then Set st = doc.Styles.Add("CabinetCorpsU2", wdStyleTypeParagraph)
    st.AutomaticallyUpdate = False
    st.ParagraphFormat = modele
    st.Font = police
    st.NextParagraphStyle = st
    ' ParagraphFormat conserve aussi le style d origine : le restaurer
    ' avant d appliquer le style reserve, sinon il annule rng.Style.
    rng.ParagraphFormat = modele
    rng.Style = st
    rng.Font = police
    If Not EnteteStatique(doc) Then
        AppliquerEspacement doc, "APPEL"
        AppliquerEspacement doc, "CORPS"
        AppliquerEspacement doc, "POLITESSE"
    End If
End Sub

' Tous les separateurs de ligne -> saut de ligne manuel (Chr 11), et
' aucune ligne vide : le bloc adresse tient en un seul paragraphe.
Public Function EnSautsDeLigne(ByVal texte As String) As String
    Dim t As String
    t = Replace(texte, vbCrLf, Chr$(11))
    t = Replace(t, vbCr, Chr$(11))
    t = Replace(t, vbLf, Chr$(11))
    Do While InStr(t, Chr$(11) & Chr$(11)) > 0
        t = Replace(t, Chr$(11) & Chr$(11), Chr$(11))
    Loop
    Do While Left$(t, 1) = Chr$(11)
        t = Mid$(t, 2)
    Loop
    Do While Right$(t, 1) = Chr$(11)
        t = Left$(t, Len(t) - 1)
    Loop
    EnSautsDeLigne = t
End Function

' Bloc adresse du correspondant : les lignes de l'adresse sont SERREES
' (interligne simple par defaut, [COURRIER] InterligneDestinataire), sans
' espacement avant/apres, tout le bloc en gras. L'interligne du corps
' (1,15) aere trop un bloc de 3 ou 4 lignes courtes.
' Appele APRES chaque etape qui peut remettre en forme les paragraphes.
Public Sub MettreEnFormeDestinataire(ByVal doc As Document)
    On Error Resume Next
    Dim rng As Range, p As Paragraph, interligne As Double
    If Not doc.Bookmarks.Exists("DESTINATAIRE") Then Exit Sub
    If EnteteStatique(doc) Then Exit Sub
    interligne = modConfig.ConfigNum("COURRIER", "InterligneDestinataire", 1)
    Set rng = doc.Bookmarks("DESTINATAIRE").Range
    ' le signet peut ne couvrir qu'une partie du bloc : on l'etend aux
    ' paragraphes entiers, sinon seules les premieres lignes sont reglees
    If rng.Paragraphs.Count > 0 Then
        rng.SetRange rng.Paragraphs(1).Range.Start, _
                     rng.Paragraphs(rng.Paragraphs.Count).Range.End
    End If
    rng.Font.Bold = True
    For Each p In rng.Paragraphs
        With p.Format
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = 0
            .SpaceAfter = 0
            If interligne <= 1 Then
                .LineSpacingRule = wdLineSpaceSingle
            Else
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(interligne)
            End If
        End With
    Next p
    If Err.Number <> 0 Then modLog.LogErreur "MettreEnFormeDestinataire : " & Err.Description
End Sub

' La date d'un courrier medical doit rester celle du jour de la
' consultation. Le modele contient des champs { CREATEDATE }, { DATE }...
' qui se RECALCULENT a chaque ouverture : une lettre relue six mois plus
' tard afficherait la date du jour. On les remplace par leur texte.
Public Sub FigerChampsDate(ByVal doc As Document)
    On Error Resume Next
    Dim story As Range, suite As Range, f As Field, i As Long
    For Each story In doc.StoryRanges
        Set suite = story
        Do While Not suite Is Nothing
            For i = suite.Fields.Count To 1 Step -1
                Set f = suite.Fields(i)
                Select Case f.Type
                    Case wdFieldCreateDate, wdFieldDate, wdFieldTime, _
                         wdFieldPrintDate, wdFieldSaveDate
                        f.Update
                        f.Unlink
                End Select
            Next i
            Set suite = suite.NextStoryRange
        Loop
    Next story
    If Err.Number <> 0 Then modLog.LogErreur "FigerChampsDate : " & Err.Description
End Sub

Private Function VariableDoc(ByVal doc As Document, ByVal nom As String) As String
    On Error Resume Next
    VariableDoc = doc.Variables(nom).Value
    Err.Clear
End Function

' Boutons A / B du PowerMic : se placer dans le bloc destinataire, dans la
' formule d'appel (les signets du modele), sans clavier ni souris.
Public Sub AllerDestinataire()
    AllerAuSignet "DESTINATAIRE"
End Sub

Public Sub AllerAppel()
    AllerAuSignet "APPEL"
End Sub

Public Sub AllerCorps()
    PlacerCurseurCorps ActiveDocument
End Sub

Private Sub AllerAuSignet(ByVal nom As String)
    On Error Resume Next
    Dim rng As Range
    If Not ActiveDocument.Bookmarks.Exists(nom) Then
        MsgBox "Signet " & nom & " absent de ce document (courrier cree par 'Nouveau courrier' ?).", vbExclamation, "Cabinet"
        Exit Sub
    End If
    Set rng = ActiveDocument.Bookmarks(nom).Range
    ' selectionner le contenu (un espace ou le texte existant) : dicter le remplace
    If Len(Trim$(rng.Text)) = 0 Then
        RemplirSignet ActiveDocument, nom, "  "
        Set rng = ActiveDocument.Bookmarks(nom).Range
        rng.Collapse wdCollapseStart
        rng.Move wdCharacter, 1
    End If
    rng.Select
End Sub

' Depannage : resserrer le bloc adresse d'un courrier DEJA ouvert.
' Sans signet DESTINATAIRE, agit sur les paragraphes selectionnes.
Public Sub ResserrerDestinataire()
    On Error GoTo Erreur
    Dim doc As Document, p As Paragraph
    Set doc = ActiveDocument
    If doc.Bookmarks.Exists("DESTINATAIRE") Then
        MettreEnFormeDestinataire doc
    ElseIf selection.Type <> wdSelectionIP Then
        For Each p In selection.Range.Paragraphs
            With p.Format
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacingRule = wdLineSpaceSingle
            End With
        Next p
        selection.Range.Font.Bold = True
    Else
        MsgBox "Selectionnez les lignes de l'adresse, puis relancez.", vbInformation, "Cabinet"
    End If
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

' Espacement d'usage des courriers (config [COURRIER] : EspaceAvantPt,
' EspaceApresPt, Interligne) applique aux paragraphes d'un signet.
Private Sub AppliquerEspacement(ByVal doc As Document, ByVal signet As String)
    On Error Resume Next
    Dim p As Paragraph
    If Not doc.Bookmarks.Exists(signet) Then Exit Sub
    For Each p In doc.Bookmarks(signet).Range.Paragraphs
        With p.Format
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = modConfig.ConfigNum("COURRIER", "EspaceAvantPt", 12)
            .SpaceAfter = modConfig.ConfigNum("COURRIER", "EspaceApresPt", 0)
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = LinesToPoints(modConfig.ConfigNum("COURRIER", "Interligne", 1.15))
        End With
    Next p
End Sub

' Remplace le contenu d'un signet et RECREE le signet sur le texte insere
Public Sub RemplirSignet(ByVal doc As Document, ByVal nom As String, ByVal texte As String)
    Dim rng As Range, aliasDragon As String
    If Not doc.Bookmarks.Exists(nom) Then Err.Raise vbObjectError + 1170, , "Signet obligatoire absent : " & nom
    ' signets du modele du medecin (connus des commandes Dragon) : reposes
    If nom = "DESTINATAIRE" Then aliasDragon = "CORRESPONDANT"
    If nom = "APPEL" Then aliasDragon = "FORMULE_APPEL"
    If Len(aliasDragon) > 0 Then
        If Not doc.Bookmarks.Exists(aliasDragon) Then aliasDragon = ""
    End If
    Set rng = doc.Bookmarks(nom).Range
    If rng.Text = texte Then Exit Sub
    rng.Text = texte
    doc.Bookmarks.Add nom, rng
    If Len(aliasDragon) > 0 Then doc.Bookmarks.Add aliasDragon, rng
End Sub

' --- Corps du courrier ------------------------------------------------

Public Function CorpsRange(ByVal doc As Document) As Range
    If doc.Bookmarks.Exists("CORPS") Then
        Set CorpsRange = doc.Bookmarks("CORPS").Range
    ElseIf selection.Type = wdSelectionNormal And Len(selection.Range.Text) > 10 Then
        ' repli : la selection manuelle du medecin fait foi
        Set CorpsRange = selection.Range
    Else
        Err.Raise vbObjectError + 301, "modCourrier", _
            "Signet CORPS introuvable dans ce document et aucune selection : " & _
            "utilisez un courrier cree par 'Nouveau courrier', ou selectionnez le texte a corriger."
    End If
End Function

Public Function RecupererCorps(ByVal doc As Document) As String
    Dim t As String
    t = CorpsRange(doc).Text
    ' retire les marques de paragraphe de tete/queue sans toucher au reste
    Do While Len(t) > 0 And (Right$(t, 1) = vbCr Or Right$(t, 1) = Chr$(7))
        t = Left$(t, Len(t) - 1)
    Loop
    Do While Len(t) > 0 And Left$(t, 1) = vbCr
        t = Mid$(t, 2)
    Loop
    RecupererCorps = t
End Function

Public Sub RemplacerCorps(ByVal doc As Document, ByVal texte As String)
    Dim rng As Range, avaitSignet As Boolean, modele As ParagraphFormat
    avaitSignet = doc.Bookmarks.Exists("CORPS")
    Set rng = CorpsRange(doc)
    ' mise en forme du corps AVANT insertion : Word donnerait sinon au texte
    ' insere celle du paragraphe suivant du modele (espacement automatique...)
    Set modele = rng.Paragraphs(1).Format.Duplicate
    rng.Text = texte & vbCr
    UniformiserParagraphes rng, modele
    If avaitSignet Then doc.Bookmarks.Add "CORPS", rng
    AppliquerEspacement doc, "CORPS"
End Sub

' Impose a tous les paragraphes de la plage la mise en forme de reference
Private Sub UniformiserParagraphes(ByVal rng As Range, ByVal modele As ParagraphFormat)
    On Error Resume Next
    Dim p As Paragraph
    For Each p In rng.Paragraphs
        p.Format = modele
    Next p
End Sub

Public Sub PlacerCurseurCorps(ByVal doc As Document)
    On Error Resume Next
    If doc.Bookmarks.Exists("CORPS") Then
        Dim rng As Range
        Set rng = doc.Bookmarks("CORPS").Range
        rng.Collapse wdCollapseStart
        rng.Move wdCharacter, 1     ' a l'interieur du signet (il s'etend en dictant)
        rng.Select
    End If
End Sub

Private Function AjouterModeleSansMacros(ByVal chemin As String) As Document
    Dim securite As Long, n As Long, description As String
    securite = Application.AutomationSecurity
    On Error GoTo Echec
    Application.AutomationSecurity = 3
    Set AjouterModeleSansMacros = Documents.Add(Template:=chemin)
    Application.AutomationSecurity = securite
    Exit Function
Echec:
    n = Err.Number: description = Err.Description
    Application.AutomationSecurity = securite
    Err.Raise n, "AjouterModeleSansMacros", description
End Function
