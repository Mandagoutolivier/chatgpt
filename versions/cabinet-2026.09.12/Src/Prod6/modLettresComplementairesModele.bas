Attribute VB_Name = "modLettresComplementairesModele"
Option Explicit

'===============================================================================
' MODULE : modLettresComplementairesModele
' VERSION : MODELE-2A — ouverture directe du modèle, sans Documents.Add
'
' Principe :
' - chaque lettre complémentaire est ouverte directement depuis lettremodeledragon.dotm en lecture seule ;
' - le VBA ne reconstruit plus la mise en page ;
' - il remplace uniquement quatre repères existant dans le modèle :
'       MTetMS / MTouMS
'       APPEL_LETTRE
'       CORPS_LETTRE
'       POLITESSE_LETTRE
' - les marques de paragraphe des repères restent en place ;
' - APPEL_LETTRE et POLITESSE_LETTRE gardent donc exactement leur paragraphe ;
' - CORPS_LETTRE peut devenir plusieurs paragraphes : chaque paragraphe créé
'   reçoit le format de paragraphe du repère CORPS_LETTRE ;
' - l'en-tête, la date et la signature ne sont jamais reconstruits ni reformatés ;
' - le champ CREATEDATE est mis à jour puis figé en texte.
'===============================================================================

Private Const LCM_APP_REGISTRE As String = "ModeleCourrierChatGPT"
Private Const LCM_SECTION_REGISTRE As String = "LettresComplementaires"
Private Const LCM_CLE_REGISTRE As String = "CheminModeleWord"

Private Const LCM_REPERE_DEST_1 As String = "MTetMS"
Private Const LCM_REPERE_DEST_2 As String = "MTouMS"
Private Const LCM_REPERE_DEST_3 As String = "MTETMS"
Private Const LCM_REPERE_DEST_4 As String = "MTOUMS"

Private Const LCM_REPERE_APPEL As String = "APPEL_LETTRE"
Private Const LCM_REPERE_CORPS As String = "CORPS_LETTRE"
Private Const LCM_REPERE_POLITESSE As String = "POLITESSE_LETTRE"

'-------------------------------------------------------------------------------
' Création publique appelée par modProdRapide
'-------------------------------------------------------------------------------

Public Function LCM_CreerLettreDepuisModele( _
    ByVal docPrincipal As Document, _
    ByVal blocDestinataire As String, _
    ByVal formuleAppel As String, _
    ByVal corpsDestination As String, _
    ByVal formulePolitesse As String, _
    ByRef messageErreur As String) As Document

    Dim cheminModele As String
    Dim docDemande As Document
    Dim rngCorps As Range

    Dim texteDestinataire As String
    Dim texteCorps As String
    Dim etape As String

    On Error GoTo GestionErreur

    Set LCM_CreerLettreDepuisModele = Nothing
    messageErreur = ""

    If docPrincipal Is Nothing Then
        messageErreur = "Le courrier principal n'est plus disponible."
        Exit Function
    End If

    cheminModele = _
        LCM_ResoudreCheminModele(docPrincipal)

    If cheminModele = "" Then

        messageErreur = _
            "Le modèle Word de dictée n'a pas été retrouvé." & vbCrLf & _
            "Exécutez LCM_ChoisirModeleLettreComplementaire puis " & _
            "sélectionnez le fichier lettremodeledragon.dotm actuellement utilisé pour la dictée."

        Exit Function
    End If

    'MODELE-2A :
    'Documents.Add avec un .dotm provoque de façon intermittente les erreurs
    '5981 / 5844 sur cette installation Word 2016.
    '
    'On n'instancie donc plus le modèle avec Documents.Add.
    'On ouvre directement le fichier modèle en LECTURE SEULE, on le modifie
    'uniquement en mémoire, on le copie dans le courrier principal, puis
    'modProdRapide le ferme sans enregistrer.
    '
    'Le fichier lettremodeledragon.dotm sur disque reste donc strictement intact.

    etape = "Réactivation du courrier principal"

    Application.ScreenUpdating = True

    If docPrincipal.Windows.Count > 0 Then
        docPrincipal.Windows(1).Visible = True
        docPrincipal.Windows(1).Activate
    Else
        docPrincipal.Activate
    End If

    DoEvents

    etape = "Ouverture directe de lettremodeledragon.dotm en lecture seule"

    Set docDemande = LCM_OuvrirSansMacros(cheminModele)

    If docDemande Is Nothing Then

        messageErreur = _
            "Word n'a pas pu ouvrir le modèle de lettre :" & _
            vbCrLf & cheminModele

        Exit Function
    End If

    docDemande.Activate

    '===========================================================================
    ' 1. DESTINATAIRE
    '===========================================================================
    etape = "Destinataire : normalisation"

    texteDestinataire = _
        LCM_NormaliserBlocDestinataire(blocDestinataire)

    If texteDestinataire = "" Then

        messageErreur = _
            "Le bloc destinataire est vide."

        GoTo Echec
    End If

    etape = "Destinataire : remplacement du repère"

    If Not LCM_RemplacerDestinataireEnConservantModele( _
        docDemande, _
        texteDestinataire, _
        etape) Then

        messageErreur = _
            "Le repère MTetMS / MTouMS est introuvable dans le modèle Word."

        GoTo Echec
    End If

    '===========================================================================
    ' 2. DATE
    '===========================================================================
    etape = "Date : mise à jour et figement"

    LCM_MettreAJourEtFigerDate docDemande

    '===========================================================================
    ' 3. FORMULE D'APPEL
    '===========================================================================
    etape = "Formule d'appel : remplacement de APPEL_LETTRE"

    If Not LCM_RemplacerRepereSimpleEnConservantModele( _
        docDemande, _
        LCM_REPERE_APPEL, _
        Trim$(formuleAppel), _
        etape) Then

        messageErreur = _
            "Le repère APPEL_LETTRE est introuvable dans le modèle Word."

        GoTo Echec
    End If

    '===========================================================================
    ' 4. CORPS
    '===========================================================================
    etape = "Corps : restauration de l'identité patient"

    texteCorps = _
        RestaurerPatientDansTexte(corpsDestination)

    etape = "Corps : normalisation"

    texteCorps = _
        LCM_NormaliserParagraphes(texteCorps)

    'Sécurité MODELE2C : même si l'API fusionne plusieurs catégories médicales,
    'on force les principaux examens décrits à commencer sur un paragraphe distinct.
    'Les libellés sont également remis en **gras Markdown** avant conversion Word.
    texteCorps = _
        LCM_ForcerParagraphesExamens(texteCorps)

    If Trim$(texteCorps) = "" Then

        messageErreur = _
            "Le corps de la lettre complémentaire est vide."

        GoTo Echec
    End If

    etape = "Corps : remplacement de CORPS_LETTRE"

    If Not LCM_RemplacerCorpsEnConservantModele( _
        docDemande, _
        texteCorps, _
        rngCorps, _
        etape) Then

        messageErreur = _
            "Le repère CORPS_LETTRE est introuvable dans le modèle Word."

        GoTo Echec
    End If

    'Le texte médical est d'abord remis au format du paragraphe CORPS_LETTRE.
    'Ensuite seulement, les **...** redeviennent du gras.
    etape = "Corps : conversion du gras Markdown"

    ConvertirMarkdownGrasDansRange rngCorps

    etape = "Corps : identité patient en gras"

    MettreIdentitePatientEnGrasDansRange rngCorps

    '===========================================================================
    ' 5. FORMULE DE POLITESSE
    '===========================================================================
    etape = "Formule de politesse : remplacement de POLITESSE_LETTRE"

    If Not LCM_RemplacerRepereSimpleEnConservantModele( _
        docDemande, _
        LCM_REPERE_POLITESSE, _
        Trim$(formulePolitesse), _
        etape) Then

        messageErreur = _
            "Le repère POLITESSE_LETTRE est introuvable dans le modèle Word."

        GoTo Echec
    End If

    etape = "Lettre complémentaire terminée"

    Set LCM_CreerLettreDepuisModele = docDemande
    Exit Function

Echec:

    On Error Resume Next

    If Not docDemande Is Nothing Then
        docDemande.Close SaveChanges:=wdDoNotSaveChanges
    End If

    Set docDemande = Nothing

    On Error GoTo 0

    Exit Function

GestionErreur:

    messageErreur = _
        "Erreur pendant la création depuis le modèle Word : " & _
        Err.Number & " - " & Err.description & vbCrLf & vbCrLf & _
        "Étape : " & etape & vbCrLf & _
        "Documents ouverts : " & Application.Documents.Count & vbCrLf & _
        "Fenêtres Word : " & Application.Windows.Count & vbCrLf & _
        "Modèle tenté : " & cheminModele

    Resume Echec

End Function

'-------------------------------------------------------------------------------
' Sélection / résolution du modèle
'-------------------------------------------------------------------------------

Public Sub LCM_ChoisirModeleLettreComplementaire()
    MsgBox "Le modele est defini dans Config\config.ini sur le NAS, rubrique PROD6, cle ModeleComplementaire.", vbInformation, "Modele partage"
End Sub
Public Sub LCM_TesterModeleLettreComplementaire()

    Dim chemin As String

    If Documents.Count = 0 Then
        MsgBox "Aucun courrier Word n'est ouvert.", vbExclamation
        Exit Sub
    End If

    chemin = _
        LCM_ResoudreCheminModele(ActiveDocument)

    If chemin = "" Then

        MsgBox _
            "Aucun modèle compatible n'a été trouvé.", _
            vbExclamation, _
            "Modèle de lettre complémentaire"

    Else

        MsgBox _
            "Modèle utilisé pour les lettres complémentaires :" & _
            vbCrLf & vbCrLf & chemin, _
            vbInformation, _
            "Modèle de lettre complémentaire"

    End If

End Sub

Public Sub LCM_TesterOuvertureModele()

    Dim chemin As String
    Dim docTest As Document

    On Error GoTo GestionErreur

    chemin = LCM_ResoudreCheminModele(ActiveDocument)

    If Trim$(chemin) = "" Then
        MsgBox "Aucun modèle n'a été trouvé.", vbExclamation
        Exit Sub
    End If

    Set docTest = LCM_OuvrirSansMacros(chemin)

    MsgBox _
        "Ouverture directe du modèle : OK" & vbCrLf & vbCrLf & _
        chemin, _
        vbInformation, _
        "MODELE-2A"

    docTest.Close SaveChanges:=wdDoNotSaveChanges
    Exit Sub

GestionErreur:

    MsgBox _
        "Échec de l'ouverture directe du modèle :" & vbCrLf & _
        Err.Number & " - " & Err.description & vbCrLf & vbCrLf & _
        chemin, _
        vbExclamation, _
        "MODELE-2A"

    On Error Resume Next

    If Not docTest Is Nothing Then
        docTest.Close SaveChanges:=wdDoNotSaveChanges
    End If

    On Error GoTo 0

End Sub

Private Function LCM_ResoudreCheminModele(ByVal docPrincipal As Document) As String
    Dim chemin As String
    chemin = modConfig.CheminNasConfigure("PROD6", "ModeleComplementaire", "Modeles\LETTRE TYPE.dot")
    If Not modFichiers.FichierExiste(chemin) Then Err.Raise vbObjectError + 962, "modLettresComplementairesModele", "Modele de lettre complementaire absent sur le NAS : " & chemin
    LCM_ResoudreCheminModele = chemin
End Function





'-------------------------------------------------------------------------------
' Remplacement des repères du modèle
'-------------------------------------------------------------------------------

Private Function LCM_RemplacerDestinataireEnConservantModele( _
    ByVal doc As Document, _
    ByVal texteDestinataire As String, _
    ByRef etape As String) As Boolean

    Dim reperes As Variant
    Dim i As Long
    Dim rng As Range

    LCM_RemplacerDestinataireEnConservantModele = False

    If doc Is Nothing Then Exit Function

    reperes = Array( _
        LCM_REPERE_DEST_1, _
        LCM_REPERE_DEST_2, _
        LCM_REPERE_DEST_3, _
        LCM_REPERE_DEST_4)

    For i = LBound(reperes) To UBound(reperes)

        etape = _
            "Destinataire : recherche de " & _
            CStr(reperes(i))

        Set rng = _
            LCM_TrouverRepere(doc, CStr(reperes(i)))

        If Not rng Is Nothing Then

            etape = _
                "Destinataire : remplacement en conservant le paragraphe du modèle"

            LCM_RemplacerTexteRangeEnConservantPolice _
                doc, _
                rng, _
                texteDestinataire

            LCM_RemplacerDestinataireEnConservantModele = True
            Exit Function

        End If

    Next i

End Function

Private Function LCM_RemplacerRepereSimpleEnConservantModele( _
    ByVal doc As Document, _
    ByVal repere As String, _
    ByVal nouveauTexte As String, _
    ByRef etape As String) As Boolean

    Dim rng As Range

    LCM_RemplacerRepereSimpleEnConservantModele = False

    If doc Is Nothing Then Exit Function

    etape = "Recherche du repère " & repere

    Set rng = _
        LCM_TrouverRepere(doc, repere)

    If rng Is Nothing Then Exit Function

    etape = _
        "Remplacement de " & repere & _
        " sans toucher à sa marque de paragraphe"

    LCM_RemplacerTexteRangeEnConservantPolice _
        doc, _
        rng, _
        nouveauTexte

    LCM_RemplacerRepereSimpleEnConservantModele = True

End Function

Private Function LCM_RemplacerCorpsEnConservantModele( _
    ByVal doc As Document, _
    ByVal nouveauTexte As String, _
    ByRef rngCorpsCree As Range, _
    ByRef etape As String) As Boolean

    Dim rngRepere As Range
    Dim rngNouveau As Range
    Dim p As Paragraph

    Dim debut As Long
    Dim fin As Long

    'Police effective du repère CORPS_LETTRE.
    Dim nomPolice As String
    Dim taillePolice As Single
    Dim gras As Long
    Dim italique As Long
    Dim soulignement As Long
    Dim petitesMajuscules As Long
    Dim toutesMajuscules As Long

    'Format effectif du paragraphe CORPS_LETTRE.
    Dim alignement As Long
    Dim retraitGauche As Single
    Dim retraitDroit As Single
    Dim retraitPremiereLigne As Single
    Dim espaceAvant As Single
    Dim espaceApres As Single
    Dim regleInterligne As Long
    Dim interligne As Single
    Dim conserverEnsemble As Long
    Dim conserverAvecSuivant As Long
    Dim controleVeuves As Long
    Dim sautAvant As Long

    LCM_RemplacerCorpsEnConservantModele = False
    Set rngCorpsCree = Nothing

    If doc Is Nothing Then Exit Function

    etape = "Corps : recherche du repère CORPS_LETTRE"

    Set rngRepere = _
        LCM_TrouverRepere(doc, LCM_REPERE_CORPS)

    If rngRepere Is Nothing Then Exit Function

    'Mémoriser la police réellement affichée par Word.
    etape = "Corps : mémorisation de la police du modèle"

    With rngRepere.Font
        nomPolice = .Name
        taillePolice = .Size
        gras = .Bold
        italique = .Italic
        soulignement = .Underline
        petitesMajuscules = .SmallCaps
        toutesMajuscules = .AllCaps
    End With

    'Mémoriser l'intégralité des propriétés utiles du paragraphe modèle.
    etape = "Corps : mémorisation de l'alinéa et de l'alignement du modèle"

    With rngRepere.Paragraphs(1).Format
        alignement = .Alignment
        retraitGauche = .LeftIndent
        retraitDroit = .RightIndent
        retraitPremiereLigne = .FirstLineIndent
        espaceAvant = .SpaceBefore
        espaceApres = .SpaceAfter
        regleInterligne = .LineSpacingRule
        interligne = .LineSpacing
        conserverEnsemble = .KeepTogether
        conserverAvecSuivant = .KeepWithNext
        controleVeuves = .WidowControl
        sautAvant = .PageBreakBefore
    End With

    debut = rngRepere.Start

    etape = "Corps : remplacement du texte repère"

    'On ne sélectionne PAS la marque de paragraphe finale du repère.
    'Elle reste donc celle du modèle.
    rngRepere.Text = nouveauTexte

    fin = debut + Len(nouveauTexte)

    Set rngNouveau = doc.Range( _
        Start:=debut, _
        End:=fin)

    etape = "Corps : application de la police du repère à tout le texte créé"

    LCM_AppliquerPoliceMemorisee _
        rngNouveau, _
        nomPolice, _
        taillePolice, _
        gras, _
        italique, _
        soulignement, _
        petitesMajuscules, _
        toutesMajuscules

    etape = "Corps : application du format CORPS_LETTRE à chaque paragraphe"

    For Each p In rngNouveau.Paragraphs

        With p.Format

            .Alignment = alignement
            .LeftIndent = retraitGauche
            .RightIndent = retraitDroit
            .FirstLineIndent = retraitPremiereLigne

            .SpaceBefore = espaceAvant
            .SpaceAfter = espaceApres

            .LineSpacingRule = regleInterligne

            If interligne > 0 Then
                .LineSpacing = interligne
            End If

            .KeepTogether = conserverEnsemble
            .KeepWithNext = conserverAvecSuivant
            .WidowControl = controleVeuves
            .PageBreakBefore = sautAvant

        End With

    Next p

    Set rngCorpsCree = rngNouveau.Duplicate

    LCM_RemplacerCorpsEnConservantModele = True

End Function

Private Sub LCM_RemplacerTexteRangeEnConservantPolice( _
    ByVal doc As Document, _
    ByVal rngRepere As Range, _
    ByVal nouveauTexte As String)

    Dim debut As Long
    Dim fin As Long
    Dim rngNouveau As Range

    Dim nomPolice As String
    Dim taillePolice As Single
    Dim gras As Long
    Dim italique As Long
    Dim soulignement As Long
    Dim petitesMajuscules As Long
    Dim toutesMajuscules As Long

    If doc Is Nothing Then Exit Sub
    If rngRepere Is Nothing Then Exit Sub

    With rngRepere.Font
        nomPolice = .Name
        taillePolice = .Size
        gras = .Bold
        italique = .Italic
        soulignement = .Underline
        petitesMajuscules = .SmallCaps
        toutesMajuscules = .AllCaps
    End With

    debut = rngRepere.Start

    'Le Range trouvé contient uniquement le texte du repère, jamais ¶.
    rngRepere.Text = nouveauTexte

    If Len(nouveauTexte) = 0 Then Exit Sub

    fin = debut + Len(nouveauTexte)

    Set rngNouveau = doc.Range( _
        Start:=debut, _
        End:=fin)

    LCM_AppliquerPoliceMemorisee _
        rngNouveau, _
        nomPolice, _
        taillePolice, _
        gras, _
        italique, _
        soulignement, _
        petitesMajuscules, _
        toutesMajuscules

End Sub

Private Sub LCM_AppliquerPoliceMemorisee( _
    ByVal rng As Range, _
    ByVal nomPolice As String, _
    ByVal taillePolice As Single, _
    ByVal gras As Long, _
    ByVal italique As Long, _
    ByVal soulignement As Long, _
    ByVal petitesMajuscules As Long, _
    ByVal toutesMajuscules As Long)

    If rng Is Nothing Then Exit Sub

    With rng.Font

        If Trim$(nomPolice) <> "" Then
            .Name = nomPolice
        End If

        If taillePolice <> wdUndefined Then
            .Size = taillePolice
        End If

        If gras <> wdUndefined Then
            .Bold = gras
        End If

        If italique <> wdUndefined Then
            .Italic = italique
        End If

        If soulignement <> wdUndefined Then
            .Underline = soulignement
        End If

        If petitesMajuscules <> wdUndefined Then
            .SmallCaps = petitesMajuscules
        End If

        If toutesMajuscules <> wdUndefined Then
            .AllCaps = toutesMajuscules
        End If

    End With

End Sub

Private Function LCM_TrouverRepere( _
    ByVal doc As Document, _
    ByVal repere As String) As Range

    Dim rng As Range

    Set LCM_TrouverRepere = Nothing

    If doc Is Nothing Then Exit Function
    If Trim$(repere) = "" Then Exit Function

    If doc.Bookmarks.Exists(repere) Then
        Set rng = doc.Bookmarks(repere).Range.Duplicate
        If Len(rng.Text) > 0 Then
            If Right$(rng.Text, 1) = vbCr Then rng.MoveEnd wdCharacter, -1
        End If
        Set LCM_TrouverRepere = rng
        Exit Function
    End If
    Set rng = doc.Content.Duplicate



    If modRechercheWord.Trouver(rng, repere) Then
        Set LCM_TrouverRepere = rng.Duplicate
    End If

End Function

Private Function LCM_NormaliserBlocDestinataire( _
    ByVal texte As String) As String

    Dim lignes() As String
    Dim resultat As String
    Dim ligne As String
    Dim i As Long

    texte = Replace(texte, vbCrLf, vbLf)
    texte = Replace(texte, vbCr, vbLf)
    texte = Replace(texte, Chr$(11), vbLf)

    lignes = Split(texte, vbLf)

    resultat = ""

    For i = LBound(lignes) To UBound(lignes)

        ligne = _
            Trim$(Replace(CStr(lignes(i)), Chr$(160), " "))

        If ligne <> "" Then

            If resultat <> "" Then

                'Retour manuel : le destinataire reste dans UN SEUL paragraphe.
                'Le retrait défini par MTetMS dans le modèle est donc conservé.
                resultat = resultat & Chr$(11)

            End If

            resultat = resultat & ligne

        End If

    Next i

    LCM_NormaliserBlocDestinataire = resultat

End Function

'-------------------------------------------------------------------------------
' Date
'-------------------------------------------------------------------------------

Private Sub LCM_MettreAJourEtFigerDate( _
    ByVal doc As Document)

    Dim i As Long
    Dim rngDate As Range
    Dim texteDate As String

    If doc Is Nothing Then Exit Sub

    'MODELE-2A :
    'le fichier modèle est maintenant ouvert directement.
    'Un champ CREATEDATE afficherait donc la date de création du MODÈLE,
    'et non celle de la lettre complémentaire.
    '
    'On remplace explicitement le résultat de DATE/CREATEDATE par la date
    'du jour, puis on supprime le champ. Le document final ne pourra donc
    'jamais réactualiser cette date à sa prochaine ouverture.

    texteDate = Format$(Date, "dddd d mmmm yyyy")

    On Error Resume Next

    For i = doc.Fields.Count To 1 Step -1

        If doc.Fields(i).Type = wdFieldCreateDate _
        Or doc.Fields(i).Type = wdFieldDate Then

            Set rngDate = doc.Fields(i).Result.Duplicate

            doc.Fields(i).Unlink

            If Not rngDate Is Nothing Then
                rngDate.Text = texteDate
            End If

            Exit For

        End If

    Next i

    On Error GoTo 0

End Sub


'-------------------------------------------------------------------------------
' Sécurisation de la structure des examens dans les lettres complémentaires - narratifs
'-------------------------------------------------------------------------------

Public Sub LCM_TesterSeparationECGNarratif()

    Dim texteTest As String
    Dim resultat As String

    texteTest = _
        "Il décrit actuellement une **dyspnée** prédominant aux changements de position, " & _
        "sans limitation nette lors des efforts sportifs. " & _
        "L’**électrocardiogramme** alterne entre un **rythme natif** et un " & _
        "**rythme électro-entraîné** en mode **VVI**."

    resultat = LCM_ForcerParagraphesExamens(texteTest)

    If InStr(1, resultat, _
        vbCr & "**Électrocardiogramme** :", _
        vbTextCompare) > 0 Then

        MsgBox _
            "OK : l'ECG narratif est séparé dans un nouveau paragraphe." & _
            vbCrLf & vbCrLf & resultat, _
            vbInformation, _
            "Test séparation ECG"

    Else

        MsgBox _
            "ÉCHEC : l'ECG est resté dans le même paragraphe." & _
            vbCrLf & vbCrLf & resultat, _
            vbExclamation, _
            "Test séparation ECG"

    End If

End Sub

Private Function LCM_ForcerParagraphesExamens( _
    ByVal texte As String) As String

    Dim libelles As Variant
    Dim i As Long

    'MODELE2C : traiter d'abord les formulations narratives naturelles.
    'Exemple : "... dyspnée. L'électrocardiogramme alterne..." devient :
    '          "... dyspnée." + nouveau paragraphe +
    '          "**Électrocardiogramme** : alterne..."
    texte = LCM_ForcerPhrasesNarrativesExamens(texte)

    'Puis sécuriser les libellés déjà structurés par l'API.
    libelles = Array( _
        "Examen clinique", _
        "Électrocardiogramme", _
        "Electrocardiogramme", _
        "Échographie cardiaque", _
        "Echographie cardiaque", _
        "Échographie", _
        "Echographie", _
        "Échocardiographie cardiaque", _
        "Echocardiographie cardiaque", _
        "Échocardiographie", _
        "Echocardiographie", _
        "Biologie", _
        "Bilan biologique", _
        "Holter rythmique", _
        "Holter ECG", _
        "Holter tensionnel", _
        "MAPA")

    For i = LBound(libelles) To UBound(libelles)
        texte = LCM_NormaliserLibelleExamen( _
            texte, _
            CStr(libelles(i)))
    Next i

    LCM_ForcerParagraphesExamens = texte

End Function

Private Function LCM_ForcerPhrasesNarrativesExamens( _
    ByVal texte As String) As String

    'ECG : formes longues avant formes courtes pour éviter les recouvrements.
    'Variantes Markdown : l'article reste hors du gras mais le nom de l'examen est déjà **...**.
    texte = LCM_NormaliserPhraseExamen(texte, "Sur l’**électrocardiogramme**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "Sur l'**electrocardiogramme**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "Sur l’**ECG**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "Sur l'**ECG**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "À l’**électrocardiogramme**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "A l'**electrocardiogramme**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "À l’**ECG**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "A l'**ECG**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "L’**électrocardiogramme**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "L'**electrocardiogramme**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "L’**ECG**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "L'**ECG**", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "Sur l’électrocardiogramme", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "Sur l'electrocardiogramme", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "Sur l’ECG", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "Sur l'ECG", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "À l’électrocardiogramme", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "A l'electrocardiogramme", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "À l’ECG", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "A l'ECG", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "L’électrocardiogramme", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "L'electrocardiogramme", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "L’ECG", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "L'ECG", "Électrocardiogramme")
    texte = LCM_NormaliserPhraseExamen(texte, "ECG", "Électrocardiogramme")

    'Examen clinique.
    texte = LCM_NormaliserPhraseExamen(texte, "À l’**examen clinique**", "Examen clinique")
    texte = LCM_NormaliserPhraseExamen(texte, "A l'**examen clinique**", "Examen clinique")
    texte = LCM_NormaliserPhraseExamen(texte, "L’**examen clinique**", "Examen clinique")
    texte = LCM_NormaliserPhraseExamen(texte, "L'**examen clinique**", "Examen clinique")
    texte = LCM_NormaliserPhraseExamen(texte, "À l’examen clinique", "Examen clinique")
    texte = LCM_NormaliserPhraseExamen(texte, "A l'examen clinique", "Examen clinique")
    texte = LCM_NormaliserPhraseExamen(texte, "L’examen clinique", "Examen clinique")
    texte = LCM_NormaliserPhraseExamen(texte, "L'examen clinique", "Examen clinique")
    texte = LCM_NormaliserPhraseExamen(texte, "Cliniquement", "Examen clinique")

    'Échographie / échocardiographie.
    texte = LCM_NormaliserPhraseExamen(texte, "À l’**échocardiographie**", "Échocardiographie")
    texte = LCM_NormaliserPhraseExamen(texte, "A l'**echocardiographie**", "Échocardiographie")
    texte = LCM_NormaliserPhraseExamen(texte, "L’**échocardiographie**", "Échocardiographie")
    texte = LCM_NormaliserPhraseExamen(texte, "L'**echocardiographie**", "Échocardiographie")
    texte = LCM_NormaliserPhraseExamen(texte, "L’**échographie cardiaque**", "Échographie cardiaque")
    texte = LCM_NormaliserPhraseExamen(texte, "L'**echographie cardiaque**", "Échographie cardiaque")
    texte = LCM_NormaliserPhraseExamen(texte, "L’**échographie transthoracique**", "Échographie cardiaque")
    texte = LCM_NormaliserPhraseExamen(texte, "L'**echographie transthoracique**", "Échographie cardiaque")
    texte = LCM_NormaliserPhraseExamen(texte, "À l’échocardiographie", "Échocardiographie")
    texte = LCM_NormaliserPhraseExamen(texte, "A l'echocardiographie", "Échocardiographie")
    texte = LCM_NormaliserPhraseExamen(texte, "L’échocardiographie", "Échocardiographie")
    texte = LCM_NormaliserPhraseExamen(texte, "L'echocardiographie", "Échocardiographie")
    texte = LCM_NormaliserPhraseExamen(texte, "L’échographie cardiaque", "Échographie cardiaque")
    texte = LCM_NormaliserPhraseExamen(texte, "L'echographie cardiaque", "Échographie cardiaque")
    texte = LCM_NormaliserPhraseExamen(texte, "L’échographie transthoracique", "Échographie cardiaque")
    texte = LCM_NormaliserPhraseExamen(texte, "L'echographie transthoracique", "Échographie cardiaque")

    'Biologie.
    texte = LCM_NormaliserPhraseExamen(texte, "Sur le plan **biologique**", "Biologie")
    texte = LCM_NormaliserPhraseExamen(texte, "Le **bilan biologique**", "Bilan biologique")
    texte = LCM_NormaliserPhraseExamen(texte, "La **biologie**", "Biologie")
    texte = LCM_NormaliserPhraseExamen(texte, "Sur le plan biologique", "Biologie")
    texte = LCM_NormaliserPhraseExamen(texte, "Le bilan biologique", "Bilan biologique")
    texte = LCM_NormaliserPhraseExamen(texte, "La biologie", "Biologie")

    'Holter / MAPA.
    texte = LCM_NormaliserPhraseExamen(texte, "Le **Holter rythmique**", "Holter rythmique")
    texte = LCM_NormaliserPhraseExamen(texte, "Le **Holter ECG**", "Holter ECG")
    texte = LCM_NormaliserPhraseExamen(texte, "Le **Holter tensionnel**", "Holter tensionnel")
    texte = LCM_NormaliserPhraseExamen(texte, "La **MAPA**", "MAPA")
    texte = LCM_NormaliserPhraseExamen(texte, "Le Holter rythmique", "Holter rythmique")
    texte = LCM_NormaliserPhraseExamen(texte, "Le Holter ECG", "Holter ECG")
    texte = LCM_NormaliserPhraseExamen(texte, "Le Holter tensionnel", "Holter tensionnel")
    texte = LCM_NormaliserPhraseExamen(texte, "La MAPA", "MAPA")

    LCM_ForcerPhrasesNarrativesExamens = texte

End Function

Private Function LCM_NormaliserPhraseExamen( _
    ByVal texte As String, _
    ByVal motif As String, _
    ByVal libelleCanonique As String) As String

    Dim pos As Long
    Dim depart As Long
    Dim posSuite As Long
    Dim texteAvant As String
    Dim texteApres As String
    Dim marqueur As String
    Dim c As String

    depart = 1
    marqueur = "**" & libelleCanonique & "** :"

    Do
        pos = InStr(depart, texte, motif, vbTextCompare)
        If pos = 0 Then Exit Do

        If LCM_EstDebutUniteMedicale(texte, pos) Then

            posSuite = pos + Len(motif)

            'Ignorer espaces puis éventuelle virgule / deux-points après le motif.
            Do While posSuite <= Len(texte)
                c = Mid$(texte, posSuite, 1)
                If c <> " " And c <> vbTab Then Exit Do
                posSuite = posSuite + 1
            Loop

            If posSuite <= Len(texte) Then
                c = Mid$(texte, posSuite, 1)
                If c = "," Or c = ":" Then
                    posSuite = posSuite + 1
                End If
            End If

            Do While posSuite <= Len(texte)
                c = Mid$(texte, posSuite, 1)
                If c <> " " And c <> vbTab Then Exit Do
                posSuite = posSuite + 1
            Loop

            texteAvant = Left$(texte, pos - 1)
            texteApres = Mid$(texte, posSuite)

            'Retirer uniquement les espaces horizontaux avant le nouveau paragraphe.
            Do While Len(texteAvant) > 0
                c = Right$(texteAvant, 1)
                If c <> " " And c <> vbTab Then Exit Do
                texteAvant = Left$(texteAvant, Len(texteAvant) - 1)
            Loop

            If Len(texteAvant) > 0 Then
                If Right$(texteAvant, 1) <> vbCr Then
                    texteAvant = texteAvant & vbCr
                End If
            End If

            texte = texteAvant & marqueur
            If Len(texteApres) > 0 Then
                texte = texte & " " & texteApres
            End If

            depart = Len(texteAvant) + Len(marqueur) + 1

        Else
            depart = pos + Len(motif)
        End If

    Loop

    LCM_NormaliserPhraseExamen = texte

End Function

Private Function LCM_EstDebutUniteMedicale( _
    ByVal texte As String, _
    ByVal positionMotif As Long) As Boolean

    Dim j As Long
    Dim c As String

    LCM_EstDebutUniteMedicale = False

    If positionMotif <= 1 Then
        LCM_EstDebutUniteMedicale = True
        Exit Function
    End If

    j = positionMotif - 1

    Do While j >= 1
        c = Mid$(texte, j, 1)
        If c <> " " And c <> vbTab Then Exit Do
        j = j - 1
    Loop

    If j < 1 Then
        LCM_EstDebutUniteMedicale = True
        Exit Function
    End If

    c = Mid$(texte, j, 1)

    If c = vbCr _
    Or c = "." _
    Or c = "!" _
    Or c = "?" _
    Or c = ";" _
    Or c = ":" Then

        LCM_EstDebutUniteMedicale = True

    End If

End Function

Private Function LCM_NormaliserLibelleExamen( _
    ByVal texte As String, _
    ByVal libelle As String) As String

    Dim marqueur As String

    'Si l'API a laissé le libellé sans Markdown, on le remet en gras.
    texte = Replace( _
        texte, _
        libelle & " :", _
        "**" & libelle & "** :", _
        1, _
        -1, _
        vbTextCompare)

    texte = Replace( _
        texte, _
        libelle & ":", _
        "**" & libelle & "** :", _
        1, _
        -1, _
        vbTextCompare)

    'Normaliser aussi le Markdown déjà présent sans espace avant les deux-points.
    texte = Replace( _
        texte, _
        "**" & libelle & "**:", _
        "**" & libelle & "** :", _
        1, _
        -1, _
        vbTextCompare)

    'Le libellé doit commencer sur un nouveau paragraphe.
    marqueur = "**" & libelle & "** :"
    texte = LCM_InsererRetourAvantMarqueur( _
        texte, _
        marqueur)

    LCM_NormaliserLibelleExamen = texte

End Function

Private Function LCM_InsererRetourAvantMarqueur( _
    ByVal texte As String, _
    ByVal marqueur As String) As String

    Dim pos As Long
    Dim depart As Long

    depart = 1

    Do
        pos = InStr( _
            depart, _
            texte, _
            marqueur, _
            vbTextCompare)

        If pos = 0 Then Exit Do

        If pos > 1 Then
            If Mid$(texte, pos - 1, 1) <> vbCr Then
                texte = _
                    Left$(texte, pos - 1) & _
                    vbCr & _
                    Mid$(texte, pos)
                pos = pos + 1
            End If
        End If

        depart = pos + Len(marqueur)

    Loop

    LCM_InsererRetourAvantMarqueur = texte

End Function

'-------------------------------------------------------------------------------
' Normalisation du corps médical
'-------------------------------------------------------------------------------

Private Function LCM_NormaliserParagraphes( _
    ByVal texte As String) As String

    Dim lignes() As String
    Dim resultat As String
    Dim ligne As String
    Dim i As Long

    texte = Replace(texte, vbCrLf, vbLf)
    texte = Replace(texte, vbCr, vbLf)
    texte = Replace(texte, Chr$(11), vbLf)

    lignes = Split(texte, vbLf)

    resultat = ""

    For i = LBound(lignes) To UBound(lignes)

        ligne = _
            Trim$(Replace(CStr(lignes(i)), Chr$(160), " "))

        If ligne <> "" Then

            If resultat <> "" Then
                resultat = resultat & vbCr
            End If

            resultat = resultat & ligne

        End If

    Next i

    LCM_NormaliserParagraphes = resultat

End Function

Private Function LCM_OuvrirSansMacros(ByVal chemin As String) As Document
    Dim securite As Long, n As Long, description As String
    securite = Application.AutomationSecurity
    On Error GoTo Echec
    Application.AutomationSecurity = 3
    Set LCM_OuvrirSansMacros = Documents.Open(FileName:=chemin, ConfirmConversions:=False, ReadOnly:=True, AddToRecentFiles:=False, Visible:=True)
    Application.AutomationSecurity = securite
    Exit Function
Echec:
    n = Err.Number: description = Err.Description
    Application.AutomationSecurity = securite
    Err.Raise n, "LCM_OuvrirSansMacros", description
End Function
