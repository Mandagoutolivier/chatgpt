Attribute VB_Name = "modProdRapide"
Option Explicit

'===============================================================================
' MODULE : modProdRapide
' VERSION : DOMICILE-GRAS-1Z - double API conservé + gras local + 6 TAB
'
' Flux principal :
'   PR_CorrigerToutEnUnClic
'       -> anonymisation locale
'       -> appel API
'       -> insertion directe du courrier corrigé
'       -> résolution automatique des destinations connues
'       -> ouverture de frmCTDestination uniquement si une destination manque
'          avec affichage de la lettre complémentaire concernée
'       -> ajout des courriers d'examens dans le même document, chacun sur
'          une nouvelle page
'
' Principes :
' - sauvegarde automatique du document final dans le dossier SORTIE configure via modSortieDragon ;
' - aucune impression automatique ;
' - pas de boîte de dialogue en fonctionnement normal ;
' - les interventions sont réservées aux erreurs ou aux destinations inconnues ;
' - le gras des pages ajoutées est réappliqué caractère par caractère afin
'   d'éviter l'inversion observée lors de la fusion Word.
'===============================================================================
' V9 : la signature des lettres complémentaires reçoit 6 tabulations réelles
' après la fusion, au lieu de 8, afin d'éviter son retour partiel à la ligne.
'===============================================================================

Private Const PR_SIGNET_DEBUT_DEMANDES As String = _
    "PR_DEBUT_DEMANDES"

Private mDemandesMultipagesGenerees As Object

Public Sub PR_SupprimerDemandesDejaAjoutees( _
    ByVal doc As Document)

    Dim positionDebut As Long
    Dim rngSupprimer As Range

    If doc Is Nothing Then Exit Sub

    If Not doc.Bookmarks.Exists( _
        PR_SIGNET_DEBUT_DEMANDES) Then

        Exit Sub
    End If

    positionDebut = _
        doc.Bookmarks( _
            PR_SIGNET_DEBUT_DEMANDES).Range.Start

    doc.Bookmarks( _
        PR_SIGNET_DEBUT_DEMANDES).Delete

    Set rngSupprimer = doc.Range( _
        Start:=positionDebut, _
        End:=doc.Content.End - 1)

    rngSupprimer.Delete

End Sub

Private Sub PR_InitialiserSuiviMultipage()

    If mDemandesMultipagesGenerees Is Nothing Then
        Set mDemandesMultipagesGenerees = CreateObject("Scripting.Dictionary")
    End If

End Sub

Public Sub PR_ReinitialiserSuiviMultipage()

    Set mDemandesMultipagesGenerees = Nothing
    gDemandesMultipagesAjouteesProdRapide = False

End Sub

'===============================================================================
' GRAS STRUCTUREL DU COURRIER PRINCIPAL
'===============================================================================
' Cette routine ne s'applique qu'au corps du courrier principal.
' Elle est appelée après MPF_AppliquerMiseEnPageCourrier afin que le gras
' ne soit pas effacé par une remise en forme générale.


'===============================================================================
' SÉCURITÉ DU RÉFÉRENT PATIENT DANS LE COURRIER PRINCIPAL
'===============================================================================
Public Function PR_CorrigerReferentPatientOrientationCCN( _
    ByVal texte As String) As String

    Dim avant As Variant
    Dim apres As Variant
    Dim i As Long

    PR_CorrigerReferentPatientOrientationCCN = texte

    If Len(Trim$(texte)) = 0 Then Exit Function

    'Cette sécurité reste volontairement ciblée sur le CCN / Centre
    'Cardiologique du Nord afin de ne pas modifier un « vous » légitime ailleurs.
    If InStr(1, texte, "CCN", vbTextCompare) = 0 _
    And InStr(1, texte, "Centre Cardiologique du Nord", vbTextCompare) = 0 Then

        Exit Function

    End If

    avant = Array( _
        "Je vous oriente donc vers le CCN", _
        "Je vous oriente vers le CCN", _
        "Nous vous orientons donc vers le CCN", _
        "Nous vous orientons vers le CCN", _
        "Je vais vous orienter vers le CCN", _
        "Nous allons vous orienter vers le CCN", _
        "Je vous adresse donc au CCN", _
        "Je vous adresse au CCN", _
        "Nous vous adressons au CCN", _
        "Nous allons vous adresser au CCN", _
        "Je vous oriente donc vers le Centre Cardiologique du Nord", _
        "Je vous oriente vers le Centre Cardiologique du Nord", _
        "Nous vous orientons vers le Centre Cardiologique du Nord", _
        "Nous allons vous orienter vers le Centre Cardiologique du Nord", _
        "Je vous adresse au Centre Cardiologique du Nord", _
        "Nous allons vous adresser au Centre Cardiologique du Nord")

    apres = Array( _
        "Je l'oriente donc vers le CCN", _
        "Je l'oriente vers le CCN", _
        "Nous l'orientons donc vers le CCN", _
        "Nous l'orientons vers le CCN", _
        "Je vais l'orienter vers le CCN", _
        "Nous allons l'orienter vers le CCN", _
        "Je l'adresse donc au CCN", _
        "Je l'adresse au CCN", _
        "Nous l'adressons au CCN", _
        "Nous allons l'adresser au CCN", _
        "Je l'oriente donc vers le Centre Cardiologique du Nord", _
        "Je l'oriente vers le Centre Cardiologique du Nord", _
        "Nous l'orientons vers le Centre Cardiologique du Nord", _
        "Nous allons l'orienter vers le Centre Cardiologique du Nord", _
        "Je l'adresse au Centre Cardiologique du Nord", _
        "Nous allons l'adresser au Centre Cardiologique du Nord")

    For i = LBound(avant) To UBound(avant)

        texte = Replace( _
            texte, _
            CStr(avant(i)), _
            CStr(apres(i)), _
            1, _
            -1, _
            vbTextCompare)

    Next i

    PR_CorrigerReferentPatientOrientationCCN = texte

End Function

Public Sub TesterCorrectionReferentPatientCCN()

    Dim texteTest As String
    Dim resultat As String

    texteTest = _
        "Je vous oriente donc vers le CCN pour une prise en charge " & _
        "spécialisée en vue d'une ablation de fibrillation auriculaire."

    resultat = _
        PR_CorrigerReferentPatientOrientationCCN(texteTest)

    If InStr(1, resultat, _
        "Je l'oriente donc vers le CCN", _
        vbTextCompare) > 0 Then

        MsgBox _
            "OK : le patient reste à la troisième personne." & _
            vbCrLf & vbCrLf & resultat, _
            vbInformation, _
            "Test référent patient"

    Else

        MsgBox _
            "ÉCHEC : la formulation n'a pas été corrigée." & _
            vbCrLf & vbCrLf & resultat, _
            vbExclamation, _
            "Test référent patient"

    End If

End Sub

Private Function PR_DemandeMultipageDejaGeneree( _
    ByVal numeroDemande As Long) As Boolean

    PR_InitialiserSuiviMultipage

    PR_DemandeMultipageDejaGeneree = _
        mDemandesMultipagesGenerees.Exists(CStr(numeroDemande))

End Function

Private Sub PR_MarquerDemandeMultipageGeneree( _
    ByVal numeroDemande As Long)

    PR_InitialiserSuiviMultipage

    If Not mDemandesMultipagesGenerees.Exists(CStr(numeroDemande)) Then
        mDemandesMultipagesGenerees.Add CStr(numeroDemande), True
    End If

End Sub

Public Sub PR_CorrigerToutEnUnClic()
    modCycleCourrier.ExecuterCycleCourrier
End Sub
Public Sub PR_AjouterDemandesAuDocumentPrincipal()

    Dim demandes As Collection
    Dim demande As Object

    Dim docPrincipal As Document
    Dim docDemande As Document

    Dim rngCorpsPrincipal As Range
    Dim rngPremierPatientPrincipal As Range

    Dim cleDestination As String
    Dim corpsDestination As String
    Dim recordDestination As String

    Dim blocDestinataire As String
    Dim formuleAppel As String
    Dim formulePolitesse As String

    Dim nombreNonResolues As Long
    Dim nombrePretes As Long
    Dim totalDemandesAjoutees As Long
    Dim listeProblemes As String
    Dim messageCreation As String

    Dim i As Long
    Dim numeroErreur As Long
    Dim descriptionErreur As String
    Dim etapeMultipage As String, etatDestinations As Object

    On Error GoTo GestionErreur

    PR_InitialiserSuiviMultipage

    If gDemandesMultipagesAjouteesProdRapide Then
        Application.StatusBar = _
            "Document deja pret - demandes d'examens deja ajoutees."
        Exit Sub
    End If

    If Not gCorrectionCabinetTestValidee Then
        MsgBox "Le courrier principal doit d'abord être corrigé.", _
               vbExclamation, "Prod Rapide"
        Exit Sub
    End If

    If modCycleCourrier.DocumentSource() Is Nothing Then
        MsgBox "Le courrier principal n'est plus mémorisé.", _
               vbExclamation, "Prod Rapide"
        Exit Sub
    End If

    If Trim$(gReponseAPICabinetTest) = "" Then
        MsgBox "La réponse API structurée n'est plus disponible.", _
               vbExclamation, "Prod Rapide"
        Exit Sub
    End If

    Set demandes = CT_ConstruireCollectionDemandes()

    If demandes.Count = 0 Then
        gDemandesMultipagesAjouteesProdRapide = True
        Application.StatusBar = _
            "Correction terminee - aucune demande d'examen - document pret."
        Exit Sub
    End If

    Set docPrincipal = modCycleCourrier.DocumentSource()

    If Not LocaliserCorpsCourrier( _
        docPrincipal, _
        rngCorpsPrincipal, _
        rngPremierPatientPrincipal) Then

        MsgBox "Impossible de localiser le corps du courrier principal.", _
               vbExclamation, "Prod Rapide"
        Exit Sub
    End If

    Set etatDestinations = modEtatCourrier.Charger(docPrincipal, "")
    nombreNonResolues = 0
    nombrePretes = 0
    listeProblemes = ""

    '===========================================================================
    ' PHASE 1 : résoudre toutes les destinations avant de créer les lettres.
    '===========================================================================

    For i = 1 To demandes.Count

        Set demande = demandes(i)

        If PR_DemandeMultipageDejaGeneree(i) Then
            GoTo ResolutionSuivante
        End If

        cleDestination = _
            UCase$(Trim$(CStr(demande("CleDestination"))))

        'V8 : si le courrier demande une prise en charge rythmologique au CCN
        'sans nommer un médecin précis, ne jamais appliquer silencieusement
        'une destination de rythmologie par défaut. La secrétaire doit choisir.
        If PR_DoitForcerCorrespondantACompleterCCN( _
            gTexteAnonymise, _
            CStr(demande("CorpsDestination"))) And Not modEtatCourrier.DestinationConfirmee(etatDestinations, i, cleDestination) Then

            cleDestination = "A_COMPLETER"
            demande("CleDestination") = cleDestination

        End If

        recordDestination = _
            CT_DestinationParCle(cleDestination)

        If Trim$(recordDestination) = "" _
        Or Not PR_RecordDestinationComplet(recordDestination) Then

            Application.ScreenUpdating = True
            Application.StatusBar = _
                "Destination a renseigner - demande " & _
                i & "/" & demandes.Count & "..."

            If Not PR_ResoudreDestinationManquante( _
                demande, _
                i, _
                demandes.Count, _
                cleDestination, _
                recordDestination) Then

                nombreNonResolues = nombreNonResolues + 1
                listeProblemes = listeProblemes & _
                    "Demande " & i & _
                    " : destination non renseignée." & vbCrLf

                GoTo ResolutionSuivante
            End If

        End If

        If Not PR_RecordDestinationComplet(recordDestination) Then

            nombreNonResolues = nombreNonResolues + 1
            listeProblemes = listeProblemes & _
                "Demande " & i & _
                " : fiche destinataire incomplète." & vbCrLf

            GoTo ResolutionSuivante
        End If

        PR_StockerRecordDestination _
            demande, _
            recordDestination

        nombrePretes = nombrePretes + 1

ResolutionSuivante:

    Next i

    '===========================================================================
    ' PHASE 2 : toutes les destinations étant connues, créer toutes les lettres.
    '===========================================================================

    'Chaque lettre sera créée directement depuis lettremodeledragon.dotm.

    For i = 1 To demandes.Count

        Set demande = demandes(i)

        If PR_DemandeMultipageDejaGeneree(i) Then
            GoTo GenerationSuivante
        End If

        recordDestination = _
            PR_LireRecordDestination(demande)

        If Trim$(recordDestination) = "" Then

            'Une demande qui était prête en phase 1 ne doit jamais disparaître
            'silencieusement pendant la génération.
            nombreNonResolues = nombreNonResolues + 1
            listeProblemes = listeProblemes & _
                "Demande " & i & _
                " : fiche destinataire introuvable au moment de la génération." & vbCrLf

            GoTo GenerationSuivante
        End If

        If Not PR_RecordDestinationComplet(recordDestination) Then

            nombreNonResolues = nombreNonResolues + 1
            listeProblemes = listeProblemes & _
                "Demande " & i & _
                " : fiche destinataire devenue incomplète." & vbCrLf

            GoTo GenerationSuivante
        End If

        cleDestination = _
            UCase$(Trim$(CStr(demande("CleDestination"))))

        corpsDestination = _
            CStr(demande("CorpsDestination"))

        corpsDestination = _
            MPF_NormaliserRetoursTexte(corpsDestination)

        blocDestinataire = _
            Trim$(ChampDestination(recordDestination, 4))

        formuleAppel = _
            Trim$(ChampDestination(recordDestination, 5))

        formulePolitesse = _
            Trim$(ChampDestination(recordDestination, 11))

        messageCreation = ""
        etapeMultipage = "Création de la lettre complémentaire depuis le modèle"

        Set docDemande = _
            LCM_CreerLettreDepuisModele( _
                docPrincipal, _
                blocDestinataire, _
                formuleAppel, _
                corpsDestination, _
                formulePolitesse, _
                messageCreation)

        If docDemande Is Nothing Then

            nombreNonResolues = nombreNonResolues + 1
            listeProblemes = listeProblemes & _
                "Demande " & i & _
                " : création depuis lettremodeledragon.dotm impossible." & _
                vbCrLf & messageCreation & vbCrLf

            GoTo GenerationSuivante
        End If

        Call modGras.AppliquerGrasDocumentComplet(docDemande)
        etapeMultipage = "Préparation de l'ajout au document principal"

        Dim corAnnexe As Object, paramAnnexe As Object
        Set paramAnnexe = modServiceNas.Parametres(): paramAnnexe("cle") = cleDestination
        Set corAnnexe = modServiceNas.Appeler("correspondent.resolve", paramAnnexe)
        PR_AjouterDocumentEnNouvellePage _
            docPrincipal, _
            docDemande, _
            etapeMultipage, i, CStr(corAnnexe("ID"))

        etapeMultipage = "Marquage de la demande comme générée"
        PR_MarquerDemandeMultipageGeneree i

        docDemande.Close SaveChanges:=wdDoNotSaveChanges
        Set docDemande = Nothing

        docPrincipal.Activate

        If LocaliserCorpsCourrier( _
            docPrincipal, _
            rngCorpsPrincipal, _
            rngPremierPatientPrincipal) Then

            Set gPlageOriginale = _
                rngCorpsPrincipal.Duplicate
        End If

GenerationSuivante:

    Next i

    docPrincipal.Activate

    If LocaliserCorpsCourrier( _
        docPrincipal, _
        rngCorpsPrincipal, _
        rngPremierPatientPrincipal) Then

        Set gPlageOriginale = _
            rngCorpsPrincipal.Duplicate
    End If

    totalDemandesAjoutees = _
        mDemandesMultipagesGenerees.Count

    'Une fois toutes les erreurs explicites éliminées, la présence d'au moins
    'une lettre effectivement ajoutée suffit à considérer le document comme prêt.
    'On ne dépend plus de l'égalité avec demandes.Count, qui peut être plus
    'restrictive que le nombre réel de lettres finales produites.
    If nombreNonResolues = 0 _
    And totalDemandesAjoutees = demandes.Count Then

        gDemandesMultipagesAjouteesProdRapide = True

        Application.StatusBar = _
            "Correction terminee - " & _
            totalDemandesAjoutees & _
            " demande(s) ajoutee(s) - document pret."

    Else

        gDemandesMultipagesAjouteesProdRapide = False

        Application.StatusBar = _
            "Correction terminee - " & _
            totalDemandesAjoutees & _
            " demande(s) ajoutee(s) - " & _
            nombreNonResolues & _
            " demande(s) non créée(s)."

        If Trim$(listeProblemes) <> "" Then

            MsgBox _
                "Certaines lettres complémentaires n'ont pas été créées :" & _
                vbCrLf & vbCrLf & _
                listeProblemes, _
                vbExclamation, _
                "Prod Rapide"

        End If

    End If

    Exit Sub

GestionErreur:

    numeroErreur = Err.Number
    descriptionErreur = Err.description

    On Error Resume Next

    If Not docDemande Is Nothing Then
        docDemande.Close SaveChanges:=wdDoNotSaveChanges
    End If

    If Not docPrincipal Is Nothing Then
        docPrincipal.Activate
    End If

    On Error GoTo 0

    MsgBox "Erreur pendant la création du document multipage :" & vbCrLf & _
           numeroErreur & " - " & descriptionErreur & vbCrLf & vbCrLf & _
           "Étape : " & etapeMultipage & vbCrLf & _
           "Documents ouverts : " & Application.Documents.Count & vbCrLf & _
           "Fenêtres Word : " & Application.Windows.Count, _
           vbExclamation, "Prod Rapide"

End Sub

Private Sub PR_StockerRecordDestination( _
    ByVal demande As Object, _
    ByVal recordDestination As String)

    If demande Is Nothing Then Exit Sub

    If demande.Exists("RecordDestination") Then

        demande("RecordDestination") = _
            recordDestination

    Else

        demande.Add _
            "RecordDestination", _
            recordDestination

    End If

End Sub

Private Function PR_LireRecordDestination( _
    ByVal demande As Object) As String

    PR_LireRecordDestination = ""

    If demande Is Nothing Then Exit Function

    If demande.Exists("RecordDestination") Then

        PR_LireRecordDestination = _
            CStr(demande("RecordDestination"))

    End If

End Function

Private Function PR_RecordDestinationComplet( _
    ByVal recordDestination As String) As Boolean

    PR_RecordDestinationComplet = False

    If Trim$(recordDestination) = "" Then _
        Exit Function

    If Trim$(ChampDestination( _
        recordDestination, 4)) = "" Then _
        Exit Function

    If Trim$(ChampDestination( _
        recordDestination, 5)) = "" Then _
        Exit Function

    If Trim$(ChampDestination( _
        recordDestination, 11)) = "" Then _
        Exit Function

    PR_RecordDestinationComplet = True

End Function

Private Function PR_ChampsDestinationManquants( _
    ByVal recordDestination As String) As String

    Dim resultat As String

    resultat = ""

    If Trim$(recordDestination) = "" Then
        PR_ChampsDestinationManquants = _
            "fiche destinataire introuvable"
        Exit Function
    End If

    If Trim$(ChampDestination( _
        recordDestination, 4)) = "" Then

        resultat = resultat & _
            "- bloc destinataire" & vbCrLf
    End If

    If Trim$(ChampDestination( _
        recordDestination, 5)) = "" Then

        resultat = resultat & _
            "- formule d'appel" & vbCrLf
    End If

    If Trim$(ChampDestination( _
        recordDestination, 11)) = "" Then

        resultat = resultat & _
            "- formule de politesse" & vbCrLf
    End If

    PR_ChampsDestinationManquants = resultat

End Function

Public Sub TesterForcageCorrespondantCCN()

    Dim sourceTest As String
    Dim corpsTest As String

    sourceTest = _
        "Nous allons donc la diriger vers le CCN pour une ablation par radiofréquence."

    corpsTest = _
        "Je vous adresse Madame [PATIENT] pour une prise en charge rythmologique " & _
        "en vue d'une ablation par radiofréquence."

    If PR_DoitForcerCorrespondantACompleterCCN( _
        sourceTest, _
        corpsTest) Then

        MsgBox _
            "OK : CCN sans médecin nommé => A_COMPLETER et choix secrétaire.", _
            vbInformation, _
            "Test correspondant CCN"
    Else
        MsgBox _
            "ÉCHEC : le choix manuel du correspondant ne serait pas forcé.", _
            vbExclamation, _
            "Test correspondant CCN"
    End If

End Sub

Public Sub TesterAbsenceForcageCCNHistorique()

    Dim sourceTest As String
    Dim corpsTest As String

    sourceTest = _
        "Une ablation par radiofréquence avait été discutée au CCN il y a deux ans, " & _
        "sans indication retenue. Le suivi est actuellement médical."

    corpsTest = _
        "Contexte rythmologique ancien avec discussion d'ablation."

    If PR_DoitForcerCorrespondantACompleterCCN( _
        sourceTest, _
        corpsTest) Then

        MsgBox _
            "ECHEC : une simple mention historique du CCN force encore un correspondant.", _
            vbExclamation, _
            "Test absence de faux positif CCN"
    Else
        MsgBox _
            "OK : CCN / ablation historiques sans orientation explicite ne forcent rien.", _
            vbInformation, _
            "Test absence de faux positif CCN"
    End If

End Sub

Public Function PR_DoitForcerCorrespondantACompleterCCN( _
    ByVal texteSource As String, _
    ByVal corpsDemande As String) As Boolean

    Dim src As String
    Dim corps As String
    Dim posCCN As Long
    Dim debutContexte As Long
    Dim longueurContexte As Long
    Dim contexte As String

    PR_DoitForcerCorrespondantACompleterCCN = False

    src = LCase$(texteSource)
    corps = LCase$(corpsDemande)

    src = Replace(Replace(src, ChrW(8217), "'"), ChrW(8216), "'")
    corps = Replace(Replace(corps, ChrW(8217), "'"), ChrW(8216), "'")

    'La demande générée doit réellement concerner la rythmologie / une ablation.
    If InStr(1, corps, "ablation", vbTextCompare) = 0 _
    And InStr(1, corps, "rythmolog", vbTextCompare) = 0 _
    And InStr(1, corps, "radiofréquence", vbTextCompare) = 0 _
    And InStr(1, corps, "radiofrequence", vbTextCompare) = 0 Then

        Exit Function

    End If

    posCCN = InStr(1, src, "ccn", vbTextCompare)
    If posCCN = 0 Then Exit Function

    'Examiner environ 120 caractères de part et d'autre de CCN.
    debutContexte = posCCN - 120
    If debutContexte < 1 Then debutContexte = 1

    longueurContexte = 240
    If debutContexte + longueurContexte - 1 > Len(src) Then
        longueurContexte = Len(src) - debutContexte + 1
    End If

    contexte = Mid$(src, debutContexte, longueurContexte)

    'V7 : la simple présence de CCN, d'une ablation ou de rythmologie dans les
    'antécédents ne suffit jamais. Il faut une orientation explicite dans le
    'même contexte avant de forcer le choix manuel du correspondant.
    If Not PR_ContexteContientOrientationExpliciteCCN(contexte) Then
        Exit Function
    End If

    'Si un médecin est explicitement nommé autour de CCN, laisser le flux normal.
    If InStr(1, contexte, "docteur ", vbTextCompare) > 0 _
    Or InStr(1, contexte, "dr ", vbTextCompare) > 0 _
    Or InStr(1, contexte, "professeur ", vbTextCompare) > 0 _
    Or InStr(1, contexte, "pr ", vbTextCompare) > 0 Then

        Exit Function

    End If

    PR_DoitForcerCorrespondantACompleterCCN = True

End Function

Private Function PR_ContexteContientOrientationExpliciteCCN( _
    ByVal contexte As String) As Boolean

    Dim futurPremierePersonne As Boolean
    Dim verbeOrientation As Boolean

    PR_ContexteContientOrientationExpliciteCCN = False

    If InStr(1, contexte, "ccn", vbTextCompare) = 0 Then Exit Function

    'Intentions explicites directement suffisantes.
    If InStr(1, contexte, "je propose", vbTextCompare) > 0 _
    Or InStr(1, contexte, "nous proposons", vbTextCompare) > 0 _
    Or InStr(1, contexte, "je préconise", vbTextCompare) > 0 _
    Or InStr(1, contexte, "je preconise", vbTextCompare) > 0 _
    Or InStr(1, contexte, "je recommande", vbTextCompare) > 0 _
    Or InStr(1, contexte, "nous recommandons", vbTextCompare) > 0 _
    Or InStr(1, contexte, "je l'adresse", vbTextCompare) > 0 _
    Or InStr(1, contexte, "je l'oriente", vbTextCompare) > 0 _
    Or InStr(1, contexte, "nous l'orientons", vbTextCompare) > 0 _
    Or InStr(1, contexte, "je le confie", vbTextCompare) > 0 _
    Or InStr(1, contexte, "je la confie", vbTextCompare) > 0 _
    Or InStr(1, contexte, "nous le confions", vbTextCompare) > 0 _
    Or InStr(1, contexte, "nous la confions", vbTextCompare) > 0 Then

        PR_ContexteContientOrientationExpliciteCCN = True
        Exit Function

    End If

    'Pour une formulation au futur, demander aussi un verbe d'orientation. Cela
    'reconnaît notamment "nous allons donc la diriger vers le CCN" sans rendre
    'la simple expression "nous allons poursuivre..." déclenchante.
    futurPremierePersonne = _
        (InStr(1, contexte, "nous allons", vbTextCompare) > 0) _
        Or (InStr(1, contexte, "je vais", vbTextCompare) > 0)

    verbeOrientation = _
        (InStr(1, contexte, "diriger", vbTextCompare) > 0) _
        Or (InStr(1, contexte, "adresser", vbTextCompare) > 0) _
        Or (InStr(1, contexte, "orienter", vbTextCompare) > 0) _
        Or (InStr(1, contexte, "confier", vbTextCompare) > 0)

    If futurPremierePersonne And verbeOrientation Then
        PR_ContexteContientOrientationExpliciteCCN = True
    End If

End Function

Private Function PR_ResoudreDestinationManquante( _
    ByVal demande As Object, _
    ByVal numeroDemande As Long, _
    ByVal nombreDemandes As Long, _
    ByRef cleDestination As String, _
    ByRef recordDestination As String) As Boolean

    Dim cleChoisie As String
    Dim blocOriginal As String
    Dim blocModifie As String
    Dim champsManquants As String
    Dim messageACompleterAffiche As Boolean

    PR_ResoudreDestinationManquante = False
    messageACompleterAffiche = False

NouvelleSelection:

    If Not messageACompleterAffiche _
    And (UCase$(Trim$(cleDestination)) = "A_COMPLETER" _
         Or Trim$(cleDestination) = "") Then

        MsgBox _
            "Un courrier complémentaire doit être créé, mais aucun " & _
            "correspondant précis n'a été identifié dans le courrier." & _
            vbCrLf & vbCrLf & _
            "Merci de sélectionner dans la fenêtre suivante le médecin " & _
            "ou correspondant concerné.", _
            vbInformation, _
            "Correspondant à préciser"

        messageACompleterAffiche = True

    End If

    gCleDestinationChoisieCabinetTest = ""

    PR_PreparerFormulaireDestination _
        CStr(demande("CorpsDestination")), _
        numeroDemande, _
        nombreDemandes

    frmCTDestination.Show

    cleChoisie = _
        UCase$(Trim$( _
            gCleDestinationChoisieCabinetTest))

    If cleChoisie = "" Then _
        Exit Function

    recordDestination = _
        CT_DestinationParCle(cleChoisie)

    If Trim$(recordDestination) = "" Then

        MsgBox _
            "Le destinataire sélectionné n'est plus disponible " & _
            "dans la base Excel." & vbCrLf & vbCrLf & _
            "Choisissez un autre destinataire pour cette demande.", _
            vbExclamation, _
            "Prod Rapide"

        GoTo NouvelleSelection
    End If

    If Not PR_RecordDestinationComplet( _
        recordDestination) Then

        champsManquants = _
            PR_ChampsDestinationManquants( _
                recordDestination)

        MsgBox _
            "La fiche du destinataire sélectionné est incomplète " & _
            "et ne permet pas de créer la lettre." & vbCrLf & vbCrLf & _
            champsManquants & vbCrLf & _
            "Choisissez un autre destinataire pour cette demande.", _
            vbExclamation, _
            "Prod Rapide"

        GoTo NouvelleSelection
    End If

    Dim structure As Object, etat As Object
    Set structure = modJson.JsonParse(gReponseAPICabinetTest)
    structure("demandes")(numeroDemande)("cle_destination") = cleChoisie
    demande("CleDestination") = cleChoisie
    gReponseAPICabinetTest = modServiceNas.JsonValeur(structure)
    Set etat = modEtatCourrier.Charger(modCycleCourrier.DocumentSource(), "")
    etat("reponse") = gReponseAPICabinetTest
    If Not etat.Exists("destinations_confirmees") Then Set etat("destinations_confirmees") = modServiceNas.Parametres()
    etat("destinations_confirmees")(CStr(numeroDemande)) = cleChoisie
    modEtatCourrier.Sauver modCycleCourrier.DocumentSource(), etat

    cleDestination = cleChoisie

    PR_ResoudreDestinationManquante = True

End Function

Private Sub PR_PreparerFormulaireDestination( _
    ByVal corpsDestination As String, _
    ByVal numeroDemande As Long, _
    ByVal nombreDemandes As Long)

    Const DECALAGE_VERTICAL As Single = 95
    Const MARGE As Single = 12

    Dim frm As Object
    Dim lblContexte As Object
    Dim txtContexte As Object
    Dim texteAffiche As String

    Set frm = frmCTDestination

    frm.Caption = _
        "Choisir le destinataire - demande " & _
        numeroDemande & "/" & nombreDemandes

    On Error Resume Next
    Set lblContexte = frm.Controls("lblPRDemandeContexte")
    Set txtContexte = frm.Controls("txtPRDemandeContexte")
    On Error GoTo 0

    'Le formulaire peut rester chargé après un Hide.
    'Dans ce cas, on ne décale les contrôles qu'une seule fois.
    If txtContexte Is Nothing Then

        frm.lblRecherche.Top = _
            frm.lblRecherche.Top + DECALAGE_VERTICAL

        frm.txtRecherche.Top = _
            frm.txtRecherche.Top + DECALAGE_VERTICAL

        frm.lstCorrespondants.Top = _
            frm.lstCorrespondants.Top + DECALAGE_VERTICAL

        frm.cmdUtiliser.Top = _
            frm.cmdUtiliser.Top + DECALAGE_VERTICAL

        frm.cmdAnnuler.Top = _
            frm.cmdAnnuler.Top + DECALAGE_VERTICAL

        frm.Height = frm.Height + DECALAGE_VERTICAL

        Set lblContexte = frm.Controls.Add( _
            "Forms.Label.1", _
            "lblPRDemandeContexte", _
            True)

        With lblContexte
            .Caption = "Sélectionnez le correspondant concerné pour cette lettre :"
            .Left = MARGE
            .Top = 8
            .Width = frm.Width - (2 * MARGE) - 8
            .Height = 16
            .Font.Bold = True
        End With

        Set txtContexte = frm.Controls.Add( _
            "Forms.TextBox.1", _
            "txtPRDemandeContexte", _
            True)

        With txtContexte
            .Left = MARGE
            .Top = 26
            .Width = frm.Width - (2 * MARGE) - 8
            .Height = 58
            .MultiLine = True
            .WordWrap = True
            .Locked = True
            .TabStop = False
            .ScrollBars = 2
            .BackColor = &H8000000F
            .Font.Size = 9
        End With

    End If

    If Not lblContexte Is Nothing Then
        lblContexte.Caption = _
            "Sélectionnez le correspondant concerné pour cette lettre :"
    End If

    texteAffiche = _
        PR_TexteDemandePourAffichage(corpsDestination)

    txtContexte.Text = texteAffiche

    On Error Resume Next
    txtContexte.SelStart = 0
    frm.txtRecherche.SetFocus
    On Error GoTo 0

End Sub

Private Function PR_TexteDemandePourAffichage( _
    ByVal texte As String) As String

    texte = Replace( _
        texte, _
        MARQUEUR_PATIENT, _
        "[PATIENT]", _
        1, _
        -1, _
        vbTextCompare)

    texte = Replace(texte, "**", "")
    texte = Trim$(texte)

    If Len(texte) > 900 Then
        texte = Left$(texte, 900) & "..."
    End If

    PR_TexteDemandePourAffichage = texte

End Function

Private Function PR_RemplacerCleDansBloc( _
    ByVal bloc As String, _
    ByVal nouvelleCle As String) As String

    Dim posPrefixe As Long
    Dim posValeur As Long
    Dim posCR As Long
    Dim posLF As Long
    Dim posFinLigne As Long

    PR_RemplacerCleDansBloc = bloc

    posPrefixe = InStr( _
        1, _
        bloc, _
        PREFIXE_CLE_DESTINATION, _
        vbTextCompare)

    If posPrefixe = 0 Then Exit Function

    posValeur = posPrefixe + Len(PREFIXE_CLE_DESTINATION)

    posCR = InStr(posValeur, bloc, vbCr, vbBinaryCompare)
    posLF = InStr(posValeur, bloc, vbLf, vbBinaryCompare)

    If posCR > 0 And posLF > 0 Then

        If posCR < posLF Then
            posFinLigne = posCR
        Else
            posFinLigne = posLF
        End If

    ElseIf posCR > 0 Then

        posFinLigne = posCR

    ElseIf posLF > 0 Then

        posFinLigne = posLF

    Else

        Exit Function
    End If

    PR_RemplacerCleDansBloc = _
        Left$(bloc, posValeur - 1) & _
        UCase$(Trim$(nouvelleCle)) & _
        Mid$(bloc, posFinLigne)

End Function

Private Sub PR_AjouterDocumentEnNouvellePage( _
    ByVal docDestination As Document, _
    ByVal docSource As Document, _
    ByRef etapeMultipage As String, _
    Optional ByVal numeroAnnexe As Long = 0, _
    Optional ByVal destinataireID As String = "")

    Dim rngInsertion As Range
    Dim rngSource As Range

    Dim debutInsertion As Long
    Dim positionSaut As Long

    'MODELE-1F :
    'La lettre source vient désormais directement de lettremodeledragon.dotm.
    'On ne modifie donc PLUS la police ni le gras après collage.
    '
    'Le code précédent forçait toute la plage ajoutée avec Font.Name,
    'Font.NameAscii, Font.NameFarEast et Font.NameOther puis parcourait tous
    'les caractères pour refaire le gras. Ces opérations sont inutiles avec
    'le nouveau modèle et peuvent provoquer l'erreur 5844, notamment sur une
    'plage contenant l'image de signature.
    '
    'On revient au collage Word "mise en forme d'origine", qui est précisément
    'prévu pour conserver la mise en forme du document source.

    etapeMultipage = "Fusion : calcul du point de saut de page"

    positionSaut = _
        docDestination.Content.End - 1

    etapeMultipage = "Fusion : création de la plage de saut"

    Set rngInsertion = docDestination.Range( _
        Start:=positionSaut, _
        End:=positionSaut)

    etapeMultipage = "Fusion : insertion du saut de page"

    rngInsertion.InsertBreak _
        Type:=wdPageBreak

    'Le premier saut de page marque le début des courriers complémentaires.
    etapeMultipage = "Fusion : création du signet de début des demandes"

    If Not docDestination.Bookmarks.Exists( _
        PR_SIGNET_DEBUT_DEMANDES) Then

        docDestination.Bookmarks.Add _
            Name:=PR_SIGNET_DEBUT_DEMANDES, _
            Range:=docDestination.Range( _
                Start:=positionSaut, _
                End:=positionSaut + 1)

    End If

    etapeMultipage = "Fusion : calcul du point d'insertion"

    debutInsertion = _
        docDestination.Content.End - 1

    Set rngInsertion = docDestination.Range( _
        Start:=debutInsertion, _
        End:=debutInsertion)

    etapeMultipage = "Fusion : préparation de la lettre source"

    Set rngSource = docSource.Content.Duplicate

    'Ne pas copier la marque finale du document source.
    If rngSource.End > rngSource.Start Then
        rngSource.End = rngSource.End - 1
    End If

    etapeMultipage = "Fusion : copie de la lettre source"

    docSource.Activate

    etapeMultipage = "Fusion : activation du courrier principal"

    docDestination.Activate
    DoEvents

    'Recréer le Range après l'activation pour éviter tout état Word ambigu.
    Set rngInsertion = docDestination.Range( _
        Start:=debutInsertion, _
        End:=debutInsertion)

    etapeMultipage = "Fusion : collage avec mise en forme d'origine"

    rngInsertion.FormattedText = rngSource.FormattedText

    'La marque de paragraphe finale du document source n'est pas copiée.
    'La signature étant le dernier paragraphe, son retrait de paragraphe
    'peut donc être perdu au collage. On positionne ici la signature
    'VISIBLE, après fusion, avec six tabulations réelles.
    etapeMultipage = "Fusion : positionnement de la signature"

    PR_PositionnerSignatureAjouteeAvec6Tabs _
        docDestination, _
        debutInsertion

    If numeroAnnexe > 0 Then modIndexAnnexes.IndexerAnnexe docDestination, docSource, debutInsertion, numeroAnnexe, destinataireID
    etapeMultipage = "Fusion : terminée"

End Sub

Private Sub PR_PositionnerSignatureAjouteeAvec6Tabs( _
    ByVal doc As Document, _
    ByVal debutZoneAjoutee As Long)

    Dim i As Long
    Dim texte As String
    Dim rngDebut As Range
    Dim nom As String, prenom As String

    If doc Is Nothing Then Exit Sub
    nom = Trim$(modConfig.Config("MEDECIN", "Nom", ""))
    prenom = Trim$(modConfig.Config("MEDECIN", "Prenom", ""))
    If Len(nom) = 0 Or Len(prenom) = 0 Then Exit Sub

    'Recherche uniquement dans la lettre qui vient d'être ajoutée,
    'en partant de la fin afin de ne jamais toucher l'en-tête.
    For i = doc.Paragraphs.Count To 1 Step -1

        If doc.Paragraphs(i).Range.Start < debutZoneAjoutee Then
            Exit For
        End If

        texte = doc.Paragraphs(i).Range.Text
        texte = Replace(texte, vbCr, "")
        texte = Replace(texte, vbLf, "")
        texte = Replace(texte, Chr$(7), "")
        texte = Replace(texte, Chr$(11), " ")

        If InStr(1, texte, nom, vbTextCompare) > 0 _
        And InStr(1, texte, prenom, vbTextCompare) > 0 Then

            'Aucun retrait de paragraphe : la position dépend uniquement
            'des six caractères TAB présents dans le texte final.
            With doc.Paragraphs(i).Format
                .LeftIndent = 0
                .RightIndent = 0
                .FirstLineIndent = 0
                .Alignment = wdAlignParagraphLeft
            End With

            Set rngDebut = doc.Range( _
                Start:=doc.Paragraphs(i).Range.Start, _
                End:=doc.Paragraphs(i).Range.Start)

            rngDebut.Text = _
                vbTab & vbTab & vbTab & vbTab & _
                vbTab & vbTab

            Exit Sub

        End If

    Next i

End Sub


Public Function PR_EnCours() As Boolean
    PR_EnCours = modCycleCourrier.CycleEnCours()
End Function
