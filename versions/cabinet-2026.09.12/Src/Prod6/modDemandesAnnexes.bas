Attribute VB_Name = "modDemandesAnnexes"
Option Explicit

'===============================================================================
' MODULE : modDemandesAnnexes
' VERSION : CABINET TEST v1 - lecture des demandes structurées
'
' Objet :
' - lire les blocs DEMANDE_DESTINATION conservés dans gReponseAPICabinetTest ;
' - extraire la clé et le corps de chaque demande ;
' - vérifier que les demandes restent disponibles après validation du courrier.
'===============================================================================

Public Function CT_ConstruireCollectionDemandes() As Collection

    Dim demandes As Collection
    Dim demande As Object

    Dim texteReponse As String
    Dim positionRecherche As Long
    Dim positionDebutBloc As Long
    Dim positionFinBloc As Long
    Dim longueurBloc As Long

    Dim blocComplet As String
    Dim cleDestination As String
    Dim corpsDestination As String

    Set demandes = New Collection

    texteReponse = gReponseAPICabinetTest

    If Trim$(texteReponse) = "" Then
        Set CT_ConstruireCollectionDemandes = demandes
        Exit Function
    End If

    positionRecherche = 1

    Do

        positionDebutBloc = InStr( _
            positionRecherche, _
            texteReponse, _
            BALISE_DEBUT_DEMANDE_DESTINATION, _
            vbTextCompare)

        If positionDebutBloc = 0 Then Exit Do

        positionFinBloc = InStr( _
            positionDebutBloc + Len(BALISE_DEBUT_DEMANDE_DESTINATION), _
            texteReponse, _
            BALISE_FIN_DEMANDE_DESTINATION, _
            vbTextCompare)

        If positionFinBloc = 0 Then Err.Raise vbObjectError + 967, , "Bloc de demande incomplet."

        longueurBloc = _
            positionFinBloc - positionDebutBloc + _
            Len(BALISE_FIN_DEMANDE_DESTINATION)

        blocComplet = Mid$( _
            texteReponse, _
            positionDebutBloc, _
            longueurBloc)

        cleDestination = CT_ExtraireCleDestination(blocComplet)
        corpsDestination = CT_ExtraireCorpsDestination(blocComplet)
        If Len(cleDestination) = 0 Or Len(corpsDestination) = 0 Then Err.Raise vbObjectError + 968, , "Demande sans destinataire ou corps."

        If Trim$(cleDestination) <> "" _
        And Trim$(corpsDestination) <> "" Then

            Set demande = CreateObject("Scripting.Dictionary")

            demande.Add "CleDestination", cleDestination
            demande.Add "CorpsDestination", corpsDestination
            demande.Add "BlocComplet", blocComplet

            demandes.Add demande

        End If

        positionRecherche = _
            positionFinBloc + _
            Len(BALISE_FIN_DEMANDE_DESTINATION)

    Loop

    Set CT_ConstruireCollectionDemandes = demandes

End Function

Private Function CT_ExtraireCleDestination( _
    ByVal bloc As String) As String

    Dim positionPrefixe As Long
    Dim positionValeur As Long
    Dim positionCR As Long
    Dim positionLF As Long
    Dim positionFinLigne As Long

    Dim cle As String

    CT_ExtraireCleDestination = ""

    positionPrefixe = InStr( _
        1, _
        bloc, _
        PREFIXE_CLE_DESTINATION, _
        vbTextCompare)

    If positionPrefixe = 0 Then Exit Function

    positionValeur = _
        positionPrefixe + _
        Len(PREFIXE_CLE_DESTINATION)

    positionCR = InStr( _
        positionValeur, _
        bloc, _
        vbCr, _
        vbBinaryCompare)

    positionLF = InStr( _
        positionValeur, _
        bloc, _
        vbLf, _
        vbBinaryCompare)

    If positionCR > 0 And positionLF > 0 Then

        If positionCR < positionLF Then
            positionFinLigne = positionCR
        Else
            positionFinLigne = positionLF
        End If

    ElseIf positionCR > 0 Then

        positionFinLigne = positionCR

    ElseIf positionLF > 0 Then

        positionFinLigne = positionLF

    Else

        positionFinLigne = Len(bloc) + 1

    End If

    cle = Mid$( _
        bloc, _
        positionValeur, _
        positionFinLigne - positionValeur)

    cle = Replace(cle, vbCr, "")
    cle = Replace(cle, vbLf, "")
    cle = Replace(cle, Chr(160), " ")
    cle = UCase$(Trim$(cle))

    CT_ExtraireCleDestination = cle

End Function

Private Function CT_ExtraireCorpsDestination( _
    ByVal bloc As String) As String

    Dim positionDebut As Long
    Dim positionTexte As Long
    Dim positionFin As Long

    Dim corps As String

    CT_ExtraireCorpsDestination = ""

    positionDebut = InStr( _
        1, _
        bloc, _
        BALISE_DEBUT_CORPS_DESTINATION, _
        vbTextCompare)

    If positionDebut = 0 Then Exit Function

    positionTexte = _
        positionDebut + _
        Len(BALISE_DEBUT_CORPS_DESTINATION)

    positionFin = InStr( _
        positionTexte, _
        bloc, _
        BALISE_FIN_CORPS_DESTINATION, _
        vbTextCompare)

    If positionFin = 0 Then Exit Function

    corps = Mid$( _
        bloc, _
        positionTexte, _
        positionFin - positionTexte)

    CT_ExtraireCorpsDestination = _
        CT_NettoyerBordsTexte(corps)

End Function

Private Function CT_NettoyerBordsTexte( _
    ByVal texte As String) As String

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

    CT_NettoyerBordsTexte = texte

End Function

Public Sub CT_TesterDemandesMemorisees()

    Dim demandes As Collection
    Dim demande As Object

    Dim message As String
    Dim apercuCorps As String
    Dim i As Long

    Set demandes = CT_ConstruireCollectionDemandes()

    If demandes.Count = 0 Then

        MsgBox _
            "Aucune demande structurée n'est actuellement mémorisée." & _
            vbCrLf & vbCrLf & _
            "Lancez une correction comportant une demande d'examen, " & _
            "puis validez le courrier.", _
            vbExclamation, _
            "Cabinet Test"

        Exit Sub

    End If

    message = _
        "Nombre de demandes mémorisées : " & _
        demandes.Count & _
        vbCrLf & vbCrLf

    For i = 1 To demandes.Count

        Set demande = demandes(i)

        apercuCorps = CStr(demande("CorpsDestination"))

        If Len(apercuCorps) > 350 Then
            apercuCorps = Left$(apercuCorps, 350) & " [...]"
        End If

        message = message & _
            "========================================" & vbCrLf & _
            "DEMANDE " & i & vbCrLf & _
            "Clé : " & _
            CStr(demande("CleDestination")) & _
            vbCrLf & vbCrLf & _
            apercuCorps & _
            vbCrLf & vbCrLf

    Next i

    If Len(message) > 6000 Then
        message = Left$(message, 6000) & _
                  vbCrLf & vbCrLf & _
                  "[Affichage tronqué]"
    End If

    MsgBox _
        message, _
        vbInformation, _
        "Demandes mémorisées — Cabinet Test"

End Sub

Public Sub CT_ForcerDestinationACompleterPourTest()

    Dim texteRecherche As String
    Dim texteRemplacement As String

    If Trim$(gReponseAPICabinetTest) = "" Then

        MsgBox _
            "Aucune réponse API n'est mémorisée." & vbCrLf & _
            "Effectuez et validez d'abord une correction Cabinet Test.", _
            vbExclamation, _
            "Test destination inconnue"

        Exit Sub

    End If

    texteRecherche = _
        PREFIXE_CLE_DESTINATION & "LOSHKAJIAN"

    texteRemplacement = _
        PREFIXE_CLE_DESTINATION & "A_COMPLETER"

    If InStr( _
        1, _
        gReponseAPICabinetTest, _
        texteRecherche, _
        vbTextCompare) = 0 Then

        MsgBox _
            "La clé LOSHKAJIAN n'a pas été trouvée dans la réponse mémorisée.", _
            vbExclamation, _
            "Test destination inconnue"

        Exit Sub

    End If

    gReponseAPICabinetTestAvantTest = _
        gReponseAPICabinetTest

    gReponseAPICabinetTest = Replace( _
        gReponseAPICabinetTest, _
        texteRecherche, _
        texteRemplacement, _
        1, _
        1, _
        vbTextCompare)

    MsgBox _
        "La destination Loshkajian a été temporairement remplacée " & _
        "par A_COMPLETER." & vbCrLf & vbCrLf & _
        "Lancez maintenant CT_CreerOuRecreerCourriersDemandes.", _
        vbInformation, _
        "Test destination inconnue"

End Sub

Public Sub CT_RestaurerReponseApresTestDestination()

    If Trim$(gReponseAPICabinetTestAvantTest) = "" Then

        MsgBox _
            "Aucune réponse antérieure n'est disponible.", _
            vbExclamation, _
            "Test destination inconnue"

        Exit Sub

    End If

    gReponseAPICabinetTest = _
        gReponseAPICabinetTestAvantTest

    gReponseAPICabinetTestAvantTest = ""

    MsgBox _
        "La réponse API d'origine a été restaurée.", _
        vbInformation, _
        "Test destination inconnue"

End Sub

Public Sub CT_RenseignerPremiereDestinationManquante()

    Dim texteRecherche As String
    Dim cleChoisie As String
    Dim positionCle As Long
    Dim recordDestination As String
    Dim nomDestinataire As String
    Dim reponse As VbMsgBoxResult

    If Trim$(gReponseAPICabinetTest) = "" Then

        MsgBox _
            "Aucune réponse API n'est mémorisée.", _
            vbExclamation, _
            "Cabinet Test"

        Exit Sub

    End If

    texteRecherche = _
        PREFIXE_CLE_DESTINATION & "A_COMPLETER"

    positionCle = InStr( _
        1, _
        gReponseAPICabinetTest, _
        texteRecherche, _
        vbTextCompare)

    If positionCle = 0 Then

        MsgBox _
            "Aucune destination à renseigner.", _
            vbInformation, _
            "Cabinet Test"

        Exit Sub

    End If

    '=========================================================
    ' Ouverture de la fenêtre conviviale de sélection.
    '=========================================================

    gCleDestinationChoisieCabinetTest = ""

    frmCTDestination.Show

    cleChoisie = _
        UCase$(Trim$(gCleDestinationChoisieCabinetTest))

    'L'utilisateur a cliqué sur Annuler.
    If cleChoisie = "" Then Exit Sub

    '=========================================================
    ' Vérification de la destination choisie dans Excel.
    '=========================================================

    recordDestination = _
        CT_DestinationParCle(cleChoisie)

    If Trim$(recordDestination) = "" Then

        MsgBox _
            "La destination sélectionnée n'est plus disponible " & _
            "dans la base Excel." & vbCrLf & vbCrLf & _
            "Aucune modification n'a été effectuée.", _
            vbExclamation, _
            "Cabinet Test"

        Exit Sub

    End If

    nomDestinataire = _
        ChampDestination(recordDestination, 2)

    '=========================================================
    ' Confirmation avant affectation.
    '=========================================================

    reponse = MsgBox( _
        "Destinataire sélectionné :" & vbCrLf & vbCrLf & _
        nomDestinataire & vbCrLf & _
        ChampDestination(recordDestination, 3) & _
        vbCrLf & vbCrLf & _
        "Affecter ce destinataire à la demande en attente ?", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Confirmer le destinataire")

    If reponse <> vbYes Then Exit Sub

    '=========================================================
    ' Remplacement d'une seule occurrence A_COMPLETER.
    ' Aucun nouvel appel API.
    '=========================================================

    gReponseAPICabinetTest = _
        Left$( _
            gReponseAPICabinetTest, _
            positionCle - 1) & _
        PREFIXE_CLE_DESTINATION & _
        cleChoisie & _
        Mid$( _
            gReponseAPICabinetTest, _
            positionCle + Len(texteRecherche))

    MsgBox _
        "La destination a été renseignée." & vbCrLf & vbCrLf & _
        "Destinataire : " & nomDestinataire & vbCrLf & _
        "Aucun nouvel appel à ChatGPT n'a été effectué." & vbCrLf & _
        "Le texte médical de la demande a été conservé." & vbCrLf & vbCrLf & _
        "Cliquez maintenant sur ""Créer demandes"" pour générer " & _
        "uniquement la demande débloquée.", _
        vbInformation, _
        "Cabinet Test"

End Sub
