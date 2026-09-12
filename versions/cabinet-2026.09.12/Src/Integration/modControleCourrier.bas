Attribute VB_Name = "modControleCourrier"
Option Explicit

Public Sub PreparerRelecture(ByVal doc As Document, ByVal source As String, ByVal resultat As String)
    Dim p As Object, r As Object, difference As Object, message As String
    Set p = modServiceNas.Parametres(): p("source") = source: p("resultat") = resultat
    Set r = modServiceNas.Appeler("clinical.compare", p)
    For Each difference In r("differences")
        message = message & vbCrLf & "- " & CStr(difference("controle"))
        Dim valeur As Variant
        For Each valeur In difference("retires"): message = message & " [retire: " & CStr(valeur) & "]": Next valeur
        For Each valeur In difference("ajoutes"): message = message & " [ajoute: " & CStr(valeur) & "]": Next valeur
    Next difference
    modIntegrationUnifie.FixerVariable doc, "RelectureEnAttente", "1"
    modIntegrationUnifie.FixerVariable doc, "RelectureValidee", "0"
    modIntegrationUnifie.FixerVariable doc, "ControlesRelecture", message
    Application.ScreenUpdating = True
    MsgBox "Le courrier et ses annexes sont prets pour votre relecture." & vbCrLf & _
        IIf(Len(message) > 0, "Differences detectees :" & message & vbCrLf, "") & _
        "Verifiez identite, destinataires, negations, doses et examens. Appuyez de nouveau sur D apres relecture pour transmettre.", vbInformation, "Relecture du courrier"
End Sub

Public Sub ValiderEtTransmettre(ByVal doc As Document)
    Dim cor As Object, id As String, f As ufListe, recherche As String, sortie As String
    id = Trim$(modIntegrationUnifie.VariableDoc(doc, "CorrespondantID"))
    If Len(id) = 0 Then
        recherche = Trim$(InputBox("Nom du destinataire principal dicte (selection d un identifiant stable) :", "Destinataire"))
        If Len(recherche) < 2 Then Exit Sub
        Set f = New ufListe
        f.Configurer "Confirmer le destinataire", modServiceNas.LireTable("CORRESPONDANTS", recherche), Array("Nom", "Prenom", "Adresse1", "Ville"), "110 pt;90 pt;180 pt;100 pt"
        f.Show vbModal
        If Not f.Annule Then Set cor = f.Resultat
        Unload f
        If cor Is Nothing Then Exit Sub
        modIntegrationUnifie.FixerVariable doc, "CorrespondantID", CStr(cor("ID"))
    Else
        Set cor = modBase.CorrespondantParID(id)
    End If
    If MsgBox("Confirmez-vous la relecture de toutes les pages et le destinataire principal : " & CStr(cor("Nom")) & " " & CStr(cor("Prenom")) & ", " & CStr(cor("Ville")) & " ?", vbYesNo + vbQuestion, "Transmettre au secretariat") <> vbYes Then Exit Sub
    modCourrier.RemplirSignet doc, "DESTINATAIRE", CStr(cor("BlocDestinataire"))
    modCourrier.MettreEnFormeDestinataire doc
    ' Chaque confirmation du texte actuel ouvre une revision distincte, meme apres un delai reseau.
    ' La consultation reste identique : aucun nouvel acte comptable.
    modIntegrationUnifie.FixerVariable doc, "PublicationID", modFichiers.IdUnique()
    modIntegrationUnifie.FixerVariable doc, "DateValidation", Format$(Now, "dd/mm/yyyy hh:nn:ss")
    modIntegrationUnifie.FixerVariable doc, "RelectureValidee", "1"
    If Not SD_EnregistrerCourrierFinal(doc, sortie) Then Err.Raise vbObjectError + 1130, , "Enregistrement final interrompu."
    modIntegrationUnifie.TransmettreSecretariat doc, sortie
    modIntegrationUnifie.FixerVariable doc, "RelectureEnAttente", "0"
    doc.Save
    MsgBox "Courrier transmis au secretariat.", vbInformation, "Cabinet"
End Sub
