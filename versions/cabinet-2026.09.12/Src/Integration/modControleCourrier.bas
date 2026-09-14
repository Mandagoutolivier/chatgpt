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
    Dim cor As Object, id As String, f As ufListe, recherche As String, sortie As String, revision As String
    modIntegrationUnifie.InitialiserPatientProd doc
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
    revision = EmpreinteCourrier(doc, CStr(cor("ID")))
    If Trim$(modIntegrationUnifie.VariableDoc(doc, "PublicationRevision")) <> revision Then
        modIntegrationUnifie.FixerVariable doc, "PublicationID", modFichiers.IdUnique()
        modIntegrationUnifie.FixerVariable doc, "PublicationRevision", revision
        modIntegrationUnifie.FixerVariable doc, "DateValidation", Format$(Now, "dd/mm/yyyy hh:nn:ss")
        modIntegrationUnifie.FixerVariable doc, "PublicationPreparee", "0"
    End If
    modIntegrationUnifie.FixerVariable doc, "RelectureValidee", "1"
    doc.Save
    modIntegrationUnifie.TransmettreSecretariat doc, sortie
    modIntegrationUnifie.FixerVariable doc, "RelectureEnAttente", "0"
    doc.Save
    MsgBox "Courrier transmis au secretariat.", vbInformation, "Cabinet"
End Sub

Public Function EmpreinteCourrier(ByVal doc As Document, ByVal destinataireID As String) As String
    Dim xml As Object, noeud As Object, nodes As Object, attr As Object, fin As Object, partie As Object, id As String
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False: xml.resolveExternals = False
    If Not xml.LoadXML(doc.Content.WordOpenXML) Then Err.Raise vbObjectError + 1167, , "Empreinte du document impossible."
    ' Les proprietes de sauvegarde et les variables de reprise ne font pas
    ' partie de la revision lisible. Corps, images, styles et mise en page oui.
    Set nodes = xml.SelectNodes("//*[local-name()='part' and (@*[local-name()='name']='/docProps/core.xml' or @*[local-name()='name']='/docProps/app.xml')] | //*[local-name()='docVars'] | //*[local-name()='rsids']")
    For Each noeud In nodes: noeud.ParentNode.RemoveChild noeud: Next noeud
    Set nodes = xml.SelectNodes("//*[@*[starts-with(local-name(),'rsid')]]")
    For Each noeud In nodes
        For Each attr In noeud.SelectNodes("@*[starts-with(local-name(),'rsid')]"): noeud.RemoveAttributeNode attr: Next attr
    Next noeud
    ' Exclure uniquement les reperes internes de navigation et l identifiant
    ' Word regenere a la sauvegarde ; conserver le texte et sa mise en page.
    Set nodes = xml.SelectNodes("//*[local-name()='bookmarkStart' and @*[local-name()='name']='_GoBack']")
    For Each noeud In nodes
        id = CStr(noeud.SelectSingleNode("@*[local-name()='id']").Text)
        Set partie = noeud.ParentNode
        Do While partie.nodeName <> "pkg:part" And Not partie.ParentNode Is Nothing
            Set partie = partie.ParentNode
        Loop
        For Each fin In partie.SelectNodes(".//*[local-name()='bookmarkEnd' and @*[local-name()='id']='" & id & "']")
            fin.ParentNode.RemoveChild fin
        Next fin
        noeud.ParentNode.RemoveChild noeud
    Next noeud
    Set nodes = xml.SelectNodes("//*[local-name()='settings']/*[local-name()='docId']")
    For Each noeud In nodes: noeud.ParentNode.RemoveChild noeud: Next noeud
    EmpreinteCourrier = modServiceNas.SHA256(destinataireID & "|" & xml.XML)
End Function
