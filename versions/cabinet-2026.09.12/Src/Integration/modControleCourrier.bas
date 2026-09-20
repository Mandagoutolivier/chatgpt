Attribute VB_Name = "modControleCourrier"
Option Explicit

Public Sub PreparerRelecture(ByVal doc As Document, ByVal source As String, ByVal resultat As String)
    Dim p As Object, r As Object, difference As Object, message As String
    Dim cor As Object, sortie As String, revision As String, valeur As Variant
    Set p = modServiceNas.Parametres(): p("source") = source: p("resultat") = resultat
    Set r = modServiceNas.Appeler("clinical.compare", p)
    For Each difference In r("differences")
        message = message & vbCrLf & "- " & CStr(difference("controle"))
        For Each valeur In difference("retires"): message = message & " [retire: " & CStr(valeur) & "]": Next valeur
        For Each valeur In difference("ajoutes"): message = message & " [ajoute: " & CStr(valeur) & "]": Next valeur
    Next difference
    Set cor = AssurerDestinataire(doc)
    If cor Is Nothing Then Exit Sub
    modCourrier.RemplirSignet doc, "DESTINATAIRE", CStr(cor("BlocDestinataire"))
    modCourrier.MettreEnFormeDestinataire doc
    revision = EmpreinteCourrier(doc, CStr(cor("ID")))
    If Trim$(modIntegrationUnifie.VariableDoc(doc, "PublicationRevision")) <> revision Then
        modIntegrationUnifie.FixerVariable doc, "PublicationID", modFichiers.IdUnique()
        modIntegrationUnifie.FixerVariable doc, "PublicationRevision", revision
        modIntegrationUnifie.FixerVariable doc, "DateValidation", Format$(Now, "dd/mm/yyyy hh:nn:ss")
        modIntegrationUnifie.FixerVariable doc, "PublicationPreparee", "0"
    End If
    modIntegrationUnifie.FixerVariable doc, "ControlesRelecture", message
    modIntegrationUnifie.FixerVariable doc, "RelectureEnAttente", "0"
    modIntegrationUnifie.FixerVariable doc, "RelectureValidee", "1"
    doc.Save
    modIntegrationUnifie.TransmettreSecretariat doc, sortie
    Application.ScreenUpdating = True
    MsgBox "Courrier transmis au secretariat." & _
        IIf(Len(message) > 0, vbCrLf & "Controles automatiques a verifier :" & message, ""), _
        vbInformation, "Cabinet"
End Sub

Public Function AssurerDestinataire(ByVal doc As Document) As Object
    Dim cor As Object, id As String, f As ufListe, recherche As String, p As Object
    id = Trim$(modIntegrationUnifie.VariableDoc(doc, "CorrespondantID"))
    If Len(id) = 0 Then
        recherche = Trim$(InputBox("Nom du destinataire principal dicte (selection d un identifiant stable) :", "Destinataire"))
        If Len(recherche) < 2 Then Exit Function
        Set f = New ufListe
        f.Configurer "Confirmer le destinataire", modServiceNas.LireTable("CORRESPONDANTS", recherche), Array("Nom", "Prenom", "Adresse1", "Ville"), "110 pt;90 pt;180 pt;100 pt"
        f.Show vbModal
        If Not f.Annule Then Set cor = f.Resultat
        Unload f
        If cor Is Nothing Then Exit Function
        id = CStr(cor("ID"))
    End If
    Set p = modServiceNas.Parametres(): p("id") = id
    Set cor = modServiceNas.Appeler("correspondent.resolve", p)
    modIntegrationUnifie.FixerVariable doc, "CorrespondantID", CStr(cor("ID"))
    Set AssurerDestinataire = cor
End Function

Public Sub ValiderEtTransmettre(ByVal doc As Document)
    Dim cor As Object, sortie As String, revision As String
    modIntegrationUnifie.InitialiserPatientProd doc
    Set cor = AssurerDestinataire(doc)
    If cor Is Nothing Then Exit Sub
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
    EmpreinteCourrier = EmpreinteXmlCourrier(doc.Content.WordOpenXML, destinataireID)
End Function

Public Function EmpreinteXmlCourrier(ByVal contenuXml As String, ByVal destinataireID As String) As String
    Dim xml As Object, noeud As Object, nodes As Object, attr As Object, fin As Object, partie As Object, id As String
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False: xml.resolveExternals = False
    If Not xml.LoadXML(contenuXml) Then Err.Raise vbObjectError + 1167, , "Empreinte du document impossible."
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
    ' Word 2016 ajoute a la premiere sauvegarde ses compteurs d identifiants
    ' VML. Ils ne decrivent aucun dessin. Les autres valeurs par defaut et
    ' regles de mise en page restent integralement dans l empreinte.
    xml.setProperty "SelectionNamespaces", _
        "xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage' " & _
        "xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' " & _
        "xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:v='urn:schemas-microsoft-com:vml'"
    Set nodes = xml.SelectNodes("/pkg:package/pkg:part[@pkg:name='/word/settings.xml']/pkg:xmlData/w:settings/w:shapeDefaults")
    For Each partie In nodes
        For Each noeud In partie.SelectNodes("o:shapedefaults")
            noeud.RemoveAttribute "spidmax"
            RetirerNoeudVmlVide noeud
        Next noeud
        For Each noeud In partie.SelectNodes("o:shapelayout/o:idmap")
            noeud.RemoveAttribute "data"
            RetirerNoeudVmlVide noeud
        Next noeud
        For Each noeud In partie.SelectNodes("o:shapelayout")
            RetirerNoeudVmlVide noeud
        Next noeud
        If partie.Attributes.Length = 0 And Not partie.HasChildNodes Then partie.ParentNode.RemoveChild partie
    Next partie
    EmpreinteXmlCourrier = modServiceNas.SHA256(destinataireID & "|" & xml.XML)
End Function

Private Sub RetirerNoeudVmlVide(ByVal noeud As Object)
    Dim attr As Object
    If noeud.HasChildNodes Then Exit Sub
    For Each attr In noeud.Attributes
        If attr.namespaceURI <> "urn:schemas-microsoft-com:vml" Or attr.baseName <> "ext" Or attr.Text <> "edit" Then Exit Sub
    Next attr
    noeud.ParentNode.RemoveChild noeud
End Sub
