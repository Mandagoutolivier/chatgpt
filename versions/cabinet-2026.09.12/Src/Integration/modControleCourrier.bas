Attribute VB_Name = "modControleCourrier"
Option Explicit

' Correction et validation humaine sont distinctes. Aucun envoi dans cette procedure.
Public Sub PreparerRelecture(ByVal doc As Document, ByVal source As String, ByVal resultat As String)
    Dim p As Object, r As Object, difference As Object, message As String
    Dim cor As Object, valeur As Variant
    InvaliderRelecture doc
    Set p = modServiceNas.Parametres(): p("source") = source: p("resultat") = resultat
    Set r = modServiceNas.Appeler("clinical.compare", p)
    For Each difference In r("differences")
        message = message & vbCrLf & "- " & CStr(difference("controle"))
        For Each valeur In difference("retires"): message = message & " [retire: " & CStr(valeur) & "]": Next valeur
        For Each valeur In difference("ajoutes"): message = message & " [ajoute: " & CStr(valeur) & "]": Next valeur
    Next difference
    Set cor = AssurerDestinataire(doc)
    If cor Is Nothing Then Exit Sub
    ' Le bloc definitif est visible AVANT la relecture, jamais remplace apres elle.
    modCourrier.RemplirSignet doc, "DESTINATAIRE", CStr(cor("BlocDestinataire"))
    modCourrier.MettreEnFormeDestinataire doc
    modIntegrationUnifie.FixerVariable doc, "RelectureDestinataireID", CStr(cor("ID"))
    modIntegrationUnifie.FixerVariable doc, "RelectureAdresseSource", modServiceNas.SHA256(CStr(cor("BlocDestinataire")))
    modIntegrationUnifie.FixerVariable doc, "RelectureAdresseDocument", modServiceNas.SHA256(doc.Bookmarks("DESTINATAIRE").Range.Text)
    modIntegrationUnifie.FixerVariable doc, "ControlesRelecture", message
    modIntegrationUnifie.FixerVariable doc, "RelectureEnAttente", "1"
    doc.Save
    Application.ScreenUpdating = True
    MsgBox "Le courrier et ses annexes sont prets pour une seule relecture." & vbCrLf & _
        IIf(Len(message) > 0, "Differences detectees :" & message & vbCrLf, "") & _
        "Verifiez identite, destinataires, negations, doses et examens. Apres relecture, D transmet directement, sans seconde boite de confirmation.", vbInformation, "Relecture du courrier"
End Sub

Public Sub InvaliderRelecture(ByVal doc As Document)
    modIntegrationUnifie.FixerVariable doc, "RelectureEnAttente", "0"
    modIntegrationUnifie.FixerVariable doc, "RelectureValidee", "0"
    modIntegrationUnifie.FixerVariable doc, "RelectureDestinataireID", ""
    modIntegrationUnifie.FixerVariable doc, "RelectureAdresseSource", ""
    modIntegrationUnifie.FixerVariable doc, "RelectureAdresseDocument", ""
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

Public Function DestinataireRelectureConforme(ByVal doc As Document, ByVal cor As Object) As Boolean
    Dim id As String
    If cor Is Nothing Then Exit Function
    If Not doc.Bookmarks.Exists("DESTINATAIRE") Then Exit Function
    id = Trim$(modIntegrationUnifie.VariableDoc(doc, "CorrespondantID"))
    If Len(id) = 0 Then Exit Function
    If id <> Trim$(modIntegrationUnifie.VariableDoc(doc, "RelectureDestinataireID")) Then Exit Function
    If id <> CStr(cor("ID")) Then Exit Function
    If modServiceNas.SHA256(CStr(cor("BlocDestinataire"))) <> Trim$(modIntegrationUnifie.VariableDoc(doc, "RelectureAdresseSource")) Then Exit Function
    If modServiceNas.SHA256(doc.Bookmarks("DESTINATAIRE").Range.Text) <> Trim$(modIntegrationUnifie.VariableDoc(doc, "RelectureAdresseDocument")) Then Exit Function
    DestinataireRelectureConforme = True
End Function

' Le geste explicite D apres relecture est la validation unique ; pas de Oui/Non redondant.
Public Sub ValiderEtTransmettre(ByVal doc As Document)
    On Error GoTo Echec
    Dim cor As Object, p As Object, sortie As String, revision As String, id As String
    Dim numero As Long, description As String
    If Trim$(modIntegrationUnifie.VariableDoc(doc, "RelectureEnAttente")) <> "1" Then Err.Raise vbObjectError + 1168, , "Preparez puis relisez le courrier avant transmission."
    modIntegrationUnifie.InitialiserPatientProd doc
    id = Trim$(modIntegrationUnifie.VariableDoc(doc, "CorrespondantID"))
    If Len(id) = 0 Then Err.Raise vbObjectError + 1168, , "Destinataire absent : reprendre la correction."
    Set p = modServiceNas.Parametres(): p("id") = id
    Set cor = modServiceNas.Appeler("correspondent.resolve", p)
    If Not DestinataireRelectureConforme(doc, cor) Then
        InvaliderRelecture doc
        Err.Raise vbObjectError + 1168, , "Le destinataire ou son adresse a change. Reprenez la correction et verifiez le nouveau destinataire. Aucun envoi."
    End If
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
    modIntegrationUnifie.FixerVariable doc, "RelectureValidee", "0"
    doc.Save
    MsgBox "Courrier transmis au secretariat.", vbInformation, "Cabinet"
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    modIntegrationUnifie.FixerVariable doc, "RelectureValidee", "0"
    doc.Save
    On Error GoTo 0
    Err.Raise numero, "ValiderEtTransmettre", description
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
