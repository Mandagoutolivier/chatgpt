Attribute VB_Name = "modEditionSecretariat"
Option Explicit
' ============================================================================
' U2 - Edition des annexes par le secretariat, SOURCE DE DEVELOPPEMENT.
' Hote : Excel / Cabinet.xlsm. Ne pas importer dans Word ou Normal.dotm.
' Version : 2026.09.20-dev2. Non compilee dans Office, non qualifiee cliniquement.
' Aucune procedure de demarrage automatique ; les entrees recoivent mDrapeau.
'
' Dependances U2 : modServiceNas, modFichiers, modDonneesTransport, modJson,
'                 modIndexAnnexes et formulaire Excel ufListe.
' Serveur attendu : publication.get et publication.revise (en developpement).
' Ne modifie jamais une archive publiee, un acte ou un reglement.
' Les copies de travail et les instantanes d'envoi sont conserves.
' ============================================================================
Private mOccupe As Boolean
Private Const TITRE As String = "Edition secretariat - developpement"

Private Sub VerifierHote()
    If InStr(1, Application.Name, "Excel", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1210, TITRE, "Module reserve au classeur Excel Cabinet.xlsm."
    End If
End Sub

Private Function Texte(ByVal d As Object, ByVal cle As String) As String
    If d Is Nothing Then Err.Raise vbObjectError + 1210, TITRE, "Objet absent."
    If Not d.Exists(cle) Then Err.Raise vbObjectError + 1210, TITRE, "Information absente : " & cle
    If VarType(d(cle)) <> vbString Then Err.Raise vbObjectError + 1210, TITRE, "Texte attendu : " & cle
    Texte = CStr(d(cle))
End Function

Private Function Hexadecimal(ByVal valeur As String, ByVal longueur As Long) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[0-9a-f]{" & CStr(longueur) & "}$"
    Hexadecimal = re.Test(valeur)
End Function

Private Function NouvelID() As String
    ' IdUnique U2 renvoie un GUID avec tirets ; le serveur exige 32 hexadecimaux.
    NouvelID = LCase$(Replace(modFichiers.IdUnique(), "-", ""))
    If Not Hexadecimal(NouvelID, 32) Then Err.Raise vbObjectError + 1210, TITRE, "Identifiant de revision invalide."
End Function

Private Function CheminCache(ByVal sourceID As String) As String
    Dim base As String
    modFichiers.VerifierIdentifiantFichier sourceID
    base = Environ$("LOCALAPPDATA")
    If Len(base) = 0 Then Err.Raise vbObjectError + 1210, TITRE, "LOCALAPPDATA introuvable."
    CheminCache = base & "\CabinetCardio\EditionsSecretariat\" & sourceID & ".json"
End Function

Private Function LireCache(ByVal sourceID As String) As Object
    Dim chemin As String, info As Object
    chemin = CheminCache(sourceID)
    If Not modFichiers.FichierExiste(chemin) Then Exit Function
    Set info = modJson.JsonParse(modFichiers.LireTexteUTF8(chemin))
    If TypeName(info) <> "Dictionary" Then Err.Raise vbObjectError + 1210, TITRE, "Suivi local invalide : conserver le fichier."
    If Texte(info, "SourceID") <> sourceID Then Err.Raise vbObjectError + 1210, TITRE, "Suivi rattache a une autre publication."
    Select Case Texte(info, "Etat")
        Case "edition", "preparee", "terminee", "abandonnee"
        Case Else: Err.Raise vbObjectError + 1210, TITRE, "Etat local inconnu : conserver le fichier."
    End Select
    Set LireCache = info
End Function

Private Sub SauverCache(ByVal sourceID As String, ByVal info As Object)
    Dim chemin As String, tmp As String
    chemin = CheminCache(sourceID)
    If Texte(info, "SourceID") <> sourceID Then Err.Raise vbObjectError + 1210, TITRE, "Publication source incoherente."
    modFichiers.EnsureDossier CreateObject("Scripting.FileSystemObject").GetParentFolderName(chemin)
    tmp = chemin & "." & NouvelID() & ".tmp"
    modFichiers.EcrireTexteUTF8 tmp, modServiceNas.JsonValeur(info)
    modFichiers.RenommerAtomique tmp, chemin, True
End Sub

Private Function LirePublication(ByVal sourceID As String, ByVal exigerCourante As Boolean) As Object
    Dim source As Object
    Set source = modServiceNas.CommandeID("publication.get", sourceID)
    If Texte(source, "PublicationID") <> sourceID Then Err.Raise vbObjectError + 1210, TITRE, "Publication recue differente."
    If Not source.Exists("EditionCourante") Then Err.Raise vbObjectError + 1210, TITRE, "Service NAS incompatible."
    If VarType(source("EditionCourante")) <> vbBoolean Then Err.Raise vbObjectError + 1210, TITRE, "Etat serveur invalide."
    If exigerCourante Then
        If Not CBool(source("EditionCourante")) Then Err.Raise vbObjectError + 1210, TITRE, "Une nouvelle version existe : rechargez la file."
    End If
    If Not Hexadecimal(Texte(source, "sha_docx"), 64) Then Err.Raise vbObjectError + 1210, TITRE, "Empreinte source invalide."
    If Len(Texte(source, "PatientID")) = 0 Or Len(Texte(source, "ConsultationID")) = 0 Then
        Err.Raise vbObjectError + 1210, TITRE, "Rattachement patient/consultation absent."
    End If
    Set LirePublication = source
End Function

Private Function RacineEdition(ByVal source As Object) As String
    Dim p As String, morceaux As Variant, i As Long, suffixe As String
    p = Texte(source, "DossierEdition")
    suffixe = "\Patients\_EditionsSecretariat"
    If Left$(p, 2) <> "\\" Or Len(p) <= Len(suffixe) Then Err.Raise vbObjectError + 1210, TITRE, "Partage d edition invalide."
    If StrComp(Right$(p, Len(suffixe)), suffixe, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 1210, TITRE, "Dossier d edition non reserve."
    If InStr(p, "/") > 0 Or InStr(p, ":") > 0 Or InStr(p, "?") > 0 Or InStr(p, "*") > 0 Then
        Err.Raise vbObjectError + 1210, TITRE, "Chemin d edition interdit."
    End If
    morceaux = Split(Mid$(p, 3), "\")
    If UBound(morceaux) < 3 Then Err.Raise vbObjectError + 1210, TITRE, "Partage d edition incomplet."
    For i = 0 To UBound(morceaux)
        If Len(morceaux(i)) = 0 Or morceaux(i) = "." Or morceaux(i) = ".." Then
            Err.Raise vbObjectError + 1210, TITRE, "Chemin relatif interdit."
        End If
    Next i
    RacineEdition = p
End Function

Private Sub VerifierCache(ByVal info As Object, ByVal source As Object)
    Dim dossier As String
    If Not Hexadecimal(Texte(info, "ID"), 32) Then Err.Raise vbObjectError + 1210, TITRE, "Identifiant du suivi invalide."
    If Texte(info, "SourceID") <> Texte(source, "PublicationID") Or Texte(info, "SourceSHA") <> Texte(source, "sha_docx") Then
        Err.Raise vbObjectError + 1210, TITRE, "Source du suivi differente : conserver la copie."
    End If
    dossier = RacineEdition(source) & "\" & Texte(info, "ID")
    If StrComp(Texte(info, "Dossier"), dossier, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 1210, TITRE, "Partage modifie : conserver la copie."
    If StrComp(Texte(info, "Chemin"), dossier & "\courrier.docx", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 1210, TITRE, "Copie de travail hors du dossier reserve."
End Sub

Public Function EditionEnCours(ByVal publication As Object) As Boolean
    Dim info As Object
    VerifierHote
    Set info = LireCache(Texte(publication, "PublicationID"))
    If info Is Nothing Then Exit Function
    EditionEnCours = (Texte(info, "Etat") = "edition" Or Texte(info, "Etat") = "preparee")
End Function

Private Function ObtenirWord() As Object
    On Error Resume Next
    Set ObtenirWord = GetObject(, "Word.Application")
    On Error GoTo 0
    If ObtenirWord Is Nothing Then Set ObtenirWord = CreateObject("Word.Application")
    ObtenirWord.Visible = True
End Function

Private Function DocumentOuvert(ByVal word As Object, ByVal chemin As String) As Object
    Dim doc As Object
    For Each doc In word.Documents
        If StrComp(CStr(doc.FullName), chemin, vbTextCompare) = 0 Then
            Set DocumentOuvert = doc
            Exit Function
        End If
    Next doc
End Function

Private Function OuvrirSansMacros(ByVal word As Object, ByVal chemin As String, ByVal lectureSeule As Boolean) As Object
    Dim securite As Long, liens As Boolean, numero As Long, description As String
    securite = word.AutomationSecurity
    liens = word.Options.UpdateLinksAtOpen
    On Error GoTo Echec
    word.AutomationSecurity = 3
    word.Options.UpdateLinksAtOpen = False
    Set OuvrirSansMacros = word.Documents.Open(chemin, False, lectureSeule, False)
Sortie:
    ' Restaurer les deux options meme si l'ouverture a echoue.
    On Error Resume Next
    Err.Clear
    word.AutomationSecurity = securite
    If Err.Number <> 0 And numero = 0 Then
        numero = Err.Number: description = "Option AutomationSecurity non restauree."
    End If
    Err.Clear
    word.Options.UpdateLinksAtOpen = liens
    If Err.Number <> 0 And numero = 0 Then
        numero = Err.Number: description = "Option UpdateLinksAtOpen non restauree."
    End If
    On Error GoTo 0
    If numero <> 0 Then Err.Raise numero, "OuvrirSansMacros", description
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    Resume Sortie
End Function

Private Sub VerifierDocument(ByVal doc As Object, ByVal info As Object, ByVal source As Object)
    If doc Is Nothing Then Err.Raise vbObjectError + 1211, TITRE, "Rouvrez la copie Word d edition."
    If doc.ReadOnly Then Err.Raise vbObjectError + 1211, TITRE, "Copie en lecture seule ou deja ouverte ailleurs."
    If StrComp(CStr(doc.FullName), Texte(info, "Chemin"), vbTextCompare) <> 0 Then Err.Raise vbObjectError + 1211, TITRE, "Mauvais document d edition."
    If modIndexAnnexes.LireMeta(doc, "PatientID") <> Texte(source, "PatientID") Then Err.Raise vbObjectError + 1211, TITRE, "Patient de la copie different."
    If modIndexAnnexes.LireMeta(doc, "ConsultationID") <> Texte(source, "ConsultationID") Then Err.Raise vbObjectError + 1211, TITRE, "Consultation de la copie differente."
    If modIndexAnnexes.ListerAnnexes(doc).Count = 0 Then Err.Raise vbObjectError + 1211, TITRE, "Aucune annexe munie des nouveaux reperes d edition."
End Sub

Public Sub OuvrirEditionAnnexes(ByVal publication As Object)
    Dim source As Object, info As Object, word As Object, doc As Object
    Dim id As String, copie As String, dossier As String, description As String
    If mOccupe Then Exit Sub
    mOccupe = True
    On Error GoTo Echec
    VerifierHote
    id = Texte(publication, "PublicationID")
    Set source = LirePublication(id, True)
    Set info = LireCache(id)
    If Not info Is Nothing Then
        If Texte(info, "Etat") = "abandonnee" Or Texte(info, "Etat") = "terminee" Then Set info = Nothing
    End If
    If info Is Nothing Then
        Set info = modServiceNas.Parametres()
        info("ID") = NouvelID(): info("SourceID") = id
        info("SourceSHA") = Texte(source, "sha_docx")
        dossier = RacineEdition(source) & "\" & Texte(info, "ID")
        If modFichiers.DossierExiste(dossier) Then Err.Raise vbObjectError + 1210, TITRE, "Dossier de travail deja existant."
        modFichiers.EnsureDossier dossier
        copie = dossier & "\courrier.docx"
        CreateObject("Scripting.FileSystemObject").CopyFile Texte(source, "CheminDocx"), copie, False
        If modDonneesTransport.EmpreinteFichierSHA256(copie) <> Texte(source, "sha_docx") Then
            Err.Raise vbObjectError + 1210, TITRE, "Copie de la publication non conforme."
        End If
        info("Chemin") = copie: info("Dossier") = dossier: info("Etat") = "edition"
        SauverCache id, info
    Else
        VerifierCache info, source
        If Texte(info, "Etat") = "preparee" Then Err.Raise vbObjectError + 1210, TITRE, "Envoi en attente : reprendre Enregistrer la version, sans modifier son instantane."
    End If
    Set word = ObtenirWord()
    Set doc = DocumentOuvert(word, Texte(info, "Chemin"))
    If doc Is Nothing Then Set doc = OuvrirSansMacros(word, Texte(info, "Chemin"), False)
    VerifierDocument doc, info, source
    modIndexAnnexes.FixerMeta doc, "EditionSecretariatID", Texte(info, "ID")
    modIndexAnnexes.FixerMeta doc, "RelectureValidee", "0"
    modIndexAnnexes.FixerMeta doc, "RelectureEnAttente", "0"
    doc.Save
    doc.Activate
Sortie:
    mOccupe = False
    Exit Sub
Echec:
    description = Err.Description
    MsgBox "Ouverture interrompue : " & description & vbCrLf & "Les copies eventuellement creees sont conservees.", vbExclamation, TITRE
    Resume Sortie
End Sub

Public Sub ChoisirDestinataireAnnexe(ByVal publication As Object)
    Dim info As Object, source As Object, word As Object, doc As Object
    Dim f As ufListe, choisi As Object, cor As Object
    Dim recherche As String, numero As String, description As String, id As String
    If mOccupe Then Exit Sub
    mOccupe = True
    On Error GoTo Echec
    VerifierHote
    id = Texte(publication, "PublicationID")
    Set source = LirePublication(id, True)
    Set info = LireCache(id)
    If info Is Nothing Then Err.Raise vbObjectError + 1211, TITRE, "Ouvrez d abord Modifier les annexes dans Word."
    VerifierCache info, source
    If Texte(info, "Etat") <> "edition" Then Err.Raise vbObjectError + 1211, TITRE, "Edition non disponible pour modification."
    Set word = ObtenirWord(): Set doc = DocumentOuvert(word, Texte(info, "Chemin"))
    VerifierDocument doc, info, source
    Set f = New ufListe
    f.Configurer "Annexe a modifier", modIndexAnnexes.ListerAnnexes(doc), Array("Annexe", "Destinataire"), "70 pt;310 pt"
    f.Show vbModal
    If Not f.Annule Then numero = CStr(f.Resultat("ID"))
    Unload f: Set f = Nothing
    If Len(numero) = 0 Then GoTo Sortie
    recherche = Trim$(InputBox("Nom du nouveau destinataire de l annexe :", "Destinataire annexe"))
    If Len(recherche) < 2 Then GoTo Sortie
    Set f = New ufListe
    f.Configurer "Choisir le destinataire", modServiceNas.LireTable("CORRESPONDANTS", recherche), Array("Nom", "Prenom", "Ville"), "150 pt;100 pt;120 pt"
    f.Show vbModal
    If Not f.Annule Then Set choisi = f.Resultat
    Unload f: Set f = Nothing
    If choisi Is Nothing Then GoTo Sortie
    Set cor = modServiceNas.CommandeID("correspondent.resolve", CStr(choisi("ID")))
    modIndexAnnexes.RemplacerDestinataireAnnexe doc, numero, cor
    doc.Save: doc.Activate
Sortie:
    mOccupe = False
    Exit Sub
Echec:
    description = Err.Description
    On Error Resume Next
    If Not f Is Nothing Then Unload f
    On Error GoTo 0
    MsgBox description, vbExclamation, TITRE
    GoTo Sortie
End Sub

Private Sub VerifierRequetePreparee(ByVal p As Object, ByVal info As Object, ByVal source As Object)
    Dim dossier As String
    If Not Hexadecimal(Texte(p, "revision_id"), 32) Then Err.Raise vbObjectError + 1212, TITRE, "Identifiant d envoi invalide."
    If Texte(p, "source_id") <> Texte(info, "SourceID") Or Texte(p, "source_sha") <> Texte(info, "SourceSHA") Then
        Err.Raise vbObjectError + 1212, TITRE, "Source de l envoi incoherente."
    End If
    If Texte(p, "patient_id") <> Texte(source, "PatientID") Then Err.Raise vbObjectError + 1212, TITRE, "Patient de l envoi incoherent."
    dossier = RacineEdition(source) & "\" & Texte(p, "revision_id")
    If StrComp(Texte(p, "docx"), dossier & "\revision.docx", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 1212, TITRE, "DOCX d envoi hors dossier reserve."
    If StrComp(Texte(p, "pdf"), dossier & "\revision.pdf", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 1212, TITRE, "PDF d envoi hors dossier reserve."
    If Not Hexadecimal(Texte(p, "sha_docx"), 64) Or Not Hexadecimal(Texte(p, "sha_pdf"), 64) Then
        Err.Raise vbObjectError + 1212, TITRE, "Empreintes d envoi invalides."
    End If
    ' Aucun recalcul pour remplacer silencieusement les empreintes deja envoyees.
    If modDonneesTransport.EmpreinteFichierSHA256(Texte(p, "docx")) <> Texte(p, "sha_docx") Then Err.Raise vbObjectError + 1212, TITRE, "Instantane DOCX modifie : reprise automatique refusee."
    If modDonneesTransport.EmpreinteFichierSHA256(Texte(p, "pdf")) <> Texte(p, "sha_pdf") Then Err.Raise vbObjectError + 1212, TITRE, "Instantane PDF modifie : reprise automatique refusee."
End Sub

Public Sub EnregistrerEditionAnnexes(ByVal publication As Object)
    Dim info As Object, source As Object, word As Object, doc As Object, snapshot As Object
    Dim p As Object, resultat As Object, cle As Variant
    Dim id As String, revisionID As String, dossier As String, description As String
    Dim accepte As Boolean
    If mOccupe Then Exit Sub
    mOccupe = True
    On Error GoTo Echec
    VerifierHote
    id = Texte(publication, "PublicationID")
    Set info = LireCache(id)
    If info Is Nothing Then Err.Raise vbObjectError + 1212, TITRE, "Aucune edition a enregistrer."
    ' Une reponse perdue peut avoir deja remplace la source : autoriser son suivi.
    Set source = LirePublication(id, False)
    VerifierCache info, source
    If Texte(info, "Etat") = "preparee" Then
        If Not info.Exists("Requete") Then Err.Raise vbObjectError + 1212, TITRE, "Requete de reprise absente."
        If TypeName(info("Requete")) <> "Dictionary" Then Err.Raise vbObjectError + 1212, TITRE, "Requete de reprise invalide."
        Set p = info("Requete")
    Else
        If Texte(info, "Etat") <> "edition" Then Err.Raise vbObjectError + 1212, TITRE, "Edition deja terminee ou abandonnee."
        If Not CBool(source("EditionCourante")) Then Err.Raise vbObjectError + 1212, TITRE, "Une nouvelle version existe. Votre copie reste conservee."
        Set word = ObtenirWord(): Set doc = DocumentOuvert(word, Texte(info, "Chemin"))
        VerifierDocument doc, info, source
        If modIndexAnnexes.LireMeta(doc, "EditionSecretariatID") <> Texte(info, "ID") Then Err.Raise vbObjectError + 1212, TITRE, "Identifiant de la copie different."
        If doc.Revisions.Count > 0 Then Err.Raise vbObjectError + 1212, TITRE, "Resolvez le suivi des modifications dans Word avant l enregistrement."
        doc.Save
        ' Instantane dans un NOUVEAU dossier : aucun instantane precedent ecrase.
        revisionID = NouvelID()
        dossier = RacineEdition(source) & "\" & revisionID
        If modFichiers.DossierExiste(dossier) Then Err.Raise vbObjectError + 1212, TITRE, "Dossier d envoi deja existant."
        modFichiers.EnsureDossier dossier
        CreateObject("Scripting.FileSystemObject").CopyFile Texte(info, "Chemin"), dossier & "\revision.docx", False
        Set snapshot = OuvrirSansMacros(word, dossier & "\revision.docx", True)
        snapshot.ExportAsFixedFormat dossier & "\revision.pdf", 17
        snapshot.Close 0: Set snapshot = Nothing
        Set p = modServiceNas.Parametres()
        p("source_id") = id: p("source_sha") = Texte(info, "SourceSHA")
        p("revision_id") = revisionID: p("patient_id") = Texte(source, "PatientID")
        p("docx") = dossier & "\revision.docx": p("pdf") = dossier & "\revision.pdf"
        p("sha_docx") = modDonneesTransport.EmpreinteFichierSHA256(Texte(p, "docx"))
        p("sha_pdf") = modDonneesTransport.EmpreinteFichierSHA256(Texte(p, "pdf"))
        Set info("Requete") = p: info("Etat") = "preparee"
        SauverCache id, info
    End If
    VerifierRequetePreparee p, info, source
    Set resultat = modServiceNas.Appeler("publication.revise", p)
    If Texte(resultat, "PublicationID") <> Texte(p, "revision_id") Then Err.Raise vbObjectError + 1212, TITRE, "Identifiant de reponse different : faire verifier le suivi."
    If Texte(resultat, "PatientID") <> Texte(source, "PatientID") Then Err.Raise vbObjectError + 1212, TITRE, "Patient de reponse different : faire verifier le suivi."
    If Texte(resultat, "ConsultationID") <> Texte(source, "ConsultationID") Then Err.Raise vbObjectError + 1212, TITRE, "Consultation de reponse differente : faire verifier le suivi."
    accepte = True
    info("Etat") = "terminee": info("PublicationID") = Texte(resultat, "PublicationID")
    SauverCache id, info
    For Each cle In resultat.Keys
        If publication.Exists(CStr(cle)) Then publication.Remove CStr(cle)
        If IsObject(resultat(cle)) Then
            Set publication(CStr(cle)) = resultat(cle)
        Else
            publication(CStr(cle)) = resultat(cle)
        End If
    Next cle
    publication("_Chemin") = Texte(resultat, "PublicationID")
    publication("ID") = Texte(resultat, "PublicationID")
    MsgBox "Instantane corrige enregistre. L original medical est conserve." & vbCrLf & _
           "Aucune facturation ni impression lancee. Les frappes posterieures a l instantane ne sont pas incluses.", vbInformation, TITRE
Sortie:
    mOccupe = False
    Exit Sub
Echec:
    description = Err.Description
    On Error Resume Next
    If Not snapshot Is Nothing Then snapshot.Close 0
    On Error GoTo 0
    If accepte Then
        MsgBox "Le serveur a accepte la revision mais la mise a jour locale a echoue : " & description & vbCrLf & _
               "Rechargez les courriers avant de continuer. Ne recreez pas l envoi.", vbExclamation, TITRE
    Else
        MsgBox "Enregistrement non confirme : " & description & vbCrLf & _
               "Les copies sont conservees. Une requete deja preparee doit etre reprise sans nouveau contenu.", vbExclamation, TITRE
    End If
    GoTo Sortie
End Sub

Public Sub AbandonnerEditionAnnexes(ByVal publication As Object)
    Dim info As Object, id As String, description As String
    If mOccupe Then Exit Sub
    mOccupe = True
    On Error GoTo Echec
    VerifierHote
    id = Texte(publication, "PublicationID")
    Set info = LireCache(id)
    If info Is Nothing Then GoTo Sortie
    If Texte(info, "Etat") = "preparee" Then Err.Raise vbObjectError + 1213, TITRE, "Envoi en attente : reprendre son enregistrement ou faire verifier son resultat avant abandon."
    If Texte(info, "Etat") = "terminee" Or Texte(info, "Etat") = "abandonnee" Then GoTo Sortie
    If MsgBox("Abandonner cette edition ? La copie reste conservee ; l original n est pas modifie.", vbYesNo + vbQuestion, TITRE) <> vbYes Then GoTo Sortie
    info("Etat") = "abandonnee"
    SauverCache id, info
Sortie:
    mOccupe = False
    Exit Sub
Echec:
    description = Err.Description
    MsgBox description, vbExclamation, TITRE
    Resume Sortie
End Sub
