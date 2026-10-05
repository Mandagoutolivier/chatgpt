Attribute VB_Name = "modRecetteC4"
Option Explicit

' Qualification Word locale, sans NAS ni API. A importer seulement dans une
' COPIE instrumentee du modele de recette, jamais dans le modele clinique.
Private Const ATTENDUS_C4 As Long = 70
Private mNombre As Long

Private Sub ExigerC4(ByVal condition As Boolean, ByVal libelle As String)
    If Not condition Then Err.Raise vbObjectError + 1494, "Recette C4", libelle
    mNombre = mNombre + 1
End Sub

Private Function BlocDragon() As String
    BlocDragon = "Docteur NOM Prenom" & vbCr & _
                 "Service FICTIF C4" & vbCr & _
                 "12 rue des Essais" & vbCr & _
                 "75000 VILLE FICTIVE"
End Function

Private Function CorrespondantFictif(ByVal ident As String, ByVal adresse As String, _
                                      Optional ByVal actif As String = "1") As Object
    Dim cor As Object
    Set cor = CreateObject("Scripting.Dictionary")
    cor("ID") = ident
    cor("Actif") = actif
    cor("AValider") = "0"
    cor("Nom") = "NOM"
    cor("Prenom") = "Prenom"
    cor("BlocDestinataire") = Replace(adresse, vbCr, vbLf)
    Set CorrespondantFictif = cor
End Function

Private Function RefuseIdentification(ByVal doc As Document, ByVal candidats As Collection, _
                                      ByVal aliases As Object) As Boolean
    Dim ident As String, numero As Long
    On Error Resume Next
    Err.Clear
    ident = modCourrier.IdDestinataireDicteParmi(doc, candidats, aliases)
    numero = Err.Number: Err.Clear
    On Error GoTo 0
    RefuseIdentification = (numero <> 0 And Len(ident) = 0)
End Function

Private Function RefuseZone(ByVal doc As Document) As Boolean
    Dim zone As Range, numero As Long
    On Error Resume Next
    Err.Clear
    Set zone = modCourrier.ZoneDestinataireDragon(doc)
    numero = Err.Number: Err.Clear
    On Error GoTo 0
    RefuseZone = (numero <> 0 And zone Is Nothing)
End Function

Private Function RefuseCorpsChevauchant(ByVal doc As Document, ByVal corps As Range) As Boolean
    Dim numero As Long
    On Error Resume Next
    Err.Clear
    modCourrier.ExigerCorpsApresDestinataire doc, corps
    numero = Err.Number: Err.Clear
    On Error GoTo 0
    RefuseCorpsChevauchant = (numero <> 0)
End Function

Private Function AliasDragon(ByVal ident As String, ByVal adresseDragon As String, _
                             ByVal empreinte As String) As Object
    Dim entree As Object
    Set entree = CreateObject("Scripting.Dictionary")
    entree("id") = ident
    entree("source_sha256") = empreinte
    entree("bloc") = Replace(adresseDragon, vbCr, vbLf)
    Set AliasDragon = entree
End Function

Private Sub MarquerParagraphe(ByVal doc As Document, ByVal index As Long, ByVal nom As String)
    Dim zone As Range
    Set zone = doc.Paragraphs(index).Range.Duplicate
    zone.MoveEnd wdCharacter, -1
    doc.Bookmarks.Add nom, zone
End Sub

Private Sub CreerFixture(ByVal chemin As String, ByVal moderne As Boolean)
    Dim doc As Document, numero As Long, description As String
    On Error GoTo Echec
    Set doc = Documents.Add(Visible:=False)
    If moderne Then
        doc.Content.Text = "EXPEDITEUR_FIX_C4" & vbCr & " " & vbCr & _
                           "DATE_FIX_C4" & vbCr & "CONCERNE_FIX_C4" & vbCr & _
                           " " & vbCr & "CORPS_FIX_C4" & vbCr & "SIGNATURE_FIX_C4" & vbCr
        MarquerParagraphe doc, 1, "EXPEDITEUR"
        MarquerParagraphe doc, 2, "DESTINATAIRE"
        MarquerParagraphe doc, 3, "DATELIEU"
        MarquerParagraphe doc, 4, "CONCERNE"
        MarquerParagraphe doc, 5, "APPEL"
        MarquerParagraphe doc, 6, "CORPS"
        MarquerParagraphe doc, 7, "SIGNATURE"
    Else
        doc.Content.Text = "ENTETE_FIX_C4" & vbCr & " " & vbCr & _
                           "DATE_FIX_C4" & vbCr & " " & vbCr & _
                           " " & vbCr & "SIGNATURE_FIX_C4" & vbCr
        MarquerParagraphe doc, 2, "CORRESPONDANT"
        MarquerParagraphe doc, 4, "FORMULE_APPEL"
    End If
    doc.SaveAs2 FileName:=chemin, FileFormat:=wdFormatXMLTemplate
    doc.Close wdDoNotSaveChanges
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "CreerFixture C4", description
End Sub

Private Sub ConfigurerFixture(ByVal racine As String, ByVal nomModele As String)
    modFichiers.EcrireTexteUTF8 racine & "\Config\config.ini", _
        "[COURRIER]" & vbCrLf & "Modele=" & nomModele & vbCrLf & _
        "AppelAuto=0" & vbCrLf & "PolitesseAuto=0" & vbCrLf & _
        "[GENERAL]" & vbCrLf & "Ville=VILLE_DATE_FIX_C4" & vbCrLf & _
        "[MEDECIN]" & vbCrLf & "Titre=Docteur" & vbCrLf & _
        "Prenom=TEST" & vbCrLf & "Nom=C4" & vbCrLf & _
        "Signature=SIGNATURE_FIX_C4"
    modConfig.DefinirRacine racine
End Sub

Private Function MemePlage(ByVal a As Range, ByVal b As Range) As Boolean
    MemePlage = (a.StoryType = b.StoryType And a.Start = b.Start And a.End = b.End)
End Function

Private Sub TesterParcours(ByVal racine As String, ByVal nom As String, ByVal moderne As Boolean)
    Dim doc As Document, zone As Range, candidats As Collection, cor As Object, aliases As Object
    Dim avant As String, blocBrut As String, fichier As String, horsAdresse As Boolean
    Dim controle As ContentControl, doublon As ContentControl, emplacement As Range
    Dim numero As Long, description As String
    On Error GoTo Echec
    ConfigurerFixture racine, nom
    Set doc = modCourrier.CreerCourrierRapide()
    ExigerC4 doc.ContentControls.Count = 1, nom & ": zone Dragon unique creee"
    ExigerC4 doc.Bookmarks.Exists("DATELIEU") = moderne, nom & ": disposition date attendue"

    doc.Activate
    modCourrier.AllerDestinataire
    ExigerC4 Selection.Type = wdSelectionIP, nom & ": curseur dans le destinataire"
    ' Imite un AutoText multiligne avec de vrais retours de paragraphe Word.
    Selection.TypeText Text:=BlocDragon()
    Set zone = modCourrier.ZoneDestinataireDragon(doc)
    blocBrut = modCourrier.BlocDestinataireDicte(doc)
    ExigerC4 InStr(1, blocBrut, "Docteur NOM Prenom", vbBinaryCompare) > 0, nom & ": debut adresse capture"
    ExigerC4 InStr(1, blocBrut, "75000 VILLE FICTIVE", vbBinaryCompare) > 0, nom & ": fin adresse capture"
    horsAdresse = (InStr(1, blocBrut, "FIX_C4", vbBinaryCompare) = 0)
    If moderne Then
        horsAdresse = horsAdresse And zone.Start >= doc.Bookmarks("EXPEDITEUR").Range.End
        horsAdresse = horsAdresse And zone.End <= doc.Bookmarks("DATELIEU").Range.Start
    End If
    ExigerC4 horsAdresse, nom & ": expediteur et date hors de l'adresse"

    ' L'AutoText n'a pas a conserver les anciens signets : la zone fait foi.
    If doc.Bookmarks.Exists("DESTINATAIRE") Then doc.Bookmarks("DESTINATAIRE").Delete
    If doc.Bookmarks.Exists("CORRESPONDANT") Then doc.Bookmarks("CORRESPONDANT").Delete
    Set aliases = New Collection ' Aucun acces APPDATA de ce profil de qualification.
    Set candidats = New Collection
    candidats.Add CorrespondantFictif("C4-UNIQUE", BlocDragon())
    avant = doc.Content.Text
    ExigerC4 modCourrier.IdDestinataireDicteParmi(doc, candidats, aliases) = "C4-UNIQUE" And _
             doc.Content.Text = avant, nom & ": ID unique sans reecrire l'adresse"
    Set candidats = New Collection
    candidats.Add CorrespondantFictif("C4-SAUT-MANUEL", Replace(BlocDragon(), vbCr, Chr$(11)))
    ExigerC4 modCourrier.IdDestinataireDicteParmi(doc, candidats, aliases) = "C4-SAUT-MANUEL", _
             nom & ": sauts manuels et paragraphes equivalents"

    ' Une variante complete validee suffit ; une empreinte de source perimee
    ' ne doit jamais autoriser le meme texte contre la fiche courante.
    Set candidats = New Collection
    Set cor = CorrespondantFictif("C4-ALIAS", Replace(BlocDragon(), _
        "Docteur NOM Prenom", "Docteur Prenom NOM"))
    cor.Remove "Nom": cor.Remove "Prenom"
    candidats.Add cor
    Set aliases = New Collection
    aliases.Add AliasDragon("C4-ALIAS", BlocDragon(), _
                           modServiceNas.SHA256(CStr(cor("BlocDestinataire"))))
    ExigerC4 modCourrier.IdDestinataireDicteParmi(doc, candidats, aliases) = "C4-ALIAS", _
             nom & ": alias complet approuve pour la source courante"
    Set aliases = New Collection
    aliases.Add AliasDragon("C4-ALIAS", BlocDragon(), String$(64, "0"))
    ExigerC4 RefuseIdentification(doc, candidats, aliases), _
             nom & ": alias perime refuse apres changement de source"
    Set aliases = New Collection

    Set candidats = New Collection
    Set cor = CorrespondantFictif("C4-INVERSE", Replace(BlocDragon(), _
        "Docteur NOM Prenom", "Docteur Prenom NOM"))
    candidats.Add cor
    ExigerC4 modCourrier.IdDestinataireDicteParmi(doc, candidats, aliases) = "C4-INVERSE", _
             nom & ": ordre Dragon NOM Prenom reconnu par les champs structures"
    ExigerC4 modCourrier.BlocCorrespondantDragonConforme(Replace(BlocDragon(), _
        "Docteur NOM Prenom", "Monsieur le Docteur NOM Prenom"), cor, aliases), _
             nom & ": titre Monsieur le Docteur reconnu"
    ExigerC4 modCourrier.BlocCorrespondantDragonConforme(Replace(BlocDragon(), _
        "Docteur NOM Prenom", "Madame le Docteur Prenom NOM"), cor, aliases), _
             nom & ": ordre Prenom NOM et titre Madame le Docteur reconnus"
    Set candidats = New Collection
    candidats.Add CorrespondantFictif("C4-INACTIF", BlocDragon(), "0")
    ExigerC4 RefuseIdentification(doc, candidats, aliases), nom & ": inactif refuse"
    Set candidats = New Collection
    Set cor = CorrespondantFictif("C4-A-VALIDER", BlocDragon())
    cor("AValider") = "1": candidats.Add cor
    ExigerC4 RefuseIdentification(doc, candidats, aliases), nom & ": adresse non validee refusee"
    Set candidats = New Collection
    candidats.Add CorrespondantFictif("C4-A", BlocDragon())
    candidats.Add CorrespondantFictif("C4-B", BlocDragon())
    ExigerC4 RefuseIdentification(doc, candidats, aliases), nom & ": deux IDs au meme bloc refuses"

    avant = doc.Content.Text
    modCourrier.ReancrerSignetsDestinataire doc
    Set zone = modCourrier.ZoneDestinataireDragon(doc)
    ExigerC4 doc.Content.Text = avant, nom & ": reancrage sans mutation du texte"
    ExigerC4 UBound(Split(doc.Content.Text, "Docteur NOM Prenom")) = 1, _
             nom & ": adresse Dragon presente une seule fois"
    ExigerC4 doc.Bookmarks.Exists("DESTINATAIRE") And _
             doc.Bookmarks.Exists("CORRESPONDANT") And _
             MemePlage(doc.Bookmarks("DESTINATAIRE").Range, zone) And _
             MemePlage(doc.Bookmarks("CORRESPONDANT").Range, zone), _
             nom & ": les deux alias encadrent le bloc entier"

    ' Partie B locale : l'association RPC est instrumentee ailleurs ; la
    ' navigation vers l'appel puis l'insertion C restent au curseur Word.
    doc.Activate
    modCourrier.AllerAppel
    ExigerC4 Selection.Type = wdSelectionIP And _
             Selection.Start >= doc.Bookmarks("APPEL").Range.Start And _
             Selection.Start <= doc.Bookmarks("APPEL").Range.End, _
             nom & ": B place le curseur dans l'appel"
    Selection.TypeText Text:="Cher Confrere," & vbCr & "Texte clinique avant "
    Selection.TypeText Text:="Monsieur FICTIF, 45 ans, "
    ExigerC4 InStr(1, doc.Content.Text, "Texte clinique avant Monsieur FICTIF, 45 ans, ", vbBinaryCompare) > 0 And _
             modCourrier.BlocDestinataireDicte(doc) = blocBrut, _
             nom & ": insertion au curseur sans ecraser le destinataire"

    fichier = racine & "\" & nom & "-apres-dragon.docx"
    doc.SaveAs2 FileName:=fichier, FileFormat:=wdFormatXMLDocument
    doc.Close wdDoNotSaveChanges: Set doc = Nothing
    Set doc = Documents.Open(FileName:=fichier, ReadOnly:=False, AddToRecentFiles:=False)
    Set zone = modCourrier.ZoneDestinataireDragon(doc)
    ExigerC4 modCourrier.BlocDestinataireDicte(doc) = blocBrut, nom & ": bloc persiste apres reouverture"
    ExigerC4 doc.Bookmarks.Exists("DESTINATAIRE") And _
             doc.Bookmarks.Exists("CORRESPONDANT") And _
             MemePlage(doc.Bookmarks("DESTINATAIRE").Range, zone) And _
             MemePlage(doc.Bookmarks("CORRESPONDANT").Range, zone), _
             nom & ": alias persistent apres reouverture"

    Set controle = doc.ContentControls(1)
    Set emplacement = doc.Paragraphs(1).Range.Duplicate
    emplacement.MoveEnd wdCharacter, -1
    Set doublon = doc.ContentControls.Add(wdContentControlRichText, emplacement)
    doublon.Tag = controle.Tag
    ExigerC4 RefuseZone(doc), nom & ": deux controles marques refuses"
    doublon.LockContentControl = False
    doublon.Delete False
    controle.LockContentControl = False
    controle.Delete False
    ExigerC4 RefuseZone(doc), nom & ": controle disparu refuse sans repli"
    doc.Close wdDoNotSaveChanges: Set doc = Nothing
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "TesterParcours C4 " & nom, description
End Sub

Private Sub TesterPatientMemorise()
    Dim doc As Document, pat As Object, annee As String
    Dim numero As Long, description As String
    On Error GoTo Echec
    Set doc = Documents.Add(Visible:=False)
    annee = CStr(Year(Date) - 40)
    modIntegrationUnifie.FixerVariable doc, "PatientID", "P-FICTIF-C4"
    modIntegrationUnifie.FixerVariable doc, "ConsultationID", "R-FICTIF-C4"
    modIntegrationUnifie.FixerVariable doc, "Patient_Nom", "NOMFICTIF"
    modIntegrationUnifie.FixerVariable doc, "Patient_Prenom", "Alice"
    modIntegrationUnifie.FixerVariable doc, "Patient_DDN", "01/01/" & annee
    modIntegrationUnifie.FixerVariable doc, "Patient_Sexe", "F"
    Set pat = modIntegrationUnifie.PatientMemorise(doc)
    ExigerC4 CStr(pat("Nom")) = "NOMFICTIF" And CStr(pat("Prenom")) = "Alice", _
             "C relit l identite memorisee sans acces reseau"
    ExigerC4 modCourrier.TexteIdentitePatient(pat) = "NOMFICTIF Alice, 40 ans, ", _
             "C forme nom prenom et age depuis le brouillon"
    ExigerC4 modCourrier.PrefixeIdentitePatient("s") = " ", "C ajoute un espace apres un mot"
    ExigerC4 modCourrier.PrefixeIdentitePatient(" ") = "", "C ne double pas un espace"
    ExigerC4 modCourrier.PrefixeIdentitePatient(vbCr) = "", "C au debut de ligne sans espace"
    ExigerC4 modCourrier.PrefixeIdentitePatient("'") = "", "C apres apostrophe sans espace"
    ExigerC4 modCourrier.PrefixeIdentitePatient(":") = " ", "C apres ponctuation separe le nom"
    ExigerC4 "Je revois" & modCourrier.PrefixeIdentitePatient("s") & _
             modCourrier.TexteIdentitePatient(pat) = "Je revois NOMFICTIF Alice, 40 ans, ", _
             "C ne colle pas l identite au mot precedent"
    ExigerC4 modCycleCourrier.AnonymiserIdentiteC4("Je revois NOMFICTIF Alice, 40 ans.", _
             "Madame", "NOMFICTIF", "Alice") = "Je revois [[PATIENT]], 40 ans.", _
             "C sans civilite est anonymise avant l API"
    ExigerC4 modCycleCourrier.AnonymiserIdentiteC4("Je revois Madame NOMFICTIF Alice.", _
             "Madame", "NOMFICTIF", "Alice") = "Je revois Madame [[PATIENT]].", _
             "civilite dictee avant C reste compatible"
    doc.Close wdDoNotSaveChanges: Set doc = Nothing
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "TesterPatientMemorise C4", description
End Sub

Private Sub TesterAppelDansDestinataire()
    Dim doc As Document, corps As Range, premier As Range, zone As Range
    Dim numero As Long, description As String
    On Error GoTo Echec
    Set doc = Documents.Add(Visible:=False)
    doc.Content.Text = "ENTETE_FIX_C4" & vbCr & _
        "Madame, Monsieur," & vbCr & "SERVICE DESTINATAIRE FICTIF" & vbCr & _
        "Beaumont, le 3 octobre 2026" & vbCr & "Cher Confrere," & vbCr & _
        "Monsieur NOM Prenom, 45 ans, presente un symptome fictif." & vbCr & _
        "Bien cordialement." & vbCr & "Docteur olivier mandagout" & vbCr
    doc.Bookmarks.Add "DESTINATAIRE", doc.Range(doc.Paragraphs(2).Range.Start, _
                                                    doc.Paragraphs(3).Range.End - 1)
    doc.Bookmarks.Add "CORRESPONDANT", doc.Bookmarks("DESTINATAIRE").Range
    MarquerParagraphe doc, 5, "APPEL"
    modCourrier.EncadrerDestinataireDragon doc
    Set zone = modCourrier.ZoneDestinataireDragon(doc)
    ExigerC4 RefuseCorpsChevauchant(doc, doc.Range(zone.Start, zone.End)), _
             "corps chevauchant directement le destinataire refuse"
    ExigerC4 modIntegrationUnifie.LocaliserCorpsUnifie(doc, corps, premier), _
             "appel reel localise apres adresse commencant par Madame, Monsieur"
    ExigerC4 corps.Start >= zone.End, "corps strictement hors de la zone destinataire"
    ExigerC4 InStr(1, corps.Text, "symptome fictif", vbTextCompare) > 0 And _
             InStr(1, corps.Text, "SERVICE DESTINATAIRE", vbTextCompare) = 0 And _
             InStr(1, corps.Text, "Beaumont, le", vbTextCompare) = 0, _
             "ni adresse ni date incluses dans le corps a remplacer"
    doc.Close wdDoNotSaveChanges: Set doc = Nothing
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise numero, "TesterAppelDansDestinataire C4", description
End Sub

Public Function ExecuterC4Destinataire(ByVal sortie As String) As String
    Dim fso As Object, racine As String, racineAvant As String, description As String
    On Error GoTo Echec
    mNombre = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(sortie) < 3 Or Mid$(sortie, 2, 2) <> ":\" Then Err.Raise vbObjectError + 1494, , "Sortie locale absolue requise."
    If fso.GetDrive(fso.GetDriveName(sortie)).DriveType <> 2 Then Err.Raise vbObjectError + 1494, , "Disque local requis."
    racineAvant = modConfig.racine()
    racine = sortie & "\c4-destinataire-" & modFichiers.IdUnique()
    fso.CreateFolder racine
    fso.CreateFolder racine & "\Config"
    fso.CreateFolder racine & "\Modeles"
    CreerFixture racine & "\Modeles\historique-c4.dotx", False
    CreerFixture racine & "\Modeles\moderne-c4.dotx", True
    TesterParcours racine, "historique-c4", False
    TesterParcours racine, "moderne-c4", True
    ExigerC4 modCourrier.CleDestinataire("AB C" & vbCr & "12 RUE") <> _
             modCourrier.CleDestinataire("A BC" & vbCr & "12 RUE"), _
             "les espaces entre mots restent significatifs"
    ExigerC4 modCourrier.CleDestinataire("16 rue de L" & ChrW$(&H2019) & "Isle-Adam") = _
             modCourrier.CleDestinataire("16 rue de L'Isle-Adam"), _
             "apostrophe Dragon et apostrophe typographique equivalentes"
    TesterPatientMemorise
    TesterAppelDansDestinataire
    ExigerC4 modFileArrivees.ProfilMedecinFile("CabinetMedecin"), _
             "fenetre des arrivees au demarrage du profil reel"
    ExigerC4 Not modFileArrivees.ProfilMedecinFile("Secretariat"), _
             "aucune fenetre medecin au secretariat"
    ExigerC4 modValidation.NomBaseCourrierPatient("DUPONT", "Alice", _
             "03/10/2026 23:41:12") = "DUPONT Alice 202610032341", _
             "courrier valide nomme Nom Prenom aaaammjjhhmm"
    ExigerC4 modValidation.NomBaseCourrierPatient("DUP/ONT", "Alice", _
             "03/10/2026 23:41:12") = "DUP ONT Alice 202610032341", _
             "caracteres interdits retires du nom de courrier"
    If mNombre <> ATTENDUS_C4 Then Err.Raise vbObjectError + 1494, , "Nombre de controles C4 inattendu."
    ExecuterC4Destinataire = "{""reussis"":" & CStr(mNombre) & ",""attendus"":" & _
        CStr(ATTENDUS_C4) & ",""echec"":false}"
Sortie:
    On Error Resume Next
    modConfig.DefinirRacine racineAvant
    On Error GoTo 0
    Exit Function
Echec:
    description = Err.Description
    ExecuterC4Destinataire = "{""reussis"":" & CStr(mNombre) & ",""attendus"":" & _
        CStr(ATTENDUS_C4) & ",""echec"":true,""description"":" & modServiceNas.JsonValeur(description) & "}"
    Resume Sortie
End Function
