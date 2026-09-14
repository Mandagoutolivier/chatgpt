Attribute VB_Name = "modRecetteU1"
Option Explicit
Private mNombre As Long

Private Sub Exiger(ByVal condition As Boolean, ByVal nom As String)
    If Not condition Then Err.Raise vbObjectError + 1190, "Recette U1", nom
    mNombre = mNombre + 1
End Sub

Public Function Executer(ByVal dossier As String) As String
    Dim d As Object, r As Object, col As Collection, v As Variant, items As Collection, ligne As Object
    Dim i As Long, numero As Long, description As String, doc As Document, h1 As String, h2 As String, id As String
    On Error GoTo Echec
    mNombre = 0
    Exiger modServiceNas.SHA256("abc") = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "SHA256 abc"
    Exiger modServiceNas.SHA256("") = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855", "SHA256 vide"
    Exiger modServiceNas.TexteUTF8(modServiceNas.OctetsUTF8(("FICTIF é µ œ " & ChrW(&H6F22) & ChrW(&H5B57)))) = ("FICTIF é µ œ " & ChrW(&H6F22) & ChrW(&H5B57)), "UTF8 Windows"
    Exiger modServiceNas.JsonValeur(Empty) = "null", "JSON Empty"
    Exiger modServiceNas.JsonValeur(Null) = "null", "JSON Null"
    Exiger modServiceNas.JsonValeur(True) = "true", "JSON booleen"
    Exiger modServiceNas.JsonValeur(12.5) = "12.5", "JSON decimal"
    Exiger modServiceNas.JsonValeur(DateSerial(2026, 9, 14)) = Chr$(34) & "2026-09-14T00:00:00" & Chr$(34), "JSON date ISO"
    Set d = modServiceNas.Parametres(): d("z") = "FICTIF": d("a") = "é"
    Set r = modServiceNas.Parametres(): r("a") = "é": r("z") = "FICTIF"
    Exiger modServiceNas.JsonValeur(d) = modServiceNas.JsonValeur(r), "Commande canonique"
    Set r = modServiceNas.DecoderReponse(200, modServiceNas.OctetsUTF8("{""result"":{""Libelle"":""FICTIF é""}}"))
    Exiger CStr(r("Libelle")) = "FICTIF é", "Reponse Unicode"
    On Error Resume Next
    Set r = modServiceNas.DecoderReponse(502, modServiceNas.OctetsUTF8("<html>proxy</html>"))
    numero = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo Echec
    Exiger numero = vbObjectError + 1103 And InStr(description, "502") > 0, "HTML 502 conserve statut"
    On Error Resume Next
    Set r = modServiceNas.DecoderReponse(409, modServiceNas.OctetsUTF8("{""error"":""FICTIF"",""code"":""destination_absente""}"))
    numero = Err.Number: Err.Clear
    On Error GoTo Echec
    Exiger numero = vbObjectError + 1141, "Destination manquante distincte"
    On Error Resume Next
    Set r = modServiceNas.DecoderReponse(200, modServiceNas.OctetsUTF8("{}"))
    numero = Err.Number: Err.Clear
    On Error GoTo Echec
    Exiger numero = vbObjectError + 1105, "Enveloppe incomplete refusee"
    On Error Resume Next
    Set r = modServiceNas.DecoderReponse(502, modServiceNas.OctetsUTF8("[]"))
    numero = Err.Number: Err.Clear
    On Error GoTo Echec
    Exiger numero = vbObjectError + 1103, "Erreur HTTP avec tableau JSON"
    On Error Resume Next
    Set r = modServiceNas.DecoderReponse(409, modServiceNas.OctetsUTF8("{""error"":null,""code"":{}}"))
    numero = Err.Number: Err.Clear
    On Error GoTo Echec
    Exiger numero = vbObjectError + 1103, "Erreur HTTP avec champs mal types"
    On Error Resume Next
    Set r = modServiceNas.DecoderReponse(200, modServiceNas.OctetsUTF8("{""result"":[]}"))
    numero = Err.Number: Err.Clear
    On Error GoTo Echec
    Exiger numero = vbObjectError + 1105, "Resultat RPC objet obligatoire"
    Set items = New Collection
    For i = 1 To 3
        Set ligne = modServiceNas.Parametres()
        ligne("DateArrivee") = IIf(i = 3, "13/09/2026", "14/09/2026"): ligne("Statut") = "Arrive"
        ligne("PatientID") = "FICTIF": ligne("RdvID") = CStr(i): ligne("ConsultationID") = "consult-FICTIF-" & CStr(i)
        ligne("Nom") = "FICTIF" & vbTab & "TEST": ligne("Prenom") = "Essai": ligne("DDN") = "01/01/1980"
        ligne("HeureArrivee") = IIf(i = 1, "10:00", "09:00"): ligne("HeureRdv") = "08:30"
        items.Add ligne
    Next i
    Set col = modAttenteLocale.FiltrerEtTrier(items, "14/09/2026")
    Exiger col.Count = 2 And col(1)("RdvID") = "2", "File filtre et tri sans SQLite"
    Exiger InStr(CStr(col(1)("Nom")), vbTab) > 0, "Tabulation conservee"
    Exiger modTexte.HeureValide("08H30") And modTexte.MinutesDepuisMinuit("08H30") = 510, "Heure majuscule"
    Set d = modServiceNas.Parametres(): d("corps_courrier") = "[[PATIENT]] FICTIF": Set d("demandes") = New Collection
    For i = 1 To 2
        Set r = modServiceNas.Parametres(): r("cle_destination") = "A_COMPLETER": r("corps") = "[[PATIENT]] FICTIF " & CStr(i)
        d("demandes").Add r
    Next i
    Set r = modOpenAI_v22_corrige.ValiderStructure(d)
    gReponseAPICabinetTest = modServiceNas.JsonValeur(r)
    Set col = CT_ConstruireCollectionDemandes()
    Exiger col.Count = 2 And col(1)("Numero") <> col(2)("Numero"), "Deux destinations inconnues distinctes"
    modEtatCourrier.DefinirDossierTests dossier & "\EtatsFictifs"
    id = modFichiers.IdUnique()
    modEtatCourrier.EcrireProtege id, "-source", "SOURCE FICTIVE é", True
    Exiger modEtatCourrier.LireProtege(id, "-source") = "SOURCE FICTIVE é", "DPAPI source relue"
    On Error Resume Next
    modEtatCourrier.EcrireProtege id, "-source", "REMPLACEMENT", True
    numero = Err.Number: Err.Clear
    On Error GoTo Echec
    Exiger numero <> 0 And modEtatCourrier.LireProtege(id, "-source") = "SOURCE FICTIVE é", "Source immuable"
    modEtatCourrier.EcrireProtege id, "-etat", "ETAPE 1"
    modEtatCourrier.EcrireProtege id, "-etat", "ETAPE 2"
    Exiger modEtatCourrier.LireProtege(id, "-etat") = "ETAPE 2", "Etat atomique remplace"
    Set doc = Documents.Add
    doc.Content.Text = "COURRIER FICTIF : TEST 5 mg."
    h1 = modControleCourrier.EmpreinteCourrier(doc, "COR-FICTIF")
    doc.SaveAs2 dossier & "\revision-fictive.docx", wdFormatXMLDocument
    h2 = modControleCourrier.EmpreinteCourrier(doc, "COR-FICTIF")
    Exiger h1 = h2, "Revision stable apres sauvegarde"
    modIntegrationUnifie.FixerVariable doc, "PublicationID", modFichiers.IdUnique()
    Exiger h2 = modControleCourrier.EmpreinteCourrier(doc, "COR-FICTIF"), "Variables de reprise hors revision"
    doc.Content.InsertAfter " MODIFICATION"
    Exiger h2 <> modControleCourrier.EmpreinteCourrier(doc, "COR-FICTIF"), "Revision change avec texte"
    doc.Close wdDoNotSaveChanges: Set doc = Nothing
    modEtatCourrier.DefinirDossierTests ""
    Executer = "{""reussis"":" & CStr(mNombre) & ",""echec"":false}"
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    modEtatCourrier.DefinirDossierTests ""
    On Error GoTo 0
    Executer = "{""reussis"":" & CStr(mNombre) & ",""echec"":true,""numero"":" & CStr(numero) & ",""description"":" & modServiceNas.JsonValeur(description) & "}"
End Function
