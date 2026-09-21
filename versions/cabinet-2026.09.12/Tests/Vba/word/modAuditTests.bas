Attribute VB_Name = "modAuditTests"
Option Explicit
Private mVerifications As Long

' Execute uniquement sur Windows/Word. Donnees fictives, aucun appel API, NAS ou imprimante.
Public Function Audit_TestsSansReseau() As String
    Dim dossier As String, doc As Document, pat As Object, racineJson As Object
    Dim texte As String, lignes As Variant, ligne As Variant, r As Range, premier As Range
    Dim n As Long, description As String, valeur As Variant, invalide As Variant
    On Error GoTo Echec
    mVerifications = 0
    dossier = Environ$("TEMP") & "\CabinetAudit-" & modFichiers.IdUnique()
    modFichiers.EnsureDossier dossier & "\Config"
    modFichiers.EcrireTexteUTF8 dossier & "\Config\config.ini", "[ECG]" & vbCrLf & "CodeExamen=EKG01"
    modConfig.DefinirRacine dossier
    Audit_Verifier modTexte.DateFrValide("29/02/2024"), "annee bissextile"
    Audit_Verifier Not modTexte.DateFrValide("29/02/2023"), "faux 29 fevrier"
    Audit_Verifier Not modTexte.DateFrValide("31/04/2024"), "jour inexistant"
    Audit_Verifier modTexte.SexeNormalise("Femme") = "F", "sexe explicite"
    Audit_Verifier Len(modTexte.SexeNormalise("Inconnu")) = 0, "aucun sexe invente"
    Audit_Verifier modTexte.IntervallesSeChevauchent(600, 30, 615, 15), "chevauchement agenda"
    Audit_Verifier Not modTexte.IntervallesSeChevauchent(600, 15, 615, 30), "creneaux contigus"

    Set racineJson = modJson.JsonParse("{""a"":[{""x"":1},2,true,null,{""x"":3}],""b"":""ete\n\u00E9""}")
    Audit_Verifier racineJson("a").Count = 5, "tableau JSON mixte"
    Audit_Verifier racineJson("a")(1)("x") = 1 And racineJson("a")(5)("x") = 3, "objets JSON imbriques"
    Audit_Verifier racineJson("a")(2) = 2 And IsNull(racineJson("a")(4)), "scalaires apres objet JSON"
    Audit_Verifier racineJson("b") = "ete" & vbLf & ChrW$(233), "echappements JSON"
    For Each invalide In Array("{}x", "{""x"":1,}", "[1,]", "{""x"":1,""x"":2}", "{""x"":01}", "{""x"":""\uZZZZ""}", "{""x"":")
        Audit_Verifier Audit_JsonRefuse(CStr(invalide)), "JSON invalide refuse"
    Next invalide
    texte = "{""status"":""completed"",""output"":[{""type"":""message"",""content"":[{""type"":""output_text"",""text"":""texte valide""}]}]}"
    Audit_Verifier Audit_ReponseRefusee(texte), "Ancien texte libre refuse"
    Dim payload As Object, demandes As New Collection
    Set payload = modServiceNas.Parametres()
    payload("corps_courrier") = "texte valide"
    Set payload("demandes") = demandes
    texte = Replace(texte, "texte valide", modJson.JsonEchapper(modServiceNas.JsonValeur(payload)))
    Audit_Verifier InStr(ExtraireTexteOpenAI(texte), "texte valide") > 0, "Responses schema JSON completed"
    Audit_Verifier Audit_ReponseRefusee(Replace(texte, "completed", "incomplete")), "Responses incomplete refuse meme avec texte"
    Audit_Verifier Audit_ReponseRefusee(Replace(texte, "completed", "failed")), "Responses failed refuse"
    Audit_Verifier Audit_ReponseRefusee(Replace(texte, "output_text", "refusal")), "refus API"

    Set pat = CreateObject("Scripting.Dictionary")
    pat("ID") = "TEST-AUDIT-001": pat("Nom") = "FICTIF": pat("Prenom") = "Elodie"
    pat("DDN") = "29/02/1960": pat("Sexe") = "F"
    texte = modGdt.ConstruireGdt(pat)
    lignes = Split(texte, vbCrLf)
    Audit_Verifier lignes(0) = "01380006302", "GDT satz 6302"
    Audit_Verifier CLng(Mid$(lignes(1), 8)) = Len(texte), "GDT longueur totale"
    Audit_Verifier InStr(texte, vbCrLf & "0103110") = 0 And InStr(texte, vbCrLf & "0173103") = 0, "GDT sans sexe ni naissance"
    For Each ligne In lignes
        If Len(ligne) > 0 Then Audit_Verifier CLng(Left$(ligne, 3)) = Len(ligne) + 2, "GDT longueur de ligne"
    Next ligne
    pat("Sexe") = "": Audit_Verifier modGdt.ConstruireGdt(pat) = texte, "GDT minimal independant du sexe"
    pat("Sexe") = "F": pat("DDN") = "31/02/1960": Audit_Verifier modGdt.ConstruireGdt(pat) = texte, "GDT minimal ne transmet jamais la naissance"
    pat.Remove "Sexe": pat.Remove "DDN"
    Audit_Verifier modGdt.ConstruireGdt(pat) = texte, "GDT accepte uniquement ID nom prenom"
    pat("ID") = "P1234567890aaaaaaaaaaaaaaaaaaaaaa1"
    texte = modGdt.ConstruireGdt(pat)
    Audit_Verifier InStr(texte, "3000" & pat("ID") & vbCrLf) > 0, "GDT identifiant long integral"
    pat("ID") = "P1234567890aaaaaaaaaaaaaaaaaaaaaa2"
    Audit_Verifier modGdt.ConstruireGdt(pat) <> texte, "GDT deux identifiants de meme prefixe distincts"
    pat("Prenom") = "El" & ChrW$(233) & "odie"
    texte = modGdt.ConstruireGdt(pat)
    Audit_Verifier InStr(texte, "3102" & pat("Prenom") & vbCrLf) > 0, "GDT accent CP1252 conserve"
    pat("DDN") = "29/02/1960": pat("Nom") = "FICTIF" & vbCrLf & "31101"
    Audit_Verifier Audit_GdtRefuse(pat), "GDT injection de champ refusee"
    pat("Nom") = ChrW$(&H4E2D): Audit_Verifier Audit_GdtRefuse(pat), "GDT caractere hors CP1252 refuse"
    Audit_Verifier modFichiers.NomFichierSur("NUL") <> "NUL", "nom reserve Windows"
    Audit_Verifier Right$(modFichiers.NomFichierSur(String$(119, "a") & ".suite"), 1) <> ".", "troncature du nom de fichier"

    Set doc = Documents.Add
    doc.Content.Text = "Entete" & vbCr & "Cher Confrere," & vbCr & "Introduction clinique." & vbCr & _
                       "Monsieur FICTIF Test, 60 ans, examen." & vbCr & "Derniere donnee medicale." & vbCr & _
                       "Bien cordialement." & vbCr & "Docteur Olivier MANDAGOUT" & vbCr
    Audit_Verifier modIntegrationUnifie.LocaliserCorpsUnifie(doc, r, premier), "corps localise"
    Audit_Verifier InStr(r.Text, "Introduction clinique") > 0 And InStr(r.Text, "Derniere donnee medicale") > 0, "premier et dernier paragraphes conserves"
    Audit_Verifier InStr(r.Text, "Bien cordialement") = 0 And InStr(r.Text, "MANDAGOUT") = 0, "politesse et signature exclues"
    Audit_Verifier doc.Bookmarks.Exists("CORPS"), "signet CORPS recree"
    Audit_TestsSansReseau = CStr(mVerifications) & " verifications VBA reussies. Aucun test Dragon/ECG/imprimante/API."
Sortie:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close wdDoNotSaveChanges
    modConfig.DefinirRacine ""
    If Len(dossier) > 0 Then CreateObject("Scripting.FileSystemObject").DeleteFolder dossier, True
    On Error GoTo 0
    If n <> 0 Then Err.Raise n, "Audit_TestsSansReseau", description
    Exit Function
Echec:
    n = Err.Number: description = Err.Description
    Resume Sortie
End Function

Public Sub Audit_LancerTests()
    On Error GoTo Echec
    MsgBox Audit_TestsSansReseau(), vbInformation, "Recette VBA"
    Exit Sub
Echec:
    MsgBox Err.Description, vbCritical, "Recette VBA en echec"
End Sub

Private Sub Audit_Verifier(ByVal condition As Boolean, ByVal nom As String)
    If Not condition Then Err.Raise vbObjectError + 990, "Audit", "ECHEC : " & nom
    mVerifications = mVerifications + 1
End Sub

Private Function Audit_JsonRefuse(ByVal texte As String) As Boolean
    Dim valeur As Variant
    On Error GoTo Attendu
    Set valeur = modJson.JsonParse(texte)
    Exit Function
Attendu:
    Audit_JsonRefuse = True
End Function

Private Function Audit_ReponseRefusee(ByVal texte As String) As Boolean
    Dim sortie As String
    On Error GoTo Attendu
    sortie = ExtraireTexteOpenAI(texte)
    Exit Function
Attendu:
    Audit_ReponseRefusee = True
End Function

Private Function Audit_GdtRefuse(ByVal patient As Object) As Boolean
    Dim sortie As String
    On Error GoTo Attendu
    sortie = modGdt.ConstruireGdt(patient)
    Exit Function
Attendu:
    Audit_GdtRefuse = True
End Function

Public Function Audit_ExecuterJson() As String
    Dim resultat As String, description As String
    On Error GoTo Echec
    resultat = Audit_TestsSansReseau()
    Audit_ExecuterJson = "{""reussis"":" & CStr(mVerifications) & ",""echec"":false}"
    Exit Function
Echec:
    description = Err.Description
    Audit_ExecuterJson = "{""reussis"":" & CStr(mVerifications) & ",""echec"":true,""description"":" & modServiceNas.JsonValeur(description) & "}"
End Function
