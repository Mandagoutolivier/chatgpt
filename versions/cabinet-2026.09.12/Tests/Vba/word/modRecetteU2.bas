Attribute VB_Name = "modRecetteU2"
Option Explicit
Private mNombre As Long

Private Sub ExigerU2(ByVal condition As Boolean, ByVal nom As String)
    If Not condition Then Err.Raise vbObjectError + 1192, "Recette U2", nom
    mNombre = mNombre + 1
End Sub

Public Function ExecuterU2(ByVal sortie As String) As String
    Dim p As Object, q As Object, fso As Object, ts As Object, cree As Boolean
    Dim dossier As String, cle As String, ancien As String, chemin As String, id As String, id2 As String
    Dim reprise As Boolean, numero As Long, description As String
    On Error GoTo Echec
    mNombre = 0
    Set p = modServiceNas.Parametres(): p("z") = "FICTIF": p("a") = "1"
    Set q = modServiceNas.Parametres(): q("a") = "1": q("z") = "FICTIF"
    cle = modCommandesLocales.CleCommandeLocale("https://fictif", "compte-a", "claim", p)
    ExigerU2 cle = modCommandesLocales.CleCommandeLocale("https://fictif", "compte-a", "claim", q), "cle canonique"
    ExigerU2 cle <> modCommandesLocales.CleCommandeLocale("https://fictif", "compte-b", "claim", p), "compte distinct"
    ExigerU2 cle <> modCommandesLocales.CleCommandeLocale("https://autre", "compte-a", "claim", p), "serveur distinct"
    ExigerU2 cle <> modCommandesLocales.CleCommandeLocale("https://fictif", "compte-a", "release", p), "operation distincte"
    q("a") = "2"
    ExigerU2 cle <> modCommandesLocales.CleCommandeLocale("https://fictif", "compte-a", "claim", q), "contenu distinct"
    Set fso = CreateObject("Scripting.FileSystemObject")
    dossier = sortie & "\commandes-U2-" & modFichiers.IdUnique()
    fso.CreateFolder dossier: cree = True
    id = modCommandesLocales.ObtenirCommandeLocale(cle, chemin, reprise, modServiceNas.SHA256("ancien jeton fictif"), dossier)
    id2 = modCommandesLocales.ObtenirCommandeLocale(cle, chemin, reprise, modServiceNas.SHA256("nouveau jeton fictif"), dossier)
    ExigerU2 reprise And id = id2, "rotation sans nouvel identifiant"
    cle = modServiceNas.SHA256("migration fictive"): ancien = modServiceNas.SHA256("ancienne cle fictive")
    Set ts = fso.CreateTextFile(dossier & "\" & ancien & ".pending", False, False): ts.Write String$(32, "a"): ts.Close
    id = modCommandesLocales.ObtenirCommandeLocale(cle, chemin, reprise, ancien, dossier)
    ExigerU2 reprise And id = String$(32, "a") And Not fso.FileExists(dossier & "\" & ancien & ".pending"), "migration U1 conserve identifiant"
    Set ts = fso.CreateTextFile(dossier & "\" & ancien & ".pending", False, False): ts.Write String$(32, "b"): ts.Close
    On Error Resume Next
    id2 = modCommandesLocales.ObtenirCommandeLocale(cle, chemin, reprise, ancien, dossier)
    numero = Err.Number: Err.Clear
    On Error GoTo Echec
    ExigerU2 numero = vbObjectError + 1108, "identifiants contradictoires refuses"
    ExigerU2 fso.FileExists(dossier & "\" & ancien & ".pending") And fso.FileExists(chemin), "fichiers contradictoires conserves"
    ExigerU2 modCycleCourrier.DocumentSource() Is Nothing And Not modCycleCourrier.CycleEnCours(), "contexte de correction libere"
    ExecuterU2 = "{""reussis"":" & CStr(mNombre) & ",""echec"":false}"
Sortie:
    On Error Resume Next
    If cree Then fso.DeleteFolder dossier, True
    On Error GoTo 0
    Exit Function
Echec:
    description = Err.Description
    ExecuterU2 = "{""reussis"":" & CStr(mNombre) & ",""echec"":true,""description"":" & modServiceNas.JsonValeur(description) & "}"
    Resume Sortie
End Function
