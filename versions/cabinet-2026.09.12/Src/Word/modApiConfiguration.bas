Attribute VB_Name = "modApiConfiguration"
Option Explicit
' =====================================================================
' modApiConfiguration - Appel de l'API Claude (correction et lettres derivees).
'  - cle API : %APPDATA%\CabinetCardio\api.key (poste medecin uniquement)
'  - envoi UTF-8 (ADODB.Stream), reponse analysee par modJson
'  - reessais sur 429/5xx et erreurs reseau
'  - mode Debug (config.ini [API] Debug=1) : le contenu exact envoye est
'    ecrit dans Logs\payload_debug.json (recette sur patients FICTIFS)
' + Orchestration de la commande "CorrigerCourrier" (Ctrl+Alt+Maj+C).
' =====================================================================

Private Const URL_API As String = "https://api.anthropic.com/v1/messages"
Private Const URL_OPENAI As String = "https://api.openai.com/v1/responses"

' Fournisseur configure : "claude" (defaut) ou "openai" (config.ini [API] Fournisseur)
Public Function FournisseurApi() As String
    FournisseurApi = "openai"
End Function
' Cle OpenAI : variable d'environnement OPENAI_API_KEY (poste du cabinet),
' sinon %APPDATA%\CabinetCardio\openai.key
Public Function LireCleOpenAICabinet() As String
    Dim chemin As String
    LireCleOpenAICabinet = Trim$(Environ$("OPENAI_API_KEY"))
    If Len(LireCleOpenAICabinet) > 0 Then Exit Function
    chemin = Environ$("APPDATA") & "\CabinetCardio\openai.key"
    If modFichiers.FichierExiste(chemin) Then
        LireCleOpenAICabinet = Trim$(Replace(Replace(modFichiers.LireTexteUTF8(chemin), vbCr, ""), vbLf, ""))
    End If
    If Len(LireCleOpenAICabinet) = 0 Then
        Err.Raise vbObjectError + 400, "modApiConfiguration", _
            "Cle OpenAI introuvable : definissez la variable d'environnement OPENAI_API_KEY" & vbCrLf & _
            "(ou creez le fichier " & chemin & ")."
    End If
End Function

Public Function LireCleApi() As String
    Dim chemin As String
    chemin = Environ$("APPDATA") & "\CabinetCardio\api.key"
    If Not modFichiers.FichierExiste(chemin) Then
        Err.Raise vbObjectError + 400, "modApiConfiguration", _
            "Cle API introuvable : " & chemin & vbCrLf & _
            "Creez ce fichier contenant uniquement votre cle API Anthropic."
    End If
    LireCleApi = Trim$(Replace(Replace(modFichiers.LireTexteUTF8(chemin), vbCr, ""), vbLf, ""))
End Function

' Appel synchrone. Renvoie le texte de la reponse ; leve une erreur claire sinon.
Public Function AppelerClaude(ByVal systeme As String, ByVal utilisateur As String, Optional ByVal fournisseur As String = "") As String
    Err.Raise vbObjectError + 439, "modApiConfiguration", "Ancien moteur desactive. Utilisez Finaliser (PowerMic D) pour le flux PROD6 unifie."
End Function
' Chaine -> octets UTF-8 sans BOM (pour http.Send)

' Prompt systeme : fichier Config\prompts\<nom> + courriers de reference
' (avecReferences=False : le marqueur est simplement retire - les lettres
' derivees ne doivent PAS imiter les comptes rendus de consultation)
Public Function ChargerPrompt(ByVal nomFichier As String, _
                              Optional ByVal avecReferences As Boolean = True) As String
    Dim chemin As String, prompt As String, refs As String
    chemin = modConfig.chemin("Config") & "\prompts\" & nomFichier
    If Not modFichiers.FichierExiste(chemin) Then
        Err.Raise vbObjectError + 406, "modApiConfiguration", "Prompt introuvable : " & chemin
    End If
    prompt = modFichiers.LireTexteUTF8(chemin)
    If avecReferences Then refs = ChargerReferences() Else refs = ""
    ChargerPrompt = Replace(prompt, "[COURRIERS_DE_REFERENCE]", refs)
End Function

Private Function ChargerReferences() As String
    Dim dossier As String, f As Variant, refs As String, n As Long, contenu As String
    Const MAX_REFS As Long = 3
    Const MAX_CARS As Long = 5000
    dossier = modConfig.chemin("Config") & "\style"
    For Each f In modFichiers.ListerFichiers(dossier, ".txt")
        If n >= MAX_REFS Then Exit For
        contenu = modFichiers.LireTexteUTF8(CStr(f))
        If Len(contenu) > MAX_CARS Then contenu = Left$(contenu, MAX_CARS)
        n = n + 1
        refs = refs & vbCrLf & "--- Courrier de référence " & n & " ---" & vbCrLf & contenu & vbCrLf
    Next f
    If n > 0 Then
        ChargerReferences = "Courriers de référence (style à imiter fidèlement) :" & vbCrLf & refs
    Else
        ChargerReferences = ""
    End If
End Function

' =====================================================================
' Commande principale : CORRIGER LE COURRIER (Ctrl+Alt+Maj+C / voix)
' =====================================================================
Public Sub CorrigerCourrier()
    modPowerMicUnifie.Unifie_D_Finaliser
End Sub
' Corps de la correction, reutilisable par FinaliserCourrier (une touche :
' corriger + demandes + enregistrer + transmettre). Renvoie True si le
' texte a ete remplace. silencieux : pas de message de succes.
Public Function CorrigerDocument(ByVal doc As Document, ByVal silencieux As Boolean) As Boolean
    Err.Raise vbObjectError + 988, "Compatibilite U2", "Moteur historique retire. Utilisez Unifie_D_Finaliser pour conserver les controles et reprises."
End Function

' Normalisation de la reponse : fins de ligne Word ; les lignes vides entre
' paragraphes (habitude de l'IA) sont supprimees - dans les courriers du
' cabinet, l'espacement de 12 pt separe les paragraphes, jamais une ligne vide.
Public Function NettoyerReponse(ByVal t As String) As String
    Dim lignes() As String, i As Long, res As String, l As String
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    lignes = Split(t, vbLf)
    For i = LBound(lignes) To UBound(lignes)
        l = Trim$(lignes(i))
        If Len(l) > 0 Then
            If Len(res) > 0 Then res = res & vbCr
            res = res & l
        End If
    Next i
    NettoyerReponse = res
End Function

Public Function PatientDuDocument(ByVal doc As Document) As Object
    On Error Resume Next
    Dim id As String
    id = doc.Variables("PatientID")
    On Error GoTo 0
    If Len(id) > 0 Then Set PatientDuDocument = modBase.PatientParID(id)
End Function

Public Function CorrespondantDuDocument(ByVal doc As Document) As Object
    On Error Resume Next
    Dim id As String
    id = doc.Variables("CorrespondantID")
    On Error GoTo 0
    If Len(id) > 0 Then Set CorrespondantDuDocument = modBase.CorrespondantParID(id)
End Function

Private Sub ArchiverBrouillon(ByVal pat As Object, ByVal corps As String)
    On Error Resume Next
    modFichiers.EcrireTexteUTF8 modPatient.DossierPatient(pat) & "\brouillon_" & _
                                modFichiers.IdUnique() & ".txt", corps
End Sub
