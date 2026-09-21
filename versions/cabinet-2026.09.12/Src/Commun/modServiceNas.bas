Attribute VB_Name = "modServiceNas"
Option Explicit

Public Function Parametres() As Object
    Set Parametres = CreateObject("Scripting.Dictionary")
    Parametres.CompareMode = 1
End Function

Public Function JsonValeur(ByVal valeur As Variant) As String
    JsonValeur = modDonneesTransport.SerialiserJson(valeur)
End Function
Public Function OctetsUTF8(ByVal texte As String) As Variant
    OctetsUTF8 = modDonneesTransport.EncoderUTF8(texte)
End Function
Public Function TexteUTF8(ByVal octets As Variant) As String
    TexteUTF8 = modDonneesTransport.DecoderUTF8(octets)
End Function
Public Function SHA256(ByVal texte As String) As String
    SHA256 = modDonneesTransport.EmpreinteSHA256(texte)
End Function
Public Function EstLecture(ByVal operation As String) As Boolean
    Select Case operation
        Case "whoami", "table.read", "record.get", "attentes", "reprises", "publications", "correspondent.resolve", "clinical.compare", "dictionary.read", "journal.read", "nir.validate", "command.result", "billing.get", "stale_arrivals", "publication.get": EstLecture = True
    End Select
End Function

Public Function Appeler(ByVal operation As String, ByVal params As Object) As Object
    Dim requete As Object, reponse As Object, http As Object, lookup As Object, check As Object
    Dim base As String, jeton As String, payload As String, chemin As String, id As String, octets As Variant
    Dim reprise As Boolean, numero As Long, description As String, statut As Long, etape As String
    Dim session As Object, compte As String, ancienneCle As String
    On Error GoTo Echec
    etape = "configuration"
    base = Trim$(modFichiers.LireTexteUTF8(Environ$("APPDATA") & "\CabinetCardio\service.url"))
    If LCase$(Left$(base, 8)) <> "https://" Then Err.Raise vbObjectError + 1100, "Service NAS", "Adresse HTTPS du service absente ou invalide."
    Do While Right$(base, 1) = "/": base = Left$(base, Len(base) - 1): Loop
    jeton = Trim$(modFichiers.LireTexteUTF8(Environ$("APPDATA") & "\CabinetCardio\service.token"))
    If Len(jeton) < 32 Or InStr(jeton, vbCr) Or InStr(jeton, vbLf) Then Err.Raise vbObjectError + 1101, , "Identifiant de connexion du poste absent ou invalide."
    If Not EstLecture(operation) Then
        etape = "commande_durable"
        Set session = Appeler("whoami", Parametres())
        compte = CStr(session("ID"))
        ancienneCle = SHA256(base & "|" & jeton & "|" & operation & "|" & JsonValeur(params))
        id = modCommandesLocales.ObtenirCommandeLocale(modCommandesLocales.CleCommandeLocale(base, compte, operation, params), chemin, reprise, ancienneCle)
        If reprise Then
            Set lookup = Parametres(): lookup("id") = id
            Set check = Appeler("command.result", lookup)
            If Not check.Exists("trouve") Then Err.Raise vbObjectError + 1105, , "Reponse de reprise incomplete."
            If CBool(check("trouve")) Then
                If Not check.Exists("resultat") Then Err.Raise vbObjectError + 1105
                Set Appeler = check("resultat")
                modFichiers.SupprimerTemporaire chemin
                Exit Function
            End If
        End If
    End If
    Set requete = Parametres(): requete("operation") = operation
    Set requete("params") = params: requete("request_id") = id
    payload = JsonValeur(requete): octets = OctetsUTF8(payload)
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 5000, 10000, 30000, 30000
    http.Open "POST", base & "/v1/rpc", False
    http.Option(6) = False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Authorization", "Bearer " & jeton
    etape = "transport": http.Send octets
    statut = http.Status
    octets = http.ResponseBody
    If statut >= 400 And statut < 500 And statut <> 408 And statut <> 429 Then modFichiers.SupprimerTemporaire chemin
    etape = "reponse"
    Set Appeler = DecoderReponse(statut, octets)
    modFichiers.SupprimerTemporaire chemin
    modLog.Diagnostic etape, "succes", 0, id
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    modLog.Diagnostic etape, "echec", numero, id
    On Error GoTo 0
    Err.Raise numero, "Service NAS", description
End Function

Public Function ItemsValides(ByVal r As Object) As Collection
    Dim item As Variant
    If TypeName(r) <> "Dictionary" Then Err.Raise vbObjectError + 1105, , "Enveloppe invalide."
    If Not r.Exists("items") Then Err.Raise vbObjectError + 1105, , "Liste absente."
    If TypeName(r("items")) <> "Collection" Then Err.Raise vbObjectError + 1105, , "Liste invalide."
    For Each item In r("items")
        If TypeName(item) <> "Dictionary" Then Err.Raise vbObjectError + 1105, , "Element de liste invalide."
    Next item
    Set ItemsValides = r("items")
End Function

Public Function LirePages(ByVal operation As String, ByVal p As Object) As Collection
    Dim r As Object, item As Object, resultat As New Collection, page As Long, suivant As Variant
    p("offset") = 0
    For page = 1 To 5000
        Set r = Appeler(operation, p)
        For Each item In ItemsValides(r): resultat.Add item: Next item
        If Not r.Exists("next") Then Err.Raise vbObjectError + 1105, , "Pagination incomplete."
        suivant = r("next")
        If IsNull(suivant) Then Set LirePages = resultat: Exit Function
        If VarType(suivant) <> vbDouble And VarType(suivant) <> vbLong And VarType(suivant) <> vbInteger Then Err.Raise vbObjectError + 1105, , "Pagination invalide."
        If CDbl(suivant) <= CDbl(p("offset")) Or CDbl(suivant) > 1000000 Or CDbl(suivant) <> Fix(CDbl(suivant)) Then Err.Raise vbObjectError + 1105, , "Pagination sans progression."
        p("offset") = CLng(suivant)
    Next page
    Err.Raise vbObjectError + 1105, , "Trop de pages : affinez la recherche."
End Function

Public Function LireTable(ByVal genre As String, Optional ByVal recherche As String = "", Optional ByVal annee As String = "", Optional ByVal jour As String = "") As Collection
    Dim p As Object
    Set p = Parametres(): p("genre") = genre: p("q") = recherche: p("year") = annee: p("date") = jour
    Set LireTable = LirePages("table.read", p)
End Function

Public Function LireID(ByVal genre As String, ByVal id As String) As Object
    Dim p As Object
    Set p = Parametres(): p("genre") = genre: p("id") = id
    Set LireID = Appeler("record.get", p)
End Function

Public Function CommandeID(ByVal operation As String, ByVal id As String) As Object
    Dim p As Object
    Set p = Parametres(): p("id") = id
    Set CommandeID = Appeler(operation, p)
End Function

Public Function DecoderReponse(ByVal statut As Long, ByVal octets As Variant) As Object
    Dim payload As String, description As String, reponse As Object, taille As Long
    On Error Resume Next
    taille = UBound(octets) + 1
    Err.Clear
    On Error GoTo 0
    If taille > 4000000 Then Err.Raise vbObjectError + 1102, , "Reponse trop volumineuse."
    If taille > 0 Then payload = TexteUTF8(octets)
    If statut < 200 Or statut >= 300 Then
        description = "Erreur HTTP " & CStr(statut) & ". "
        Select Case statut
            Case 401, 403: description = description & "Connexion ou droits du compte a verifier."
            Case 409: description = description & "Conflit : rechargez les donnees."
            Case 422: description = description & "Informations invalides ou incompletes."
            Case Else: description = description & "Service indisponible. Reprenez la meme action."
        End Select
        ' Un proxy peut renvoyer du HTML : son corps ne doit pas masquer le statut.
        On Error Resume Next
        Set reponse = modJson.JsonParse(payload)
        Err.Clear
        On Error GoTo 0
        If TypeName(reponse) = "Dictionary" Then
            If reponse.Exists("error") Then
                If VarType(reponse("error")) = vbString Then description = CStr(reponse("error"))
            End If
            If reponse.Exists("code") Then
                If VarType(reponse("code")) = vbString Then
                    If CStr(reponse("code")) = "destination_absente" Or CStr(reponse("code")) = "destination_ambigue" Then Err.Raise vbObjectError + 1141, "Service NAS", description
                End If
            End If
        End If
        Err.Raise vbObjectError + 1103, "Service NAS", description
    End If
    Set reponse = modJson.JsonParse(payload)
    If TypeName(reponse) <> "Dictionary" Then Err.Raise vbObjectError + 1105, , "Enveloppe de reponse invalide."
    If Not reponse.Exists("result") Then Err.Raise vbObjectError + 1105, , "Reponse incomplete."
    If TypeName(reponse("result")) <> "Dictionary" Then Err.Raise vbObjectError + 1105, , "Resultat invalide."
    Set DecoderReponse = reponse("result")
End Function
