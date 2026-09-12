Attribute VB_Name = "modServiceNas"
Option Explicit
Private mEnAttente As Object

Public Function Parametres() As Object
    Set Parametres = CreateObject("Scripting.Dictionary")
    Parametres.CompareMode = 1
End Function

Public Function JsonValeur(ByVal valeur As Variant) As String
    Dim k As Variant, element As Variant, texte As String
    If IsObject(valeur) Then
        If TypeName(valeur) = "Collection" Then
            For Each element In valeur
                If Len(texte) > 0 Then texte = texte & ","
                texte = texte & JsonValeur(element)
            Next element
            JsonValeur = "[" & texte & "]"
        Else
            For Each k In valeur.Keys
                If Len(texte) > 0 Then texte = texte & ","
                texte = texte & Chr$(34) & modJson.JsonEchapper(CStr(k)) & Chr$(34) & ":" & JsonValeur(valeur(k))
            Next k
            JsonValeur = "{" & texte & "}"
        End If
    ElseIf IsNull(valeur) Then
        JsonValeur = "null"
    ElseIf VarType(valeur) = vbBoolean Then
        If valeur Then JsonValeur = "true" Else JsonValeur = "false"
    ElseIf IsNumeric(valeur) And VarType(valeur) <> vbString Then
        JsonValeur = Replace(CStr(valeur), ",", ".")
    Else
        JsonValeur = Chr$(34) & modJson.JsonEchapper(CStr(valeur)) & Chr$(34)
    End If
End Function

Public Function Appeler(ByVal operation As String, ByVal params As Object) As Object
    Dim requete As Object, reponse As Object, http As Object, flux As Object
    Dim base As String, jeton As String, payload As String, cle As String, octets As Variant
    base = Trim$(modFichiers.LireTexteUTF8(Environ$("APPDATA") & "\CabinetCardio\service.url"))
    If LCase$(Left$(base, 8)) <> "https://" Then Err.Raise vbObjectError + 1100, "Service NAS", "Adresse HTTPS du service absente ou invalide."
    Do While Right$(base, 1) = "/": base = Left$(base, Len(base) - 1): Loop
    jeton = Trim$(modFichiers.LireTexteUTF8(Environ$("APPDATA") & "\CabinetCardio\service.token"))
    If Len(jeton) < 32 Or InStr(jeton, vbCr) Or InStr(jeton, vbLf) Then Err.Raise vbObjectError + 1101, "Service NAS", "Identifiant de connexion du poste absent ou invalide."
    If mEnAttente Is Nothing Then Set mEnAttente = CreateObject("Scripting.Dictionary")
    cle = operation & "|" & JsonValeur(params)
    If Not mEnAttente.Exists(cle) Then mEnAttente(cle) = modFichiers.IdUnique()
    Set requete = Parametres()
    requete("operation") = operation
    Set requete("params") = params
    requete("request_id") = CStr(mEnAttente(cle))
    payload = JsonValeur(requete)
    Set flux = CreateObject("ADODB.Stream")
    flux.Type = 2: flux.Charset = "utf-8": flux.Open
    flux.WriteText payload: flux.Position = 0: flux.Type = 1: flux.Position = 3
    octets = flux.Read: flux.Close
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 5000, 10000, 30000, 30000
    http.Open "POST", base & "/v1/rpc", False
    http.Option(6) = False ' Aucun renvoi du jeton vers une redirection HTTP.
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Authorization", "Bearer " & jeton
    http.Send octets
    If Len(http.ResponseText) > 4000000 Then Err.Raise vbObjectError + 1102, "Service NAS", "Reponse trop volumineuse."
    Set reponse = modJson.JsonParse(http.ResponseText)
    If http.Status < 200 Or http.Status >= 300 Then
        If http.Status < 500 Then mEnAttente.Remove cle
        If reponse.Exists("error") Then Err.Raise vbObjectError + 1103, "Service NAS", CStr(reponse("error"))
        Err.Raise vbObjectError + 1104, "Service NAS", "Erreur HTTP " & http.Status
    End If
    If Not reponse.Exists("result") Then Err.Raise vbObjectError + 1105, "Service NAS", "Reponse incomplete."
    Set Appeler = reponse("result")
    mEnAttente.Remove cle
End Function

Public Function LireTable(ByVal genre As String, Optional ByVal recherche As String = "", _
                          Optional ByVal annee As String = "", Optional ByVal jour As String = "") As Collection
    Dim p As Object, r As Object, item As Object, resultat As New Collection
    Set p = Parametres(): p("genre") = genre: p("q") = recherche: p("year") = annee: p("date") = jour: p("offset") = 0
    Do
        Set r = Appeler("table.read", p)
        For Each item In r("items"): resultat.Add item: Next item
        If IsNull(r("next")) Then Exit Do
        p("offset") = CLng(r("next"))
    Loop
    Set LireTable = resultat
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
