Attribute VB_Name = "modJson"
Option Explicit
' =====================================================================
' modJson - JSON minimal et sur pour l'API Claude :
'  - JsonEchapper : echappe une chaine pour l'inclure dans un JSON
'  - JsonParse    : analyse un document JSON complet
'       objet  -> Scripting.Dictionary
'       tableau-> Collection
'       autres -> String / Double / Boolean / Null
'  - JsonTexteReponse / JsonErreurMessage : extraction reponse API
' =====================================================================

Private mTxt As String
Private mPos As Long
Private mLng As Long
Private mProfondeur As Long

Public Function JsonEchapper(ByVal s As String) As String
    Dim i As Long, c As String, code As Long, res As String
    ' Precomptage inutile : concatenation par blocs via Replace pour les cas frequents
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    ' caracteres de controle restants
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        code = AscW(c)
        If code >= 0 And code < 32 Then
            res = res & "\u" & Right$("0000" & Hex$(code), 4)
        Else
            res = res & c
        End If
    Next i
    JsonEchapper = res
End Function

Public Function JsonParse(ByVal texte As String) As Variant
    mTxt = texte: mPos = 1: mLng = Len(texte): mProfondeur = 0
    If mLng > 4000000 Then Err.Raise vbObjectError + 209, "modJson", "JSON trop volumineux."
    SauterBlancs
    If Mid$(mTxt, mPos, 1) = "{" Or Mid$(mTxt, mPos, 1) = "[" Then
        Set JsonParse = LireValeur()
    Else
        JsonParse = LireValeur()
    End If
    SauterBlancs
    If mPos <= mLng Then Err.Raise vbObjectError + 210, "modJson", "Donnees apres la fin du JSON."
    mTxt = ""
End Function

Private Function LireValeur() As Variant
    Dim c As String
    SauterBlancs
    If mPos > mLng Then Err.Raise vbObjectError + 200, "modJson", "JSON tronque"
    c = Mid$(mTxt, mPos, 1)
    mProfondeur = mProfondeur + 1
    If mProfondeur > 128 Then Err.Raise vbObjectError + 211, "modJson", "JSON trop imbrique."
    Select Case c
        Case "{": Set LireValeur = LireObjet()
        Case "[": Set LireValeur = LireTableau()
        Case """": LireValeur = LireChaine()
        Case "t"
            Attendre "true": LireValeur = True
        Case "f"
            Attendre "false": LireValeur = False
        Case "n"
            Attendre "null": LireValeur = Null
        Case Else
            LireValeur = LireNombre()
    End Select
    mProfondeur = mProfondeur - 1
End Function

Private Function LireObjet() As Object
    Dim dict As Object, cle As String, c As String
    Set dict = CreateObject("Scripting.Dictionary")
    mPos = mPos + 1                      ' {
    SauterBlancs
    If Mid$(mTxt, mPos, 1) = "}" Then mPos = mPos + 1: Set LireObjet = dict: Exit Function
    Do
        SauterBlancs
        cle = LireChaine()
        SauterBlancs
        If Mid$(mTxt, mPos, 1) <> ":" Then Err.Raise vbObjectError + 201, "modJson", "':' attendu pos " & mPos
        mPos = mPos + 1
        If dict.Exists(cle) Then Err.Raise vbObjectError + 212, "modJson", "Cle JSON dupliquee."
        SauterBlancs
        If Mid$(mTxt, mPos, 1) = "{" Or Mid$(mTxt, mPos, 1) = "[" Then
            Set dict(cle) = LireValeur()
        Else
            dict(cle) = LireValeur()
        End If
        SauterBlancs
        c = Mid$(mTxt, mPos, 1)
        mPos = mPos + 1
        If c = "}" Then Exit Do
        If c <> "," Then Err.Raise vbObjectError + 202, "modJson", "',' ou '}' attendu pos " & mPos
    Loop
    Set LireObjet = dict
End Function

Private Function LireTableau() As Collection
    Dim col As New Collection, c As String
    mPos = mPos + 1                      ' [
    SauterBlancs
    If Mid$(mTxt, mPos, 1) = "]" Then mPos = mPos + 1: Set LireTableau = col: Exit Function
    Do
        col.Add LireValeur()
        SauterBlancs
        c = Mid$(mTxt, mPos, 1)
        mPos = mPos + 1
        If c = "]" Then Exit Do
        If c <> "," Then Err.Raise vbObjectError + 203, "modJson", "',' ou ']' attendu pos " & mPos
    Loop
    Set LireTableau = col
End Function

Private Function LireChaine() As String
    Dim res As String, c As String, code As Long, debut As Long
    If Mid$(mTxt, mPos, 1) <> """" Then Err.Raise vbObjectError + 204, "modJson", "chaine attendue pos " & mPos
    mPos = mPos + 1
    debut = mPos
    Do While mPos <= mLng
        c = Mid$(mTxt, mPos, 1)
        If c = """" Then
            res = res & Mid$(mTxt, debut, mPos - debut)
            mPos = mPos + 1
            LireChaine = res
            Exit Function
        ElseIf c = "\" Then
            res = res & Mid$(mTxt, debut, mPos - debut)
            mPos = mPos + 1
            c = Mid$(mTxt, mPos, 1)
            Select Case c
                Case """": res = res & """"
                Case "\": res = res & "\"
                Case "/": res = res & "/"
                Case "b": res = res & Chr$(8)
                Case "f": res = res & Chr$(12)
                Case "n": res = res & vbLf
                Case "r": res = res & vbCr
                Case "t": res = res & vbTab
                Case "u"
                    Dim reHex As Object
                    Set reHex = CreateObject("VBScript.RegExp")
                    reHex.Pattern = "^[0-9A-Fa-f]{4}$"
                    If Not reHex.Test(Mid$(mTxt, mPos + 1, 4)) Then Err.Raise vbObjectError + 213, "modJson", "Sequence unicode invalide."
                    code = CLng("&H" & Mid$(mTxt, mPos + 1, 4))
                    res = res & ChrW$(code)
                    mPos = mPos + 4
                Case Else
                    Err.Raise vbObjectError + 205, "modJson", "echappement inconnu \" & c
            End Select
            mPos = mPos + 1
            debut = mPos
        Else
            If AscW(c) >= 0 And AscW(c) < 32 Then Err.Raise vbObjectError + 214, "modJson", "Caractere de controle non echappe."
            mPos = mPos + 1
        End If
    Loop
    Err.Raise vbObjectError + 206, "modJson", "chaine non terminee"
End Function

Private Function LireNombre() As Double
    Dim debut As Long, c As String, s As String
    debut = mPos
    Do While mPos <= mLng
        c = Mid$(mTxt, mPos, 1)
        If InStr("0123456789+-.eE", c) > 0 Then mPos = mPos + 1 Else Exit Do
    Loop
    s = Mid$(mTxt, debut, mPos - debut)
    If Len(s) = 0 Then Err.Raise vbObjectError + 207, "modJson", "nombre attendu pos " & debut
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^-?(0|[1-9][0-9]*)(\.[0-9]+)?([eE][+-]?[0-9]+)?$"
    If Not re.Test(s) Then Err.Raise vbObjectError + 215, "modJson", "Nombre JSON invalide."
    LireNombre = Val(s)                  ' Val ignore la locale (point decimal)
End Function

Private Sub Attendre(ByVal mot As String)
    If Mid$(mTxt, mPos, Len(mot)) <> mot Then
        Err.Raise vbObjectError + 208, "modJson", "'" & mot & "' attendu pos " & mPos
    End If
    mPos = mPos + Len(mot)
End Sub

Private Sub SauterBlancs()
    Dim c As String
    Do While mPos <= mLng
        c = Mid$(mTxt, mPos, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then mPos = mPos + 1 Else Exit Do
    Loop
End Sub

' --- Extraction specifique reponse API Claude ------------------------

' Concatene les blocs "text" de content[]. Renvoie "" si absent.
Public Function JsonTexteReponse(ByVal racine As Variant) As String
    On Error GoTo Absent
    Dim bloc As Variant, res As String
    For Each bloc In racine("content")
        If bloc("type") = "text" Then res = res & bloc("text")
    Next bloc
    JsonTexteReponse = res
    Exit Function
Absent:
    JsonTexteReponse = ""
End Function

' Reponse OpenAI : Responses API (output[].content[] type output_text) ou,
' a defaut, Chat Completions (choices[0].message.content). "" si absent.
Public Function JsonTexteReponseOpenAI(ByVal racine As Variant) As String
    Dim item As Variant, bloc As Variant, res As String
    If Not IsObject(racine) Then Err.Raise vbObjectError + 216, "modJson", "Reponse API inattendue."
    If Not racine.Exists("status") Then Err.Raise vbObjectError + 217, "modJson", "Statut API absent."
    If CStr(racine("status")) <> "completed" Then Err.Raise vbObjectError + 218, "modJson", "Reponse API incomplete ou interrompue."
    If Not racine.Exists("output") Then Err.Raise vbObjectError + 219, "modJson", "Sortie API absente."
    For Each item In racine("output")
        If item("type") = "message" Then
            For Each bloc In item("content")
                If bloc("type") = "refusal" Then Err.Raise vbObjectError + 220, "modJson", "Reponse API refusee."
                If bloc("type") = "output_text" Then
                    If Len(res) > 0 Then res = res & vbCrLf
                    res = res & CStr(bloc("text"))
                End If
            Next bloc
        End If
    Next item
    JsonTexteReponseOpenAI = res
End Function
Public Function JsonErreurMessage(ByVal racine As Variant) As String
    On Error GoTo Absent
    JsonErreurMessage = racine("error")("message")
    Exit Function
Absent:
    JsonErreurMessage = ""
End Function
