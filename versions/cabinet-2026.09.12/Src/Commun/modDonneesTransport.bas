Attribute VB_Name = "modDonneesTransport"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32.dll" (ByRef provider As LongPtr, ByVal container As LongPtr, ByVal providerName As LongPtr, ByVal providerType As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" (ByVal provider As LongPtr, ByVal algorithm As Long, ByVal key As LongPtr, ByVal flags As Long, ByRef hash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" (ByVal hash As LongPtr, ByRef data As Any, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" (ByVal hash As LongPtr, ByVal param As Long, ByRef data As Any, ByRef length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" (ByVal hash As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal provider As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function CryptAcquireContextW Lib "advapi32.dll" (ByRef provider As Long, ByVal container As Long, ByVal providerName As Long, ByVal providerType As Long, ByVal flags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal provider As Long, ByVal algorithm As Long, ByVal key As Long, ByVal flags As Long, ByRef hash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hash As Long, ByRef data As Any, ByVal length As Long, ByVal flags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hash As Long, ByVal param As Long, ByRef data As Any, ByRef length As Long, ByVal flags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hash As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal provider As Long, ByVal flags As Long) As Long
#End If

Public Function SerialiserJson(ByVal valeur As Variant) As String
    Dim k As Variant, element As Variant, texte As String, keys As Variant, i As Long, j As Long, swap As Variant
    If IsObject(valeur) Then
        Select Case TypeName(valeur)
        Case "Collection"
            For Each element In valeur
                If Len(texte) > 0 Then texte = texte & ","
                texte = texte & SerialiserJson(element)
            Next element
            SerialiserJson = "[" & texte & "]"
        Case "Dictionary"
            keys = valeur.Keys
            For i = 0 To valeur.Count - 2
                For j = i + 1 To valeur.Count - 1
                    If StrComp(CStr(keys(i)), CStr(keys(j)), vbBinaryCompare) > 0 Then
                        swap = keys(i): keys(i) = keys(j): keys(j) = swap
                    End If
                Next j
            Next i
            For Each k In keys
                If Len(texte) > 0 Then texte = texte & ","
                texte = texte & SerialiserJson(CStr(k)) & ":" & SerialiserJson(valeur(k))
            Next k
            SerialiserJson = "{" & texte & "}"
        Case Else
            Err.Raise vbObjectError + 1106, , "Type objet JSON non pris en charge."
        End Select
    Else
        Select Case VarType(valeur)
        Case vbNull, vbEmpty: SerialiserJson = "null"
        Case vbBoolean
            If valeur Then SerialiserJson = "true" Else SerialiserJson = "false"
        Case vbDate
            SerialiserJson = Chr$(34) & Format$(CDate(valeur), "yyyy-mm-dd") & "T" & Format$(CDate(valeur), "hh:nn:ss") & Chr$(34)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            SerialiserJson = Replace(CStr(valeur), ",", ".")
        Case vbString
            SerialiserJson = Chr$(34) & modJson.JsonEchapper(CStr(valeur)) & Chr$(34)
        Case Else
            Err.Raise vbObjectError + 1106, , "Type JSON non pris en charge."
        End Select
    End If
End Function


Public Function EncoderUTF8(ByVal texte As String) As Variant
    Dim flux As Object
    Set flux = CreateObject("ADODB.Stream")
    flux.Type = 2: flux.Charset = "utf-8": flux.Open
    flux.WriteText texte: flux.Position = 0: flux.Type = 1: flux.Position = 3
    EncoderUTF8 = flux.Read: flux.Close
End Function


Public Function DecoderUTF8(ByVal octets As Variant) As String
    Dim flux As Object
    Set flux = CreateObject("ADODB.Stream")
    flux.Type = 1: flux.Open: flux.Write octets
    flux.Position = 0: flux.Type = 2: flux.Charset = "utf-8"
    DecoderUTF8 = flux.ReadText: flux.Close
End Function


Public Function EmpreinteSHA256(ByVal texte As String) As String
#If VBA7 Then
    Dim provider As LongPtr, hash As LongPtr
#Else
    Dim provider As Long, hash As Long
#End If
    Dim bytes() As Byte, digest(0 To 31) As Byte, length As Long, i As Long, numero As Long
    Dim etape As String, erreurWindows As Long
    On Error GoTo Echec
    etape = "utf8": If Len(texte) > 0 Then bytes = EncoderUTF8(texte)
    etape = "contexte"
    If CryptAcquireContextW(provider, 0, 0, 24, &HF0000000) = 0 Then Err.Raise 5
    etape = "initialisation"
    If CryptCreateHash(provider, &H800C&, 0, 0, hash) = 0 Then Err.Raise 5
    If Len(texte) > 0 Then
        etape = "donnees"
        If CryptHashData(hash, bytes(0), UBound(bytes) + 1, 0) = 0 Then Err.Raise 5
    End If
    etape = "resultat": length = 32
    If CryptGetHashParam(hash, 2, digest(0), length, 0) = 0 Then Err.Raise 5
    For i = 0 To 31: EmpreinteSHA256 = EmpreinteSHA256 & LCase$(Right$("0" & Hex$(digest(i)), 2)): Next i
Sortie:
    If hash <> 0 Then CryptDestroyHash hash
    If provider <> 0 Then CryptReleaseContext provider, 0
    On Error GoTo 0
    If numero <> 0 Then Err.Raise vbObjectError + 1107, "EmpreinteSHA256", "Calcul de l empreinte impossible : " & etape & " (" & CStr(numero) & ", Windows " & CStr(erreurWindows) & ")."
    Exit Function
Echec:
    numero = Err.Number: erreurWindows = Err.LastDllError: Resume Sortie
End Function
