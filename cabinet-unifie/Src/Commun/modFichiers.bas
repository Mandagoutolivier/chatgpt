Attribute VB_Name = "modFichiers"
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef value As GUID) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef value As GUID, ByVal buffer As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal source As LongPtr, ByVal destination As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CoCreateGuid Lib "ole32" (ByRef value As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (ByRef value As GUID, ByVal buffer As Long, ByVal size As Long) As Long
Private Declare Function MoveFileExW Lib "kernel32" (ByVal source As Long, ByVal destination As Long, ByVal flags As Long) As Long
#End If
Private mVerrous As Object

Public Sub Pause(ByVal ms As Long)
    If ms > 0 Then Sleep ms
End Sub

Public Function FichierExiste(ByVal chemin As String) As Boolean
    FichierExiste = CreateObject("Scripting.FileSystemObject").FileExists(chemin)
End Function

Public Function DossierExiste(ByVal chemin As String) As Boolean
    DossierExiste = CreateObject("Scripting.FileSystemObject").FolderExists(chemin)
End Function

Public Sub EnsureDossier(ByVal chemin As String)
    Dim fso As Object, parent As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(chemin) Then Exit Sub
    parent = fso.GetParentFolderName(chemin)
    If Len(parent) = 0 Or parent = chemin Then Err.Raise vbObjectError + 110, "modFichiers", "Dossier inaccessible : " & chemin
    EnsureDossier parent
    ' La creation concurrente du meme dossier est acceptable.
    On Error GoTo Echec
    fso.CreateFolder chemin
    Exit Sub
Echec:
    If Not fso.FolderExists(chemin) Then Err.Raise vbObjectError + 111, "modFichiers", "Creation impossible : " & chemin
End Sub

Public Function ListerFichiers(ByVal dossier As String, ByVal extension As String) As Collection
    Dim fso As Object, f As Object, noms As New Collection, resultat As New Collection, n As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dossier) Then Err.Raise vbObjectError + 112, "modFichiers", "Dossier inaccessible : " & dossier
    For Each f In fso.GetFolder(dossier).Files
        If LCase$(Right$(f.Name, Len(extension))) = LCase$(extension) Then InsererTrie noms, f.Name
    Next f
    For Each n In noms
        resultat.Add dossier & "\" & n
    Next n
    Set ListerFichiers = resultat
End Function

Private Sub InsererTrie(ByVal col As Collection, ByVal valeur As String)
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(col(i), valeur, vbTextCompare) > 0 Then
            col.Add valeur, , i
            Exit Sub
        End If
    Next i
    col.Add valeur
End Sub

Public Function LireTexteUTF8(ByVal chemin As String) As String
    Dim st As Object, n As Long, description As String
    On Error GoTo Echec
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile chemin
    LireTexteUTF8 = st.ReadText(-1)
    st.Close
    If Len(LireTexteUTF8) > 0 Then
        If AscW(Left$(LireTexteUTF8, 1)) = -257 Then LireTexteUTF8 = Mid$(LireTexteUTF8, 2)
    End If
    Exit Function
Echec:
    n = Err.Number: description = Err.Description
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    On Error GoTo 0
    Err.Raise n, "modFichiers.LireTexteUTF8", description
End Function

Public Sub EcrireTexteUTF8(ByVal chemin As String, ByVal contenu As String)
    EcrireAvecEncodage chemin, contenu, "utf-8"
End Sub

Public Sub EcrireTexteAnsi(ByVal chemin As String, ByVal contenu As String)
    EcrireAvecEncodage chemin, contenu, "windows-1252"
End Sub

Private Sub EcrireAvecEncodage(ByVal chemin As String, ByVal contenu As String, ByVal charset As String)
    Dim st As Object, n As Long, description As String
    On Error GoTo Echec
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = charset: st.Open
    st.WriteText contenu
    st.SaveToFile chemin, 2
    st.Close
    Exit Sub
Echec:
    n = Err.Number: description = Err.Description
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    On Error GoTo 0
    Err.Raise n, "modFichiers.EcrireAvecEncodage", description
End Sub

' Source et destination doivent etre sur le meme volume (temporaire voisin).
Public Sub RenommerAtomique(ByVal source As String, ByVal destination As String, Optional ByVal remplacer As Boolean = False)
    Dim flags As Long, code As Long
    flags = 8   ' MOVEFILE_WRITE_THROUGH
    If remplacer Then flags = flags Or 1
    If MoveFileExW(StrPtr(source), StrPtr(destination), flags) = 0 Then
        code = Err.LastDllError
        Err.Raise vbObjectError + 113, "modFichiers", "Publication impossible (Windows " & code & ") : " & destination
    End If
End Sub

Public Function DossierTempLocal() As String
    DossierTempLocal = Environ$("LOCALAPPDATA") & "\CabinetCardio\Temp"
    EnsureDossier DossierTempLocal
End Function

' Copie ephemere unique. L'appelant la supprime apres fermeture du classeur.
Public Function CopieLocale(ByVal chemin As String, Optional ByVal essais As Long = 4, Optional ByVal delaiMs As Long = 500) As String
    Dim fso As Object, dest As String, verrou As String, n As Long, description As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    verrou = fso.GetBaseName(chemin)
    If Not AcquerirVerrou(verrou, essais * delaiMs + 2000) Then Err.Raise vbObjectError + 114, "modFichiers", "Base occupee : " & chemin
    On Error GoTo Echec
    dest = DossierTempLocal() & "\" & IdUnique() & "_" & fso.GetFileName(chemin)
    fso.CopyFile chemin, dest, False
    RelacherVerrou verrou
    CopieLocale = dest
    Exit Function
Echec:
    n = Err.Number: description = Err.Description
    RelacherVerrou verrou
    SupprimerTemporaire dest
    Err.Raise n, "modFichiers.CopieLocale", description
End Function

Public Sub SupprimerTemporaire(ByVal chemin As String)
    On Error Resume Next
    If Len(chemin) > 0 Then CreateObject("Scripting.FileSystemObject").DeleteFile chemin, True
End Sub

Private Function CheminVerrou(ByVal nom As String) As String
    VerifierIdentifiantFichier nom
    EnsureDossier modConfig.Chemin("Base") & "\locks"
    CheminVerrou = modConfig.Chemin("Base") & "\locks\" & nom & ".lock"
End Function

' Verrou de fichier SMB conserve OUVERT : libere par Windows apres un crash.
' Le fichier .lock persiste ; son anciennete ne permet jamais de voler un verrou.
Public Function AcquerirVerrou(ByVal nom As String, Optional ByVal timeoutMs As Long = 5000) As Boolean
    Dim chemin As String, numero As Integer, debut As Double, ecoule As Double, erreur As Long
    If mVerrous Is Nothing Then
        Set mVerrous = CreateObject("Scripting.Dictionary")
        mVerrous.CompareMode = 1
    End If
    If mVerrous.Exists(nom) Then Exit Function
    chemin = CheminVerrou(nom)
    debut = Timer
    Do
        numero = FreeFile
        On Error Resume Next
        Err.Clear
        Open chemin For Binary Access Read Write Lock Read Write As #numero
        erreur = Err.Number
        On Error GoTo 0
        If erreur = 0 Then
            mVerrous.Add nom, numero
            AcquerirVerrou = True
            Exit Function
        End If
        ecoule = Timer - debut
        If ecoule < 0 Then ecoule = ecoule + 86400#
        If ecoule * 1000# >= timeoutMs Then Exit Do
        Pause 100
    Loop
End Function

Public Sub RelacherVerrou(ByVal nom As String)
    Dim numero As Integer
    If mVerrous Is Nothing Then Exit Sub
    If Not mVerrous.Exists(nom) Then Exit Sub
    numero = CInt(mVerrous(nom))
    Close #numero
    mVerrous.Remove nom
End Sub

Public Sub VerifierIdentifiantFichier(ByVal valeur As String)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    If NomWindowsReserve(valeur) Then Err.Raise vbObjectError + 122, "modFichiers", "Nom reserve par Windows."
    re.Pattern = "^[A-Za-z0-9_-][A-Za-z0-9_ .-]{0,150}$"
    If Not re.Test(valeur) Or Right$(valeur, 1) = "." Or Right$(valeur, 1) = " " Then
        Err.Raise vbObjectError + 115, "modFichiers", "Identifiant de fichier invalide."
    End If
End Sub

Public Function EcrireDrapeau(ByVal dossier As String, ByVal nomBase As String, ByVal dict As Object) As String
    Dim contenu As String, cle As Variant, tmp As String, final As String, n As Long, description As String
    VerifierIdentifiantFichier nomBase
    EnsureDossier dossier
    For Each cle In dict.Keys
        If InStr(CStr(cle), "=") > 0 Or InStr(CStr(cle), vbLf) > 0 Or InStr(CStr(cle), vbCr) > 0 Then Err.Raise vbObjectError + 116, "modFichiers", "Cle de drapeau invalide."
        contenu = contenu & cle & "=" & Replace(Replace(CStr(dict(cle)), vbCr, " "), vbLf, " ") & vbCrLf
    Next cle
    final = dossier & "\" & nomBase & ".txt"
    If FichierExiste(final) Then
        If LireTexteUTF8(final) <> contenu Then Err.Raise vbObjectError + 117, "modFichiers", "Un evenement different utilise deja cet identifiant."
        EcrireDrapeau = final
        Exit Function
    End If
    tmp = dossier & "\" & IdUnique() & ".tmp"
    On Error GoTo Echec
    EcrireTexteUTF8 tmp, contenu
    RenommerAtomique tmp, final
    EcrireDrapeau = final
    Exit Function
Echec:
    n = Err.Number: description = Err.Description
    SupprimerTemporaire tmp
    Err.Raise n, "modFichiers.EcrireDrapeau", description
End Function

Public Function LireDrapeau(ByVal chemin As String) As Object
    Dim dict As Object, lignes() As String, i As Long, pos As Long, cle As String
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    lignes = Split(Replace(LireTexteUTF8(chemin), vbCrLf, vbLf), vbLf)
    For i = LBound(lignes) To UBound(lignes)
        pos = InStr(lignes(i), "=")
        If pos > 1 Then
            cle = Left$(lignes(i), pos - 1)
            If dict.Exists(cle) Then Err.Raise vbObjectError + 118, "modFichiers", "Cle dupliquee dans le drapeau : " & cle
            dict.Add cle, Mid$(lignes(i), pos + 1)
        End If
    Next i
    Set LireDrapeau = dict
End Function

Public Function IdUnique() As String
    Dim g As GUID, buffer As String
    If CoCreateGuid(g) <> 0 Then Err.Raise vbObjectError + 119, "modFichiers", "Creation GUID impossible."
    buffer = String$(39, vbNullChar)
    If StringFromGUID2(g, StrPtr(buffer), 39) = 0 Then Err.Raise vbObjectError + 120, "modFichiers", "Conversion GUID impossible."
    IdUnique = LCase$(Mid$(buffer, 2, 36))
End Function

Public Sub SauvegardeHorodatee(ByVal chemin As String, Optional ByVal maxVersions As Long = 60)
    Dim fso As Object, dossier As String, prefixe As String, dest As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(chemin) Then Err.Raise vbObjectError + 121, "modFichiers", "Base a sauvegarder absente : " & chemin
    dossier = modConfig.Chemin("Sauvegardes")
    EnsureDossier dossier
    prefixe = fso.GetBaseName(chemin) & "_"
    dest = dossier & "\" & prefixe & Format$(Now, "yyyymmdd_hhnnss") & "_" & IdUnique() & "." & fso.GetExtensionName(chemin)
    fso.CopyFile chemin, dest, False
    ' Pas de purge automatique : retention a regler sur le Synology.
End Sub

Public Function NomFichierSur(ByVal s As String) As String
    Dim c As Variant, i As Long
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, CStr(c), " ")
    Next c
    For i = 0 To 31
        s = Replace(s, Chr$(i), " ")
    Next i
    s = Left$(Trim$(s), 120)
    Do While Len(s) > 0
        If Right$(s, 1) <> "." And Right$(s, 1) <> " " Then Exit Do
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then s = "document"
    If NomWindowsReserve(s) Then s = "_" & s
    NomFichierSur = s
End Function

' Preserve les zeros initiaux et traite les saisies comme du texte, jamais comme des formules.
Public Sub EcrireCelluleTexte(ByVal cellule As Object, ByVal valeur As String)
    cellule.NumberFormat = "@"
    cellule.Value2 = "'" & valeur
End Sub

Private Function NomWindowsReserve(ByVal valeur As String) As Boolean
    Dim nom As String
    nom = UCase$(Split(valeur & ".", ".")(0))
    NomWindowsReserve = (nom = "CON" Or nom = "PRN" Or nom = "AUX" Or nom = "NUL" Or nom Like "COM[1-9]" Or nom Like "LPT[1-9]")
End Function
