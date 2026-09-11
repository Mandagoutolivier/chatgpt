Attribute VB_Name = "modBase"
Option Explicit
' =====================================================================
' modBase - Lecture des bases (poste medecin) SANS verrou :
' copie locale du xlsx puis lecture par une instance Excel INVISIBLE
' (l'ADO/ACE se fige a l'interieur de Word : ne pas y revenir).
' Renvoie des Collections de Scripting.Dictionary (cle = nom de colonne,
' insensible a la casse). Cache memoire de 60 s.
' =====================================================================

Private mPatients As Collection
Private mCorrespondants As Collection
Private mCharge As Date
Private Const CACHE_S As Long = 60

' Lit une feuille d'un classeur (copie locale) et renvoie la Collection.
Public Function LireTable(ByVal fichier As String, ByVal feuille As String, _
                          Optional ByVal colonneNonVide As String = "ID") As Collection
    Dim xl As Object, wb As Object, res As Collection, copie As String
    Set xl = OuvrirExcel()
    On Error GoTo Nettoyage
    copie = modFichiers.CopieLocale(fichier)
    Set wb = xl.Workbooks.Open(copie, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Set res = LireFeuille(wb, feuille, colonneNonVide)
    wb.Close False
    xl.Quit
    modFichiers.SupprimerTemporaire copie
    Set LireTable = res
    Exit Function
Nettoyage:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    xl.Quit
    modFichiers.SupprimerTemporaire copie
    On Error GoTo 0
    Err.Raise numErr, "modBase", descErr
End Function

Private Function OuvrirExcel() As Object
    Dim xl As Object
    Set xl = CreateObject("Excel.Application")
    xl.visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    xl.EnableEvents = False
    Set OuvrirExcel = xl
End Function

Private Function LireFeuille(ByVal wb As Object, ByVal feuille As String, _
                             ByVal colonneNonVide As String) As Collection
    Dim ws As Object, donnees As Variant, col As Collection, dict As Object
    Dim r As Long, c As Long, nLignes As Long, nCols As Long
    Set ws = wb.Worksheets(feuille)
    Set col = New Collection
    donnees = ws.UsedRange.Value
    If IsEmpty(donnees) Or Not IsArray(donnees) Then
        Set LireFeuille = col
        Exit Function
    End If
    nLignes = UBound(donnees, 1)
    nCols = UBound(donnees, 2)
    For r = 2 To nLignes
        Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = 1
        For c = 1 To nCols
            dict(Nz(donnees(1, c))) = Nz(donnees(r, c))
        Next c
        If Len(colonneNonVide) = 0 Then
            col.Add dict
        ElseIf dict.Exists(colonneNonVide) Then
            If Len(dict(colonneNonVide)) > 0 Then col.Add dict
        End If
    Next r
    Set LireFeuille = col
End Function

Private Function Nz(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        Nz = ""
    ElseIf VarType(v) = vbDate Then
        Nz = Format$(v, "dd/mm/yyyy")
    Else
        Nz = Trim$(CStr(v))
    End If
End Function

' ---------------------------------------------------------------------
' Installation neuve : la base n'existe pas encore. On la cree VIDE avec
' ses colonnes plutot que d'echouer sur "Fichier introuvable" - le
' medecin peut alors creer son premier patient depuis le poste secretaire.
' xl : instance Excel invisible deja ouverte par l'appelant.
' ---------------------------------------------------------------------
Private Sub CreerBaseSiAbsente(ByVal xl As Object, ByVal fichier As String, ByVal feuilles As Variant)
    If Not modFichiers.FichierExiste(fichier) Then Err.Raise vbObjectError + 200, "modBase", _
        "Base NAS absente. Initialisez le NAS depuis l installateur du secretariat : " & fichier
End Sub
' Charge PATIENTS et CORRESPONDANTS en UNE session Excel
Private Sub ChargerBases(ByVal forcer As Boolean)
    Dim xl As Object, wb As Object, tous As Collection, actifs As Collection, c As Object
    If Not forcer And Not mPatients Is Nothing Then
        If DateDiff("s", mCharge, Now) <= CACHE_S Then Exit Sub
    End If
    modLog.Etape "base : ouverture d'Excel invisible"
    Set xl = OuvrirExcel()
    On Error GoTo Nettoyage
    modLog.Etape "base : verification de Patients.xlsx"
    CreerBaseSiAbsente xl, modConfig.FichierPatients(), modSchemas.FeuillesBasePatients()
    modLog.Etape "base : copie locale de Patients.xlsx"
    Dim copie As String
    copie = modFichiers.CopieLocale(modConfig.FichierPatients())
    modLog.Etape "base : ouverture de la copie " & copie
    Set wb = xl.Workbooks.Open(copie, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    modLog.Etape "base : lecture des feuilles"
    Set mPatients = LireFeuille(wb, "PATIENTS", "ID")
    Set tous = LireFeuille(wb, "CORRESPONDANTS", "ID")
    wb.Close False
    xl.Quit
    modFichiers.SupprimerTemporaire copie
    modLog.Etape "base : " & mPatients.Count & " patients charges"
    Set actifs = New Collection
    For Each c In tous
        If c("Actif") <> "0" Then actifs.Add c
    Next c
    Set mCorrespondants = actifs
    mCharge = Now
    Exit Sub
Nettoyage:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    xl.Quit
    modFichiers.SupprimerTemporaire copie
    On Error GoTo 0
    Err.Raise numErr, "modBase", descErr
End Sub

Public Function patients(Optional ByVal forcer As Boolean = False) As Collection
    ChargerBases forcer
    Set patients = mPatients
End Function

Public Function Correspondants(Optional ByVal forcer As Boolean = False) As Collection
    ChargerBases forcer
    Set Correspondants = mCorrespondants
End Function

Public Function PatientParID(ByVal id As String, Optional ByVal forcer As Boolean = False) As Object
    Dim p As Object
    For Each p In patients(forcer)
        If p("ID") = id Then Set PatientParID = p: Exit Function
    Next p
    Set PatientParID = Nothing
End Function

Public Function CorrespondantParID(ByVal id As String) As Object
    Dim c As Object
    For Each c In Correspondants()
        If c("ID") = id Then Set CorrespondantParID = c: Exit Function
    Next c
    Set CorrespondantParID = Nothing
End Function
