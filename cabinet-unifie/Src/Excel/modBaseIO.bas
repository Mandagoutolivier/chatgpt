Attribute VB_Name = "modBaseIO"
Option Explicit
' =====================================================================
' modBaseIO - Ecritures dans les bases (poste secretaire UNIQUEMENT).
' Toute ecriture est une TRANSACTION COURTE : verrou cooperatif ->
' sauvegarde (1re ecriture du jour) -> ouvrir -> ecrire -> fermer ->
' relacher. Les fichiers ne restent JAMAIS ouverts.
' =====================================================================

Private mEcran As Boolean
Private mEvenements As Boolean
Private mTransaction As Boolean
Private mJourSauvegarde As String   ' "yyyymmdd|fichier" deja sauvegardes

' --- lecture (copie locale, fenetre cachee) ---------------------------

Public Function LireTableX(ByVal fichier As String, ByVal Feuille As String, Optional ByVal colonneNonVide As String = "ID") As Collection
    Dim wb As Workbook, copie As String, ecran As Boolean, events As Boolean
    Dim numero As Long, description As String, securite As Long
    ecran = Application.ScreenUpdating: events = Application.EnableEvents: securite = Application.AutomationSecurity
    On Error GoTo Echec
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.AutomationSecurity = 3
    copie = modFichiers.CopieLocale(fichier)
    Set wb = Workbooks.Open(copie, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    wb.Windows(1).Visible = False
    Set LireTableX = LireFeuilleX(wb, Feuille, colonneNonVide)
    wb.Close False
    Set wb = Nothing
Sortie:
    Application.ScreenUpdating = ecran: Application.EnableEvents = events: Application.AutomationSecurity = securite
    modFichiers.SupprimerTemporaire copie
    If numero <> 0 Then Err.Raise numero, "modBaseIO.LireTableX", description
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    GoTo Sortie
End Function
Private Function LireFeuilleX(ByVal wb As Workbook, ByVal Feuille As String, _
                              ByVal colonneNonVide As String) As Collection
    Dim ws As Worksheet, donnees As Variant, col As Collection, dict As Object
    Dim r As Long, c As Long
    Set ws = wb.Worksheets(Feuille)
    Set col = New Collection
    donnees = ws.UsedRange.Value
    If IsEmpty(donnees) Or Not IsArray(donnees) Then Set LireFeuilleX = col: Exit Function
    For r = 2 To UBound(donnees, 1)
        Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = 1
        For c = 1 To UBound(donnees, 2)
            dict(Nz(donnees(1, c))) = Nz(donnees(r, c))
        Next c
        If Len(colonneNonVide) = 0 Then
            col.Add dict
        ElseIf dict.Exists(colonneNonVide) Then
            If Len(dict(colonneNonVide)) > 0 Then col.Add dict
        End If
    Next r
    Set LireFeuilleX = col
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

' --- transactions d'ecriture -----------------------------------------

' Ajoute une ligne. Si prefixeID est fourni ("P", "C", "R"...), genere
' l'identifiant (colonne ID) et le renvoie.
Public Function AjouterLigne(ByVal fichier As String, ByVal Feuille As String, _
                             ByVal valeurs As Object, Optional ByVal prefixeID As String = "") As String
    Dim wb As Workbook, ws As Worksheet, ligne As Long, id As String
    OuvrirTransaction fichier, wb
    On Error GoTo Echec
    Set ws = wb.Worksheets(Feuille)
    If StrComp(Feuille, "RDV", vbTextCompare) = 0 Then VerifierCreneauSousVerrou ws, valeurs
    ligne = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1      ' xlUp
    If ligne < 2 Then ligne = 2
    If Len(prefixeID) > 0 Then
        id = ProchainID(ws, prefixeID)
        valeurs("ID") = id
    End If
    EcrireValeurs ws, ligne, valeurs
    FermerTransaction fichier, wb, True
    AjouterLigne = id
    Exit Function
Echec:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    FermerTransaction fichier, wb, False
    Err.Raise numErr, "modBaseIO", descErr
End Function

' Modifie la premiere ligne ou colonneCle = valeurCle.
Public Sub ModifierLigne(ByVal fichier As String, ByVal Feuille As String, _
                         ByVal colonneCle As String, ByVal valeurCle As String, _
                         ByVal valeurs As Object)
    Dim wb As Workbook, ws As Worksheet, ligne As Long
    OuvrirTransaction fichier, wb
    On Error GoTo Echec
    Set ws = wb.Worksheets(Feuille)
    ligne = TrouverLigne(ws, colonneCle, valeurCle)
    If ligne = 0 Then Err.Raise vbObjectError + 600, "modBaseIO", _
        "Introuvable : " & colonneCle & "=" & valeurCle & " dans " & Feuille
    EcrireValeurs ws, ligne, valeurs
    FermerTransaction fichier, wb, True
    Exit Sub
Echec:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    FermerTransaction fichier, wb, False
    Err.Raise numErr, "modBaseIO", descErr
End Sub

' Ajoute plusieurs lignes d'un coup (journal comptable)
Public Sub AjouterLignes(ByVal fichier As String, ByVal Feuille As String, ByVal lignes As Collection)
    Dim wb As Workbook, ws As Worksheet, ligne As Long, valeurs As Object
    OuvrirTransaction fichier, wb
    On Error GoTo Echec
    Set ws = wb.Worksheets(Feuille)
    ligne = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
    If ligne < 2 Then ligne = 2
    For Each valeurs In lignes
        EcrireValeurs ws, ligne, valeurs
        ligne = ligne + 1
    Next valeurs
    FermerTransaction fichier, wb, True
    Exit Sub
Echec:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    FermerTransaction fichier, wb, False
    Err.Raise numErr, "modBaseIO", descErr
End Sub

' --- plomberie --------------------------------------------------------

Private Sub OuvrirTransaction(ByVal fichier As String, ByRef wb As Workbook)
    Dim autre As Workbook, verrou As String, numero As Long, description As String
    If mTransaction Then Err.Raise vbObjectError + 601, "modBaseIO", "Une transaction est deja en cours."
    For Each autre In Workbooks
        If StrComp(autre.FullName, fichier, vbTextCompare) = 0 Then Err.Raise vbObjectError + 602, "modBaseIO", "Fermez la base ouverte dans Excel : " & fichier
    Next autre
    verrou = NomVerrou(fichier)
    If Not modFichiers.AcquerirVerrou(verrou, 8000) Then Err.Raise vbObjectError + 603, "modBaseIO", "Base occupee. Reessayez dans quelques secondes."
    mEcran = Application.ScreenUpdating: mEvenements = Application.EnableEvents: mTransaction = True
    On Error GoTo Echec
    Application.ScreenUpdating = False: Application.EnableEvents = False
    SauvegardeQuotidienne fichier
    Set wb = Workbooks.Open(fichier, UpdateLinks:=0, ReadOnly:=False, Notify:=False, AddToMru:=False)
    If wb.ReadOnly Then Err.Raise vbObjectError + 604, "modBaseIO", "La base est ouverte en lecture seule."
    wb.Windows(1).Visible = False
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Set wb = Nothing
    modFichiers.RelacherVerrou verrou
    Application.ScreenUpdating = mEcran: Application.EnableEvents = mEvenements: mTransaction = False
    On Error GoTo 0
    Err.Raise numero, "modBaseIO.OuvrirTransaction", description
End Sub
Private Sub FermerTransaction(ByVal fichier As String, ByVal wb As Workbook, ByVal enregistrer As Boolean)
    Dim numero As Long, description As String
    If Not mTransaction Then Exit Sub
    On Error GoTo Echec
    If Not wb Is Nothing Then
        If enregistrer Then
            If wb.ReadOnly Then Err.Raise vbObjectError + 605, "modBaseIO", "Enregistrement refuse : base en lecture seule."
            wb.Windows(1).Visible = True
            wb.Save
            If Not wb.Saved Then Err.Raise vbObjectError + 606, "modBaseIO", "La base n a pas ete enregistree."
        End If
        wb.Close SaveChanges:=False
    End If
Sortie:
    modFichiers.RelacherVerrou NomVerrou(fichier)
    Application.ScreenUpdating = mEcran: Application.EnableEvents = mEvenements: mTransaction = False
    If numero <> 0 Then Err.Raise numero, "modBaseIO.FermerTransaction", description
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    GoTo Sortie
End Sub
Private Function NomVerrou(ByVal fichier As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    NomVerrou = fso.GetBaseName(fichier)
End Function

Private Sub SauvegardeQuotidienne(ByVal fichier As String)
    Dim cle As String
    cle = Format$(Date, "yyyymmdd") & "|" & fichier
    If InStr(mJourSauvegarde, cle) = 0 Then
        modFichiers.SauvegardeHorodatee fichier
        mJourSauvegarde = mJourSauvegarde & ";" & cle
    End If
End Sub

Private Sub EcrireValeurs(ByVal ws As Worksheet, ByVal ligne As Long, ByVal valeurs As Object)
    Dim entetes As Object, c As Long, cle As Variant
    Set entetes = CreateObject("Scripting.Dictionary")
    entetes.CompareMode = 1
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(-4159).Column   ' xlToLeft
        entetes(Trim$(CStr(ws.Cells(1, c).Value))) = c
    Next c
    For Each cle In valeurs.Keys
        If entetes.Exists(CStr(cle)) Then
            With ws.Cells(ligne, entetes(CStr(cle)))
                If LCase$(CStr(cle)) = "montant" Then
                    .NumberFormat = "0.00"
                    .Value2 = CCur(Val(Replace(CStr(valeurs(cle)), ",", ".")))
                Else
                    .NumberFormat = "@"
                    .Value2 = "'" & CStr(valeurs(cle))
                End If
            End With
        Else
            Err.Raise vbObjectError + 608, "modBaseIO", "Colonne absente : " & CStr(cle)
        End If
    Next cle
End Sub

Private Function TrouverLigne(ByVal ws As Worksheet, ByVal colonneCle As String, _
                              ByVal valeurCle As String) As Long
    Dim entetes As Object, c As Long, r As Long, colCle As Long, derniere As Long
    Set entetes = CreateObject("Scripting.Dictionary")
    entetes.CompareMode = 1
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(-4159).Column
        entetes(Trim$(CStr(ws.Cells(1, c).Value))) = c
    Next c
    If Not entetes.Exists(colonneCle) Then Exit Function
    colCle = entetes(colonneCle)
    derniere = ws.Cells(ws.Rows.Count, colCle).End(-4162).Row
    For r = 2 To derniere
        If Trim$(CStr(ws.Cells(r, colCle).Value)) = valeurCle Then
            TrouverLigne = r
            Exit Function
        End If
    Next r
End Function

Private Function ProchainID(ByVal ws As Worksheet, ByVal prefixe As String) As String
    Dim derniere As Long, r As Long, v As String, num As Long, maxi As Long
    derniere = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    For r = 2 To derniere
        v = Trim$(CStr(ws.Cells(r, 1).Value))
        If Left$(v, Len(prefixe)) = prefixe Then
            num = Val(mID$(v, Len(prefixe) + 1))
            If num > maxi Then maxi = num
        End If
    Next r
    ProchainID = prefixe & Format$(maxi + 1, "00000")
End Function

' Cree un classeur de base s'il n'existe pas (ex : agenda d'une nouvelle annee)
Public Sub CreerClasseurSiAbsent(ByVal fichier As String, ByVal Feuille As String, ByVal entetes As Variant)
    Dim wb As Workbook, i As Long, verrou As String, tmp As String, numero As Long, description As String
    Dim ecran As Boolean, events As Boolean
    If modFichiers.FichierExiste(fichier) Then Exit Sub
    verrou = NomVerrou(fichier)
    If Not modFichiers.AcquerirVerrou(verrou, 8000) Then Err.Raise vbObjectError + 607, "modBaseIO", "Creation de base deja en cours."
    ecran = Application.ScreenUpdating: events = Application.EnableEvents
    On Error GoTo Echec
    If modFichiers.FichierExiste(fichier) Then GoTo Sortie
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = Feuille
    For i = LBound(entetes) To UBound(entetes)
        wb.Worksheets(1).Cells(1, i - LBound(entetes) + 1).Value2 = entetes(i)
    Next i
    wb.Worksheets(1).Rows(1).Font.Bold = True
    tmp = Left$(fichier, InStrRev(fichier, "\")) & modFichiers.IdUnique() & ".xlsx"
    wb.SaveAs tmp, 51
    wb.Close False
    Set wb = Nothing
    modFichiers.RenommerAtomique tmp, fichier
Sortie:
    modFichiers.RelacherVerrou verrou
    Application.ScreenUpdating = ecran: Application.EnableEvents = events
    modFichiers.SupprimerTemporaire tmp
    If numero <> 0 Then Err.Raise numero, "modBaseIO.CreerClasseurSiAbsent", description
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    GoTo Sortie
End Sub

' Verification d'idempotence SOUS LE MEME VERROU que l'insertion.
Public Function AjouterSeanceUnique(ByVal fichier As String, ByVal lignes As Collection, ByVal seanceID As String) As Boolean
    Dim wb As Workbook, ws As Worksheet, valeurs As Object, ligne As Long
    Dim numero As Long, description As String, colPatient As Long, c As Long
    If lignes Is Nothing Then Err.Raise vbObjectError + 656, , "Seance sans actes."
    If lignes.Count = 0 Or Len(Trim$(seanceID)) = 0 Then Err.Raise vbObjectError + 656, , "Seance sans actes ou identifiant."
    OuvrirTransaction fichier, wb
    On Error GoTo Echec
    Set ws = wb.Worksheets("JOURNAL")
    ligne = TrouverLigne(ws, "SeanceID", seanceID)
    If ligne > 0 Then
        For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If CStr(ws.Cells(1, c).Value2) = "PatientID" Then colPatient = c
        Next c
        If colPatient = 0 Then Err.Raise vbObjectError + 657, , "Colonne PatientID absente du journal."
        For Each valeurs In lignes
            If CStr(ws.Cells(ligne, colPatient).Value2) <> CStr(valeurs("PatientID")) Then
                Err.Raise vbObjectError + 658, , "Identifiant de consultation deja utilise pour un autre patient."
            End If
        Next valeurs
        FermerTransaction fichier, wb, False
        Exit Function
    End If
    ligne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    For Each valeurs In lignes
        EcrireValeurs ws, ligne, valeurs
        ligne = ligne + 1
    Next valeurs
    FermerTransaction fichier, wb, True
    AjouterSeanceUnique = True
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    FermerTransaction fichier, wb, False
    Err.Raise numero, "modBaseIO.AjouterSeanceUnique", description
End Function

Public Sub MarquerFeuilleImprimee(ByVal fichier As String, ByVal seanceID As String)
    Dim wb As Workbook, ws As Worksheet, ligne As Long, valeurs As Object, c As Long, colSeance As Long
    Dim numero As Long, description As String
    OuvrirTransaction fichier, wb
    On Error GoTo Echec
    Set ws = wb.Worksheets("JOURNAL")
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If CStr(ws.Cells(1, c).Value2) = "SeanceID" Then colSeance = c
    Next c
    If colSeance = 0 Then Err.Raise vbObjectError + 609, "modBaseIO", "Colonne SeanceID absente."
    Set valeurs = CreateObject("Scripting.Dictionary")
    valeurs("FeuilleSoinsImprimee") = "O"
    For ligne = 2 To ws.Cells(ws.Rows.Count, colSeance).End(xlUp).Row
        If CStr(ws.Cells(ligne, colSeance).Value2) = seanceID Then EcrireValeurs ws, ligne, valeurs
    Next ligne
    FermerTransaction fichier, wb, True
    Exit Sub
Echec:
    numero = Err.Number: description = Err.Description
    FermerTransaction fichier, wb, False
    Err.Raise numero, "modBaseIO.MarquerFeuilleImprimee", description
End Sub

Private Sub VerifierCreneauSousVerrou(ByVal ws As Worksheet, ByVal valeurs As Object)
    Dim entetes As Object, c As Long, ligne As Long, k As Variant, statut As String
    Set entetes = CreateObject("Scripting.Dictionary")
    entetes.CompareMode = 1
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        entetes(Trim$(CStr(ws.Cells(1, c).Value2))) = c
    Next c
    For Each k In Array("Date", "Heure", "DureeMin", "Statut")
        If Not entetes.Exists(k) Then Err.Raise vbObjectError + 654, , "Colonne agenda absente : " & k
    Next k
    For ligne = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        statut = CStr(ws.Cells(ligne, entetes("Statut")).Value2)
        If statut <> "Annule" Then
            If CStr(ws.Cells(ligne, entetes("Date")).Value2) = CStr(valeurs("Date")) Then
                If modTexte.IntervallesSeChevauchent(modTexte.MinutesDepuisMinuit(CStr(valeurs("Heure"))), CLng(valeurs("DureeMin")), _
                    modTexte.MinutesDepuisMinuit(CStr(ws.Cells(ligne, entetes("Heure")).Value2)), CLng(ws.Cells(ligne, entetes("DureeMin")).Value2)) Then
                    Err.Raise vbObjectError + 655, , "Ce rendez-vous chevauche un creneau occupe. Choisissez un autre horaire."
                End If
            End If
        End If
    Next ligne
End Sub
