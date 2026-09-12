Attribute VB_Name = "modJournal"
Option Explicit
' =====================================================================
' modJournal - Journal comptable (Actes\Journal_AAAA.xlsx, feuille
' JOURNAL, une ligne par acte). Exploitable directement par le comptable.
' =====================================================================

Public Function EntetesJournal() As Variant
    EntetesJournal = Array("Date", "SeanceID", "PatientID", "Nom", "Prenom", "DDN", "NIR", _
                           "CodeActe", "Montant", "ModePaiement", "TiersPayant", "Paye", _
                           "DateEncaissement", "FeuilleSoinsImprimee", "Notes")
End Function

Public Sub AssurerJournalAnnee()
    modBaseIO.CreerClasseurSiAbsent modConfig.FichierJournal(), "JOURNAL", EntetesJournal()
End Sub

Public Sub Ajouter(ByVal lignes As Collection)
    AssurerJournalAnnee
    modBaseIO.AjouterLignes modConfig.FichierJournal(), "JOURNAL", lignes
End Sub

Public Sub OuvrirJournal()
    On Error GoTo Echec
    Dim wb As Workbook, ws As Worksheet, lignes As Collection, d As Object, entetes As Variant, r As Long, c As Long
    Set lignes = modBaseIO.LireTableX(modConfig.FichierJournal(), "JOURNAL", "SeanceID")
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    ws.Name = "Journal exporte": entetes = EntetesJournal()
    For c = LBound(entetes) To UBound(entetes): ws.Cells(1, c + 1).Value2 = entetes(c): Next c
    r = 2
    For Each d In lignes
        For c = LBound(entetes) To UBound(entetes)
            If d.Exists(CStr(entetes(c))) Then
                If CStr(entetes(c)) = "Montant" Then
                    ws.Cells(r, c + 1).Value2 = CDbl(Val(Replace(CStr(d(entetes(c))), ",", ".")))
                    ws.Cells(r, c + 1).NumberFormat = "0.00"
                Else
                    modFichiers.EcrireCelluleTexte ws.Cells(r, c + 1), CStr(d(entetes(c)))
                End If
            End If
        Next c
        r = r + 1
    Next d
    ws.Rows(1).Font.Bold = True: ws.Columns.AutoFit
    Application.StatusBar = "Export du journal serveur : les modifications de cette copie ne modifient pas la comptabilite."
    Exit Sub
Echec:
    MsgBox "Export impossible : " & Err.Description, vbExclamation, "Cabinet"
End Sub
