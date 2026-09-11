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
    AssurerJournalAnnee
    On Error GoTo Occupe
    Dim wb As Workbook
    Set wb = Workbooks.Open(modConfig.FichierJournal(), ReadOnly:=True)
    wb.Windows(1).visible = True     ' fichiers enregistres fenetre masquee (anciennes versions)
    wb.Activate
    Application.StatusBar = "Journal ouvert en lecture seule : les actes s'y ajoutent automatiquement."
    Exit Sub
Occupe:
    MsgBox "Impossible d'ouvrir le journal : " & Err.Description, vbExclamation, "Cabinet"
End Sub
