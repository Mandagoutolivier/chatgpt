Attribute VB_Name = "modAgenda"
Option Explicit
' =====================================================================
' modAgenda - Rendez-vous (poste secretaire). Fichier Agenda_AAAA.xlsx,
' feuille RDV. Statuts : Prevu / Arrive / Honore / Absent / Annule.
' =====================================================================

Public Function EntetesRdv() As Variant
    EntetesRdv = Array("ID", "PatientID", "Date", "Heure", "DureeMin", "TypeActe", _
                       "Statut", "HeureArrivee", "Notes", "DateCreation")
End Function

Public Sub AssurerAgendaAnnee(Optional ByVal annee As Long = 0)
    modBaseIO.CreerClasseurSiAbsent modConfig.FichierAgenda(annee), "RDV", EntetesRdv()
End Sub

Public Function AjouterRdv(ByVal patientID As String, ByVal dateRdv As String, _
                           ByVal heure As String, ByVal dureeMin As String, _
                           ByVal typeActe As String, ByVal notes As String) As String
    Dim d As Object
    If Not modTexte.DateFrValide(dateRdv) Then Err.Raise vbObjectError + 650, "modAgenda", "Date du rendez-vous invalide."
    If Not modTexte.HeureValide(heure) Then Err.Raise vbObjectError + 651, "modAgenda", "Heure du rendez-vous invalide."
    If Not IsNumeric(dureeMin) Then Err.Raise vbObjectError + 652, "modAgenda", "Duree invalide."
    If Val(dureeMin) <> Fix(Val(dureeMin)) Or Val(dureeMin) < 1 Or Val(dureeMin) > 480 Then Err.Raise vbObjectError + 653, "modAgenda", "Duree invalide."
    If modTexte.MinutesDepuisMinuit(heure) + Val(dureeMin) > 1440 Then Err.Raise vbObjectError + 656, , "Le rendez-vous depasse minuit."
    dateRdv = Format$(modTexte.DateFr(dateRdv), "dd/mm/yyyy")
    heure = Format$(modTexte.MinutesDepuisMinuit(heure) \ 60, "00") & ":" & Format$(modTexte.MinutesDepuisMinuit(heure) Mod 60, "00")
    AssurerAgendaAnnee AnneeDeDate(dateRdv)
    Set d = CreateObject("Scripting.Dictionary")
    d("PatientID") = patientID
    d("Date") = dateRdv
    d("Heure") = heure
    d("DureeMin") = dureeMin
    d("TypeActe") = typeActe
    d("Statut") = "Prevu"
    d("Notes") = notes
    d("DateCreation") = Format$(Now, "dd/mm/yyyy hh:nn")
    AjouterRdv = modBaseIO.AjouterLigne(modConfig.FichierAgenda(AnneeDeDate(dateRdv)), "RDV", d, "R")
End Function

Public Sub MarquerStatut(ByVal rdvID As String, ByVal statut As String, _
                         Optional ByVal annee As Long = 0)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Statut") = statut
    If statut = "Arrive" Then d("HeureArrivee") = Format$(Now, "hh:nn")
    modBaseIO.ModifierLigne modConfig.FichierAgenda(annee), "RDV", "ID", rdvID, d
End Sub

' RDV d'une date (defaut aujourd'hui), tries par heure
Public Function RdvDuJour(Optional ByVal dateRdv As String = "") As Collection
    Dim tous As Collection, jour As Collection, r As Object
    Dim pats As Object, p As Object
    If Len(dateRdv) = 0 Then dateRdv = Format$(Date, "dd/mm/yyyy")
    AssurerAgendaAnnee AnneeDeDate(dateRdv)
    Set tous = modBaseIO.LireTableX(modConfig.FichierAgenda(AnneeDeDate(dateRdv)), "RDV")
    ' index patients pour l'affichage
    Set pats = CreateObject("Scripting.Dictionary")
    For Each p In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        Set pats(p("ID")) = p
    Next p
    Set jour = New Collection
    For Each r In tous
        If r("Date") = dateRdv Then
            If pats.Exists(r("PatientID")) Then
                r("Nom") = pats(r("PatientID"))("Nom")
                r("Prenom") = pats(r("PatientID"))("Prenom")
            Else
                r("Nom") = "?"
                r("Prenom") = ""
            End If
            InsererParHeure jour, r
        End If
    Next r
    Set RdvDuJour = jour
End Function

Private Sub InsererParHeure(ByVal col As Collection, ByVal r As Object)
    Dim i As Long
    For i = 1 To col.Count
        If col(i)("Heure") > r("Heure") Then
            col.Add r, , i
            Exit Sub
        End If
    Next i
    col.Add r
End Sub

Private Function AnneeDeDate(ByVal dateTexte As String) As Long
    AnneeDeDate = Year(modTexte.DateFr(dateTexte))
End Function
