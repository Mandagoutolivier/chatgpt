Attribute VB_Name = "modAttenteLocale"
Option Explicit
' Liste en memoire. Chaque ouverture relit le NAS ; aucune base locale.

Public Function FiltrerEtTrier(ByVal items As Collection, ByVal jour As String) As Collection
    Dim col As New Collection, d As Object, i As Long, insere As Boolean, cle As Variant
    For Each d In items
        For Each cle In Array("DateArrivee", "Statut", "PatientID", "RdvID", "ConsultationID", "Nom", "Prenom", "DDN", "HeureArrivee", "HeureRdv")
            If Not d.Exists(CStr(cle)) Then Err.Raise vbObjectError + 973, , "Arrivee incomplete. Rechargez la file."
        Next cle
        If CStr(d("DateArrivee")) = jour And CStr(d("Statut")) = "Arrive" Then
            If Len(Trim$(CStr(d("PatientID")))) = 0 Or Len(Trim$(CStr(d("RdvID")))) = 0 Or Len(Trim$(CStr(d("ConsultationID")))) = 0 Then Err.Raise vbObjectError + 973, , "Arrivee sans identifiant."
            d("SourceNas") = CStr(d("ConsultationID"))
            d("Patient") = CStr(d("Nom")) & " " & CStr(d("Prenom"))
            insere = False
            For i = 1 To col.Count
                If CleTri(d) < CleTri(col(i)) Then
                    col.Add d, , i: insere = True: Exit For
                End If
            Next i
            If Not insere Then col.Add d
        End If
    Next d
    Set FiltrerEtTrier = col
End Function

Private Function CleTri(ByVal d As Object) As String
    CleTri = CStr(d("HeureArrivee")) & "|" & CStr(d("HeureRdv")) & "|" & CStr(d("ConsultationID"))
End Function

Public Function ChoisirAttente() As Object
    Dim r As Object, col As Collection, f As ufListe
    Set r = modServiceNas.Appeler("attentes", modServiceNas.Parametres())
    Set col = FiltrerEtTrier(modServiceNas.ItemsValides(r), Format$(Date, "dd/mm/yyyy"))
    If col.Count = 0 Then MsgBox "Aucun patient arrive en attente aujourd hui.", vbInformation, "Cabinet": Exit Function
    Set f = New ufListe
    f.Configurer "Patients arrives", col, Array("HeureArrivee", "Patient", "DDN"), "50 pt;200 pt;70 pt"
    f.Show vbModal
    If Not f.Annule Then Set ChoisirAttente = f.Resultat
    Unload f
End Function

Public Sub SynchroniserAttentes()
    Dim r As Object, col As Collection
    Set r = modServiceNas.Appeler("attentes", modServiceNas.Parametres())
    Set col = FiltrerEtTrier(modServiceNas.ItemsValides(r), Format$(Date, "dd/mm/yyyy"))
End Sub

Public Sub ConsommerAttente(ByVal attente As Object)
    Dim r As Object, k As Variant
    Set r = modServiceNas.CommandeID("claim", CStr(attente("SourceNas")))
    For Each k In r.Keys: attente(k) = r(k): Next k
End Sub

Public Sub LibererReservation(ByVal attente As Object)
    Dim r As Object
    If attente Is Nothing Then Exit Sub
    If Not attente.Exists("ReservationNas") Then Exit Sub
    Set r = modServiceNas.CommandeID("release", CStr(attente("ReservationNas")))
End Sub
Public Sub EnregistrerBrouillon(ByVal attente As Object, ByVal cheminBrouillon As String)
    Dim p As Object, r As Object
    Set p = modServiceNas.Parametres(): p("id") = CStr(attente("ReservationNas")): p("path") = cheminBrouillon
    Set r = modServiceNas.Appeler("draft", p)
End Sub
