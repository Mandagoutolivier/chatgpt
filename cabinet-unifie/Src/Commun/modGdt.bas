Attribute VB_Name = "modGdt"
Option Explicit
' =====================================================================
' modGdt - Envoi de l'identite patient au logiciel ECG Resting12Lead
' par fichier GDT (dialecte valide en essai reel le 31/08/2026) :
'   satz 6302, GDT 02.00, jeu de caracteres 9206=3 (ANSI),
'   date de naissance JJMMAAAA, fichier IMPORT.GDT en ANSI cp1252.
' Cote Resting12Lead : Parametres > Parametres Interface >
'   "Saisie Auto info Patient" = GDT (et "Connection Systeme Info DMS"
'   DEcochee - les deux modes sont exclusifs), Interface GDT In/Out
'   pointant sur le meme dossier que [ECG] DossierGdt.
' =====================================================================

' Ecrit IMPORT.GDT pour ce patient. dossierOverride : pour les tests.
' Renvoie le chemin ecrit. Erreur claire si non configure.
Public Function EcrireGdtPatient(ByVal pat As Object, _
                                 Optional ByVal dossierOverride As String = "") As String
    Dim dossier As String, chemin As String, contenu As String
    dossier = dossierOverride
    If Len(dossier) = 0 Then dossier = modConfig.Config("ECG", "DossierGdt", "")
    If Len(dossier) = 0 Then
        Err.Raise vbObjectError + 800, "modGdt", _
            "Envoi ECG non configure : renseignez [ECG] DossierGdt dans Config\config.ini " & _
            "(dossier surveille par Resting12Lead, ex : C:\Mandagout)."
    End If
    If Right$(dossier, 1) = "\" Then dossier = Left$(dossier, Len(dossier) - 1)
    If Not modFichiers.DossierExiste(dossier) Then
        Err.Raise vbObjectError + 801, "modGdt", "Dossier ECG introuvable : " & dossier
    End If

    contenu = ConstruireGdt(pat)
    chemin = dossier & "\IMPORT.GDT"
    modFichiers.EcrireTexteAnsi chemin, contenu
    modLog.LogInfo "GDT ecrit pour " & pat("ID") & " -> " & chemin
    EcrireGdtPatient = chemin
End Function

Public Function ConstruireGdt(ByVal pat As Object) As String
    Dim lignes As Collection, l As Variant, total As Long, contenu As String
    Set lignes = New Collection
    lignes.Add LigneGdt("8000", "6302")
    lignes.Add "PLACEHOLDER"
    lignes.Add LigneGdt("9206", "3")
    lignes.Add LigneGdt("9218", "02.00")
    lignes.Add LigneGdt("3000", pat("ID"))
    lignes.Add LigneGdt("3101", UCase$(pat("Nom")))
    lignes.Add LigneGdt("3102", pat("Prenom"))
    lignes.Add LigneGdt("3103", DdnVersGdt(pat("DDN")))
    lignes.Add LigneGdt("3110", SexeVersGdt(pat("Sexe")))
    lignes.Add LigneGdt("8402", modConfig.Config("ECG", "CodeExamen", "EKG01"))

    ' champ 8100 = longueur totale, sa propre ligne comprise (14 octets)
    total = 14
    For Each l In lignes
        If l <> "PLACEHOLDER" Then total = total + Len(l) + 2
    Next l

    For Each l In lignes
        If l = "PLACEHOLDER" Then
            contenu = contenu & LigneGdt("8100", Format$(total, "00000")) & vbCrLf
        Else
            contenu = contenu & l & vbCrLf
        End If
    Next l
    ConstruireGdt = contenu
End Function

Private Function SexeVersGdt(ByVal sexe As String) As String
    sexe = UCase$(Trim$(sexe))
    If sexe = "M" Or sexe = "1" Or sexe = "H" Then
        SexeVersGdt = "1"
    ElseIf sexe = "F" Or sexe = "2" Then
        SexeVersGdt = "2"
    Else
        SexeVersGdt = ""
    End If
End Function

Private Function LigneGdt(ByVal champ As String, ByVal valeur As String) As String
    LigneGdt = Format$(Len(champ & valeur) + 5, "000") & champ & valeur
End Function

' "01/01/1935" -> "01011935"
Private Function DdnVersGdt(ByVal ddn As String) As String
    Dim p() As String
    p = Split(Trim$(ddn), "/")
    If UBound(p) = 2 Then
        DdnVersGdt = Format$(Val(p(0)), "00") & Format$(Val(p(1)), "00") & Format$(Val(p(2)), "0000")
    Else
        DdnVersGdt = ""
    End If
End Function
