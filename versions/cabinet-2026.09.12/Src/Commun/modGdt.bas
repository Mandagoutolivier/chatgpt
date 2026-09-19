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
            "(en recette, dossier dedie non surveille par le profil ECG clinique)."
    End If
    If Right$(dossier, 1) = "\" Then dossier = Left$(dossier, Len(dossier) - 1)
    If Not modFichiers.DossierExiste(dossier) Then
        Err.Raise vbObjectError + 801, "modGdt", "Dossier ECG introuvable : " & dossier
    End If

    contenu = ConstruireGdt(pat)
    chemin = dossier & "\IMPORT.GDT"
    Dim tmp As String, numero As Long, description As String
    tmp = dossier & "\" & modFichiers.IdUnique() & ".tmp"
    On Error GoTo Echec
    modFichiers.EcrireTexteAnsi tmp, contenu
    modFichiers.RenommerAtomique tmp, chemin, True
    modLog.LogInfo "GDT ecrit pour " & pat("ID") & " -> " & chemin
    EcrireGdtPatient = chemin
    Exit Function
Echec:
    numero = Err.Number: description = Err.Description
    modFichiers.SupprimerTemporaire tmp
    Err.Raise numero, "modGdt", description
End Function

Public Function ConstruireGdt(ByVal pat As Object) As String
    Dim lignes As Collection, l As Variant, total As Long, contenu As String
    If Len(Trim$(CStr(pat("ID")))) = 0 Or Len(Trim$(CStr(pat("Nom")))) = 0 Or Len(Trim$(CStr(pat("Prenom")))) = 0 Then
        Err.Raise vbObjectError + 805, "modGdt", "Identite patient incomplete."
    End If
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
    Select Case modTexte.SexeNormalise(sexe)
        Case "M": SexeVersGdt = "1"
        Case "F": SexeVersGdt = "2"
        Case Else: Err.Raise vbObjectError + 802, "modGdt", "Sexe patient non renseigne pour l ECG."
    End Select
End Function
Private Function LigneGdt(ByVal champ As String, ByVal valeur As String) As String
    Dim st As Object, retour As String, i As Long
    If Len(champ & valeur) + 5 > 999 Then Err.Raise vbObjectError + 806, "modGdt", "Champ GDT trop long."
    For i = 0 To 31
        If InStr(valeur, Chr$(i)) > 0 Then Err.Raise vbObjectError + 807, "modGdt", "Caractere de controle interdit dans le GDT."
    Next i
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "windows-1252": st.Open
    st.WriteText valeur
    st.Position = 0
    retour = st.ReadText
    st.Close
    If retour <> valeur Then Err.Raise vbObjectError + 808, "modGdt", "Un caractere de l identite ne peut pas etre transmis en CP1252."
    LigneGdt = Format$(Len(champ & valeur) + 5, "000") & champ & valeur
End Function
' "01/01/1935" -> "01011935"
Private Function DdnVersGdt(ByVal ddn As String) As String
    If Not modTexte.DateFrValide(ddn) Then Err.Raise vbObjectError + 803, "modGdt", "Date de naissance invalide."
    If modTexte.DateFr(ddn) > Date Then Err.Raise vbObjectError + 804, "modGdt", "Date de naissance future."
    DdnVersGdt = Format$(modTexte.DateFr(ddn), "ddmmyyyy")
End Function
