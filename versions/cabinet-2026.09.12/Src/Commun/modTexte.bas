Attribute VB_Name = "modTexte"
Option Explicit
' =====================================================================
' modTexte - Utilitaires texte : pliage (minuscules sans accents) pour
' les recherches et l'anonymisation insensibles a la casse/aux accents.
' =====================================================================

Public Function Plier(ByVal s As String) As String
    Dim src As String, dst As String, i As Long, p As Long, c As String, res As String
    s = LCase$(s)
    ' digrammes d'abord
    s = Replace(s, ChrW$(230), "ae")       ' ae lie
    s = Replace(s, ChrW$(339), "oe")       ' oe lie
    src = ChrW$(224) & ChrW$(226) & ChrW$(228) & ChrW$(225) & ChrW$(227) & ChrW$(229) & _
          ChrW$(233) & ChrW$(232) & ChrW$(234) & ChrW$(235) & _
          ChrW$(237) & ChrW$(236) & ChrW$(238) & ChrW$(239) & _
          ChrW$(243) & ChrW$(242) & ChrW$(244) & ChrW$(246) & ChrW$(245) & _
          ChrW$(250) & ChrW$(249) & ChrW$(251) & ChrW$(252) & _
          ChrW$(253) & ChrW$(255) & ChrW$(231) & ChrW$(241)
    dst = "aaaaaaeeeeiiiiooooouuuuyycn"
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        p = InStr(src, c)
        If p > 0 Then res = res & Mid$(dst, p, 1) Else res = res & c
    Next i
    Plier = res
End Function

' --- Sexe / civilite ---------------------------------------------------
' Le sexe des fiches est saisi de facons variees (M/F, H, Mme, Madame,
' Monsieur, Homme/Femme, 1/2, Masculin/Feminin...). Tout passe par
' SexeNormalise -> "M", "F" ou "" (inconnu). Un sexe INCONNU ne donne
' plus "Monsieur" par defaut : pas de civilite, "ne(e) le".
Public Function SexeNormalise(ByVal v As String) As String
    Dim t As String
    t = Plier(Trim$(v))
    If Len(t) = 0 Then Exit Function
    Select Case t
        Case "mme", "madame", "mlle", "mademoiselle", "femme", "feminin", "f", "2", "w", "woman", "female"
            SexeNormalise = "F"
        Case "m", "m.", "mr", "monsieur", "homme", "masculin", "h", "1", "male", "man"
            SexeNormalise = "M"
    End Select
End Function

' Sexe d'une fiche (patient ou correspondant), quelle que soit la colonne
' utilisee : Sexe, Civilite, Genre, Titre (Mme/M./Dr sans info).
Public Function SexePatient(ByVal pat As Object) As String
    Dim k As Variant
    If pat Is Nothing Then Exit Function
    For Each k In Array("Sexe", "Civilite", "Genre", "Sex", "Titre")
        If pat.Exists(k) Then
            SexePatient = SexeNormalise(CStr(pat(k)))
            If Len(SexePatient) > 0 Then Exit Function
        End If
    Next k
End Function

' Date de naissance d'une fiche, quelle que soit la colonne (DDN,
' DateNaissance, Date de naissance, Naissance, NeLe).
Public Function DdnPatient(ByVal pat As Object) As String
    Dim k As Variant
    If pat Is Nothing Then Exit Function
    For Each k In Array("DDN", "DateNaissance", "Date de naissance", "Date naissance", "Naissance", "NeLe", "DateDeNaissance")
        If pat.Exists(k) Then
            DdnPatient = Trim$(CStr(pat(k)))
            If Len(DdnPatient) > 0 Then Exit Function
        End If
    Next k
End Function

' Civilite d'apres le sexe (M/F, ou toute valeur normalisable) ; "" si inconnu
Public Function Civilite(ByVal sexe As String) As String
    Select Case SexeNormalise(sexe)
        Case "F": Civilite = "Madame"
        Case "M": Civilite = "Monsieur"
    End Select
End Function

Public Function CiviliteCourte(ByVal sexe As String) As String
    Select Case SexeNormalise(sexe)
        Case "F": CiviliteCourte = "Mme"
        Case "M": CiviliteCourte = "M."
    End Select
End Function

Public Function NeLe(ByVal sexe As String) As String
    Select Case SexeNormalise(sexe)
        Case "F": NeLe = "née le"
        Case "M": NeLe = "né le"
        Case Else: NeLe = "né(e) le"
    End Select
End Function

' ---------------------------------------------------------------------
' Dates et heures : validation STRICTE (le 31 février est refusé, une
' heure doit être hh:mm entre 00:00 et 23:59).
' ---------------------------------------------------------------------
Public Function DateFrValide(ByVal s As String) As Boolean
    Dim p() As String, j As Long, m As Long, a As Long
    s = Trim$(s)
    p = Split(s, "/")
    If UBound(p) <> 2 Then Exit Function
    If Len(p(0)) < 1 Or Len(p(0)) > 2 Then Exit Function
    If Len(p(1)) < 1 Or Len(p(1)) > 2 Then Exit Function
    If Len(p(2)) <> 4 Then Exit Function
    If Not (EstEntier(p(0)) And EstEntier(p(1)) And EstEntier(p(2))) Then Exit Function
    j = Val(p(0)): m = Val(p(1)): a = Val(p(2))
    If m < 1 Or m > 12 Then Exit Function
    If a < 1880 Or a > Year(Date) + 5 Then Exit Function
    If j < 1 Or j > JoursDuMois(m, a) Then Exit Function
    DateFrValide = True
End Function

Public Function DateFr(ByVal s As String) As Date
    Dim p() As String
    If Not DateFrValide(s) Then Err.Raise vbObjectError + 810, "modTexte", "Date invalide : " & s
    p = Split(Trim$(s), "/")
    DateFr = DateSerial(Val(p(2)), Val(p(1)), Val(p(0)))
End Function

Public Function JoursDuMois(ByVal m As Long, ByVal a As Long) As Long
    Select Case m
        Case 1, 3, 5, 7, 8, 10, 12: JoursDuMois = 31
        Case 4, 6, 9, 11: JoursDuMois = 30
        Case 2
            If (a Mod 4 = 0 And a Mod 100 <> 0) Or (a Mod 400 = 0) Then JoursDuMois = 29 Else JoursDuMois = 28
    End Select
End Function

Private Function EstEntier(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    EstEntier = True
End Function

Public Function HeureValide(ByVal s As String) As Boolean
    Dim p() As String
    s = Replace(LCase$(Trim$(s)), "h", ":")
    p = Split(s, ":")
    If UBound(p) <> 1 Then Exit Function
    If Not (EstEntier(Trim$(p(0))) And EstEntier(Trim$(p(1)))) Then Exit Function
    If Val(p(0)) < 0 Or Val(p(0)) > 23 Then Exit Function
    If Val(p(1)) < 0 Or Val(p(1)) > 59 Then Exit Function
    HeureValide = True
End Function

' Minutes depuis minuit ; -1 si l'heure est invalide
Public Function MinutesDepuisMinuit(ByVal heure As String) As Long
    Dim p() As String
    If Not HeureValide(heure) Then MinutesDepuisMinuit = -1: Exit Function
    p = Split(Replace(LCase$(Trim$(heure)), "h", ":"), ":")
    MinutesDepuisMinuit = Val(p(0)) * 60 + Val(p(1))
End Function

' Deux creneaux [debut, debut+duree[ se recouvrent-ils ?
Public Function IntervallesSeChevauchent(ByVal debut1 As Long, ByVal duree1 As Long, _
                                         ByVal debut2 As Long, ByVal duree2 As Long) As Boolean
    If duree1 <= 0 Then duree1 = 1
    If duree2 <= 0 Then duree2 = 1
    IntervallesSeChevauchent = (debut1 < debut2 + duree2) And (debut2 < debut1 + duree1)
End Function
