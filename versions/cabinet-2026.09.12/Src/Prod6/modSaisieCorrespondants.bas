Attribute VB_Name = "modSaisieCorrespondants"
Option Explicit
Public Sub NC_OuvrirNouveauCorrespondant()
    On Error GoTo Echec
    Dim d As Object, p As Object, r As Object
    Set d = modServiceNas.Parametres()
    d("Nom") = Trim$(InputBox("Nom du medecin ou de la structure :", "Nouveau correspondant"))
    If Len(d("Nom")) = 0 Then Exit Sub
    d("Prenom") = Trim$(InputBox("Prenom (laisser vide pour une structure) :", "Nouveau correspondant"))
    d("Adresse1") = Trim$(InputBox("Adresse :", "Nouveau correspondant"))
    d("CP") = Trim$(InputBox("Code postal :", "Nouveau correspondant"))
    d("Ville") = Trim$(InputBox("Ville :", "Nouveau correspondant"))
    d("TypesExamen") = Trim$(InputBox("Codes des examens separes par ; (laisser vide pour un medecin traitant) :", "Nouveau correspondant"))
    d("BlocDestinataire") = d("Nom") & " " & d("Prenom") & vbCrLf & d("Adresse1") & vbCrLf & d("CP") & " " & d("Ville")
    d("Actif") = "1": d("AValider") = "0": d("ParDefaut") = "0"
    d("FormuleAppel") = "Cher Confrere,": d("FormulePolitesse") = "Bien confraternellement."
    If MsgBox("Enregistrer ce correspondant sur le NAS ?" & vbCrLf & d("BlocDestinataire"), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Set p = modServiceNas.Parametres(): Set p("data") = d
    Set r = modServiceNas.Appeler("correspondent.save", p)
    MsgBox "Correspondant enregistre : " & CStr(r("ID")), vbInformation, "Cabinet"
    Exit Sub
Echec:
    MsgBox "Enregistrement interrompu : " & Err.Description, vbExclamation, "Cabinet"
End Sub
