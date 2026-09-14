Attribute VB_Name = "modBaseIO"
Option Explicit
' Les bases ne sont plus ouvertes dans Excel : toutes les mutations passent
' par les transactions du service Synology. Les signatures preservent l'UI.

Public Function LireTableX(ByVal fichier As String, ByVal Feuille As String, Optional ByVal colonneNonVide As String = "ID") As Collection
    Dim p As Object, r As Object
    If UCase$(Feuille) = "JOURNAL" Then
        Set LireTableX = LireJournal(AnneeFichier(fichier))
    Else
        Set LireTableX = modServiceNas.LireTable(UCase$(Feuille), "", IIf(UCase$(Feuille) = "RDV", AnneeFichier(fichier), ""))
    End If
End Function

Public Function AjouterLigne(ByVal fichier As String, ByVal Feuille As String, ByVal valeurs As Object, Optional ByVal prefixeID As String = "") As String
    Dim p As Object, r As Object
    Set p = modServiceNas.Parametres(): p("genre") = UCase$(Feuille): Set p("data") = valeurs
    Set r = modServiceNas.Appeler("table.add", p)
    AjouterLigne = CStr(r("ID"))
End Function

Public Sub ModifierLigne(ByVal fichier As String, ByVal Feuille As String, ByVal colonneCle As String, ByVal valeurCle As String, ByVal valeurs As Object)
    Dim p As Object, r As Object
    If colonneCle <> "ID" Then Err.Raise vbObjectError + 1110, , "Modification sans identifiant stable interdite."
    If Not valeurs.Exists("_revision") Then Err.Raise vbObjectError + 1111, , "Revision de la fiche absente : rechargez-la."
    valeurs("ID") = valeurCle
    Set p = modServiceNas.Parametres(): p("genre") = UCase$(Feuille): Set p("data") = valeurs
    Set r = modServiceNas.Appeler("table.update", p)
End Sub

Public Sub AjouterLignes(ByVal fichier As String, ByVal Feuille As String, ByVal lignes As Collection)
    Err.Raise vbObjectError + 1112, , "Utilisez EnregistrerSeance pour une ecriture comptable."
End Sub

Public Sub CreerClasseurSiAbsent(ByVal fichier As String, ByVal Feuille As String, ByVal entetes As Variant)
    ' Le schema serveur est versionne ; aucun classeur de base n'est cree.
End Sub

Public Function EnregistrerActes(ByVal lignes As Collection, ByVal seanceID As String, ByVal publicationID As String) As Object
    Dim p As Object
    Set p = modServiceNas.Parametres(): p("id") = seanceID: p("publication_id") = publicationID: Set p("lignes") = lignes
    Set EnregistrerActes = modServiceNas.Appeler("bill", p)
End Function

Private Function AnneeFichier(ByVal fichier As String) As String
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "_(\d{4})\.xlsx$": re.IgnoreCase = True
    Set matches = re.Execute(fichier)
    If matches.Count > 0 Then AnneeFichier = matches(0).SubMatches(0)
End Function

Public Function LireJournal(Optional ByVal annee As String = "", Optional ByVal seanceID As String = "") As Collection
    Dim p As Object
    Set p = modServiceNas.Parametres(): p("year") = annee: p("id") = seanceID
    Set LireJournal = modServiceNas.LirePages("journal.read", p)
End Function
