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

Public Function AjouterSeanceUnique(ByVal fichier As String, ByVal lignes As Collection, ByVal seanceID As String, _
                                    Optional ByRef selectionDifferente As Boolean = False) As Boolean
    Dim p As Object, r As Object
    Set p = modServiceNas.Parametres(): p("id") = seanceID: Set p("lignes") = lignes
    Set r = modServiceNas.Appeler("bill", p)
    If r.Exists("selection_differente") Then selectionDifferente = CBool(r("selection_differente"))
    AjouterSeanceUnique = CBool(r("ajoute"))
End Function

Public Sub MarquerFeuilleImprimee(ByVal fichier As String, ByVal seanceID As String)
    Dim r As Object
    Set r = modServiceNas.CommandeID("printed", seanceID)
End Sub

Private Function AnneeFichier(ByVal fichier As String) As String
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "_(\d{4})\.xlsx$": re.IgnoreCase = True
    Set matches = re.Execute(fichier)
    If matches.Count > 0 Then AnneeFichier = matches(0).SubMatches(0)
End Function

Public Function LireJournal(Optional ByVal annee As String = "", Optional ByVal seanceID As String = "") As Collection
    Dim p As Object, r As Object, ligne As Object, resultat As New Collection
    Set p = modServiceNas.Parametres(): p("year") = annee: p("id") = seanceID: p("offset") = 0
    Do
        Set r = modServiceNas.Appeler("journal.read", p)
        For Each ligne In r("items"): resultat.Add ligne: Next ligne
        If IsNull(r("next")) Then Exit Do
        p("offset") = r("next")
    Loop
    Set LireJournal = resultat
End Function
