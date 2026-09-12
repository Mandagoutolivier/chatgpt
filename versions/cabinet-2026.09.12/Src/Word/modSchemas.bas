Attribute VB_Name = "modSchemas"
Option Explicit
' =====================================================================
' modSchemas - Colonnes des bases partagees par les deux postes.
' Sert a CREER une base absente (installation neuve) et a completer les
' colonnes manquantes d'une base existante : le logiciel ne doit jamais
' s'arreter sur "Fichier introuvable : Patients.xlsx".
' Aucune donnee nominative n'est stockee ici, seulement des en-tetes.
' =====================================================================

' Feuille PATIENTS de Base\Patients.xlsx
Public Function EntetesPatients() As Variant
    EntetesPatients = Array("ID", "Nom", "Prenom", "NomNaissance", "Sexe", "DDN", "NIR", _
                            "Adresse1", "Adresse2", "CP", "Ville", "Tel", "Mobile", "Email", _
                            "MedTraitantID", "Mutuelle", "ALD", _
                            "AssureNom", "AssurePrenom", "AssureDDN", "AssureNIR", _
                            "Notes", "DateCreation", "Actif", "Provisoire")
End Function

' Feuille CORRESPONDANTS du meme classeur
Public Function EntetesCorrespondants() As Variant
    EntetesCorrespondants = Array("ID", "Titre", "Nom", "Prenom", "Specialite", _
                                  "Adresse1", "Adresse2", "CP", "Ville", "Tel", "Email", _
                                  "BlocDestinataire", "Tutoiement", _
                                  "FormuleAppel", "FormulePolitesse", "Notes", "Actif")
End Function

' Feuilles de Base\Patients.xlsx, dans l'ordre
Public Function FeuillesBasePatients() As Variant
    FeuillesBasePatients = Array("PATIENTS", "CORRESPONDANTS")
End Function

Public Function EntetesDe(ByVal feuille As String) As Variant
    Select Case UCase$(Trim$(feuille))
        Case "PATIENTS":       EntetesDe = EntetesPatients()
        Case "CORRESPONDANTS": EntetesDe = EntetesCorrespondants()
        Case Else:             EntetesDe = Array("ID")
    End Select
End Function
