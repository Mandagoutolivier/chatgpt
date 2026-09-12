Attribute VB_Name = "modSubstitutions"
Option Explicit
' =====================================================================
' modSubstitutions - Corrections locales appliquees au texte dicte
' AVANT l'anonymisation et l'appel API (reprend les macros historiques
' du cabinet). Regles dans Config\substitutions.txt :
'   texte cherche|remplacement   (remplacement litteral, sensible casse)
'   MAJ|mot                      (mot en MAJUSCULES, insensible casse)
' =====================================================================

Public Function AppliquerSubstitutions(ByVal texte As String) As String
    Dim chemin As String, contenu As String, lignes() As String
    Dim i As Long, p As Long, motif As String, remp As String, ligne As String
    chemin = modConfig.chemin("Config") & "\substitutions.txt"
    If Not modFichiers.FichierExiste(chemin) Then
        AppliquerSubstitutions = texte
        Exit Function
    End If
    contenu = modFichiers.LireTexteUTF8(chemin)
    contenu = Replace(contenu, vbCrLf, vbLf)
    lignes = Split(contenu, vbLf)
    For i = LBound(lignes) To UBound(lignes)
        ligne = lignes(i)
        If Len(Trim$(ligne)) > 0 And Left$(Trim$(ligne), 1) <> "#" Then
            p = InStr(ligne, "|")
            If p > 0 Then
                motif = Left$(ligne, p - 1)
                remp = Mid$(ligne, p + 1)
                If motif = "MAJ" Then
                    remp = Trim$(remp)
                    If Len(remp) > 1 Then
                        texte = Replace(texte, remp, UCase$(remp), 1, -1, vbTextCompare)
                    End If
                ElseIf Len(motif) > 0 Then
                    texte = Replace(texte, motif, remp)
                End If
            End If
        End If
    Next i
    AppliquerSubstitutions = texte
End Function
