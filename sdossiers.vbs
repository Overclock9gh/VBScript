'Auteur Saïd Hamdane
'24 septembre 2019

'Objectif: Afficher tous les sous-dossiers d'un dossier spécifique

Option Explicit

Dim SystemeFic, SousDoss, Chemin

Set SystemeFic = CreateObject("Scripting.FileSystemObject")
Chemin = InputBox("Entrez l'emplacement voulu", "Recherche des sous-dossiers", "C:\Program Files")

If Chemin=Empty Then
	MsgBox "Erreur, il n'y a aucun chemin spécifié", vbCritical
ElseIf SystemeFic.FolderExists(Chemin)=false	Then
	MsgBox "Erreur, le chemin n'existe pas", vbCritical
Else
	ShowSubfolders SystemeFic.GetFolder(Chemin)
End If 

Sub ShowSubFolders(Folder)
    For Each SousDoss in Folder.SubFolders
        Wscript.Echo SousDoss.Name
    Next
End Sub