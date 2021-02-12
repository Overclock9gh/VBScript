'Auteur Saïd Hamdane
'24 septembre 2019
'*****************************************************************************************************************************************************************************************************
'Objectif: Mettre les noms de dossiers dans un fichier texte
'Paramètres d'entré: fso: l'objet qui utilise les utilitées de Scripting.FileSystemObject
'                    fichier: le fichier ou on écrit les noms de dossiers
'                    objFic: l'objet du fichier crée 
'                    chaineCar:  
'                    nombre
'Paramètres de sortie:
'*****************************************************************************************************************************************************************************************************

Option Explicit

Dim fso, fichier, objFic, chaineCar, nombre, i

Set fso=CreateObject("Scripting.FileSystemObject")

fichier="c:\TP1\dossiers.txt"
Set objFic = fso.CreateTextFile(fichier,True)
nombre = InputBox("Entrez le nombre de dossier")

If IsNumeric(nombre) Then
	For i=1 To nombre  
		chaineCar = InputBox("Entrez le nom du dossier")
		objFic.WriteLine chaineCar
	Next	
Else
	MsgBox "Erreur, le nombre de dossier n'est pas numérique", vbCritical
End If

objFic.Close