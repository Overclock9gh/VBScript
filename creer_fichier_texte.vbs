'Auteur Sa�d Hamdane
'24 septembre 2019
'*****************************************************************************************************************************************************************************************************
'Objectif: Mettre les noms de dossiers dans un fichier texte
'Param�tres d'entr�: fso: l'objet qui utilise les utilit�es de Scripting.FileSystemObject
'                    fichier: le fichier ou on �crit les noms de dossiers
'                    objFic: l'objet du fichier cr�e 
'                    chaineCar:  
'                    nombre
'Param�tres de sortie:
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
	MsgBox "Erreur, le nombre de dossier n'est pas num�rique", vbCritical
End If

objFic.Close