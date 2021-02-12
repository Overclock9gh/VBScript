'Auteur Sa�d Hamdane
'24 septembre 2019
'*****************************************************************************************************************************************************************************************************
'Objectif: Cr�er les dossiers d'un document texte dans le dossier TP1
'Param�tres d'entr�: fso: l'objet qui utilise les utilit�es de Scripting.FileSystemObject
'                    fichier: le fichier ou on �crit les noms de dossiers
'                    objFic: l'objet du fichier cr�e 
'                    chaineCar:  
'                    nombre
'Param�tres de sortie:
'*****************************************************************************************************************************************************************************************************

Option Explicit

dim chemin, objFileSys, objReadFile, subCreateFolder, ligne

chemin = "c:\TP1\"

Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objReadFile = objFileSys.OpenTextFile("c:\TP1\dossiers.txt")

Do until objReadFile.AtEndOfStream 
    ligne=CStr(objReadFile.ReadLine())
    If objFileSys.FolderExists(chemin + ligne) Then
    	WScript.Echo "ERREUR, le dossier existe d�j�"
    Else
    	objFileSys.Createfolder(chemin + ligne) 
    	WScript.Echo "Le dossier a �t� cr��"
    End If
Loop
objReadFile.Close

