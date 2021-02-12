'Auteur Saïd Hamdane
'24 septembre 2019
'*****************************************************************************************************************************************************************************************************
'Objectif: Créer les dossiers d'un document texte dans le dossier TP1
'Paramètres d'entré: fso: l'objet qui utilise les utilitées de Scripting.FileSystemObject
'                    fichier: le fichier ou on écrit les noms de dossiers
'                    objFic: l'objet du fichier crée 
'                    chaineCar:  
'                    nombre
'Paramètres de sortie:
'*****************************************************************************************************************************************************************************************************

Option Explicit

dim chemin, objFileSys, objReadFile, subCreateFolder, ligne

chemin = "c:\TP1\"

Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objReadFile = objFileSys.OpenTextFile("c:\TP1\dossiers.txt")

Do until objReadFile.AtEndOfStream 
    ligne=CStr(objReadFile.ReadLine())
    If objFileSys.FolderExists(chemin + ligne) Then
    	WScript.Echo "ERREUR, le dossier existe déjà"
    Else
    	objFileSys.Createfolder(chemin + ligne) 
    	WScript.Echo "Le dossier a été créé"
    End If
Loop
objReadFile.Close

