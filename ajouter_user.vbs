'Auteur Saïd Hamdane
'24 septembre 2019
'*****************************************************************************************************************************************************************************************************
'Objectif: Changer les permissions de user lulu et les Administrateurs en control total sur le dossier C:\TP1
'Paramètres d'entré: fso: l'objet qui utilise les utilitées de Scripting.FileSystemObject
'                    fichier: le fichier ou on écrit les noms de dossiers
'                    objFic: l'objet du fichier crée 
'                    chaineCar:  
'                    nombre
'Paramètres de sortie:
'*****************************************************************************************************************************************************************************************************

Option Explicit

Dim strUser, strUser2, strPassword, objReseau, strOrdinateur, colComptes, objUser

'Declaration
strUser = "admin2"
strUser2 = "admin3 "
strPassword = "password44"

'Computer Name
Set objReseau = WScript.CreateObject("WScript.Network")
strOrdinateur = objReseau.ComputerName

'Creation of the local user1
Set colComptes = GetObject("WinNT://" & strOrdinateur & "" )
Set objUser = colComptes.Create("user", strUser)
objUser.SetPassword strPassword
objUser.SetInfo

'Creation of the local user2
Set colComptes = GetObject("WinNT://" & strOrdinateur & "" )
Set objUser = colComptes.Create("user", strUser2)
objUser.SetPassword strPassword
objUser.SetInfo



