'Auteur Sa�d Hamdane
'24 septembre 2019
'*****************************************************************************************************************************************************************************************************
'Objectif: Changer les permissions de user lulu et les Administrateurs en control total sur le dossier C:\TP1
'Param�tres d'entr�: fso: l'objet qui utilise les utilit�es de Scripting.FileSystemObject
'                    fichier: le fichier ou on �crit les noms de dossiers
'                    objFic: l'objet du fichier cr�e 
'                    chaineCar:  
'                    nombre
'Param�tres de sortie:
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



