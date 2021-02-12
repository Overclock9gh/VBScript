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

Dim strOrdinateur, objWMISer, colUsers, objItem, obj, intUAC
strOrdinateur  = "."

Set objWMISer = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strOrdinateur  & "\root\cimv2")

Set colUsers = objWMISer.ExecQuery _
    ("Select * from Win32_UserAccount Where LocalAccount = True")

For Each objItem in colUsers

	Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Full Name: " & objItem.FullName
    Wscript.Echo "Password Expires: " & objItem.PasswordExpires
    Wscript.Echo
    
    If objItem.PasswordExpires Then
    	objItem.PasswordExpires = False
    	objItem.Put_
    Else 
    	objItem.PasswordExpires = True 
    	objItem.Put_
    End if 	
    
    Wscript.Echo "Password Expires: " & objItem.PasswordExpires
    
Next