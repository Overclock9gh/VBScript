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