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

Dim strFolder, strUser, strUser2, strCommand, strCommand2, objShellExec, strOutput, strOutput2, objShellExec2

strFolder = "C:\TP1" 
strUser = "HAMDANE\lulu"
strUser2 = "Administrateurs"
	 
SetPermissions
	     
Function SetPermissions()

	Const WshFinished = 1, WshFailed = 2
	Dim objShell, objFSO
	strCommand = "icacls " + strFolder + " /grant " + strUser + ":(OI)(CI)F /T"
	strCommand2 = "icacls " + strFolder + " /grant " + strUser2 + ":(OI)(CI)F /T"
	 
	Set objShell = CreateObject("Wscript.Shell")
	Set objShellExec = objShell.Exec(strCommand)
	Set objShellExec2 = objShell.Exec(strCommand2)

msgbox objShellExec.StdOut.ReadAll
msgbox objShellExec2.StdOut.ReadAll

End Function

