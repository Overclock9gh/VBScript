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

Dim oShell, sPath

Set oShell = WScript.CreateObject ("WScript.Shell")
oShell.Run("net use s: \\HAMDANE\scripts")

WScript.Sleep 5000

sPath = "s:"
Set oShell = CreateObject("WScript.Shell")
oShell.Run "explorer /n," & sPath, 1, False

CreateObject("WScript.Shell").Run("https://cmontmorency.moodle.decclic.qc.ca/")



