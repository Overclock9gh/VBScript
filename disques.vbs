'Auteur Saïd Hamdane
'24 septembre 2019

'Objectif: Afficher tous les lecteurs disques

Option Explicit

Dim objSysFic, disques, objDisque

Set objSysFic = CreateObject("Scripting.FileSystemObject")
Set disques = objSysFic.Drives

For Each objDisque in disques
    Wscript.Echo "Disque: " & objDisque.DriveLetter
Next