'Auteur Saïd Hamdane
'24 septembre 2019

'Objectif: Afficher un fichier choisis par l'utilisateur dans le répértoire courant

Option Explicit

Dim objfso, interpreteur, strfichier, CurrentDirectory, WshShell

Set objfso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject ("WScript.Shell")

currentDirectory = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

interpreteur = InStr (1, WScript.FullName, "cscript", VbTextCompare)
    If interpreteur > 0 Then
        InputBox"Entrer le nom d'un fichier", strfichier
        
        If objfso.FileExists(objfso.BuildPath(CurrentDirectory, strfichier)) Then
            
            Set objTextFile = objfso.OpenTextFile _
            (strfichier, ForReading)

            Do Until objTextFile.AtEndOfStream
            strComputer = objTextFile.ReadLine
             Wscript.Echo strComputer
            Loop

        Else
            WScript.Echo "Le fichier n'existe pas"
        End if

    Else
	    MsgBox "Erreur, le programme ne permet pas l'exécution en wscript"
    End If
