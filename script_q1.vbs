'Auteur Saïd Hamdane
'24 septembre 2019

'Objectif: Afficher un message d'encouragement avec un nom donné 

Option Explicit

Dim interpreteur, strnom

strnom = "Bob"

If strnom=empty Then
    MsgBox "Erreur, il n'y a aucun nom en argument"
Else
interpreteur = InStr (1, WScript.FullName, "cscript", VbTextCompare)
    If interpreteur > 0 Then
	    WScript.Echo "Bon examen " + strnom + "!"
    Else
	    MsgBox "Bon examen " + strnom + "!"
    End If
End if

