'** Script Information *******************************************************************
'** Description	: Test quickly the CustomSettings.ini file								**
'** Author(s)	: Florian Valente														**
'** Date		: April 23, 2013														**
'** Version		: 1.0																	**
'*****************************************************************************************

Option Explicit

'** Variables Declaration **
Dim oWSH					: Set oWSH = CreateObject("WScript.Shell")
Dim oFSO					: Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim strMININTPath			: strMININTPath = "C:\MININT\SMSOSD\OSDLOGS\"

' Delete VARIABLES.DAT file
If (oFSO.FileExists(strMININTPath & "VARIABLES.DAT")) Then oFSO.DeleteFile strMININTPath & "VARIABLES.DAT", True

' Check if the script is located in the DS Scripts folder
If (oFSO.FolderExists("..\Control\")) Then
	oWSH.Run "ZTIGather.wsf /inifile:..\Control\CustomSettings.ini"
Else
	MsgBox "Please place this script in the DS Scripts folder."
	WScript.Quit
End If

' Launch notepad with BDD.LOG file
oWSH.Run "notepad.exe " & strMININTPath & "BDD.LOG", 1, False

Set oWSH = Nothing
Set oFSO = Nothing