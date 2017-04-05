FileDelete(@ScriptDir & "\Irbis Hotkeys Documentation.txt")
$scriptFile = FileOpen(@ScriptDir & '\IrbisHotkeys.au3', 0)
$docFile = FileOpen("Irbis Hotkeys Documentation.txt", 1)
If $scriptFile = -1 Or $docFile = -1 Then
    MsgBox(4096, "ERROR", "")
    Exit
EndIf

SrchCln()


Func SrchCln()

Do
	 $Char = FileRead($scriptFile, 1)
	 If @error = -1 Then
		 FileClose($docFile)
		 FileClose($docFile2)
		 FileClose($scriptFile)
		 Exit
	 EndIf
Until $Char = ";"

WriteDoc()
EndFunc

Func WriteDoc()
	$String = FileReadLine($scriptFile)
	If StringInStr($String, "INSTANCE")	Or StringInStr($String, "==>")	Or StringInStr($String, "Run")	Or StringInStr($String, "****") Then
	SrchCln()
	EndIf
	FileWrite($docFile, $String & @CR)
	FileWrite($docFile2, $String & @CR)
	SrchCln()
EndFunc
