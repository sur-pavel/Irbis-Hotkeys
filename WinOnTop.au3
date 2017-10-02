#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Users\Андрей\Downloads\Icons8-Windows-8-Programming-Pin.ico
#AutoIt3Wrapper_Outfile=..\WinOnTop.Exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
HotKeySet("^{F8}", "OnTop")
HotKeySet("^{F9}", "OnTopOff")

While 1
	Sleep(100)
WEnd


Func OnTop()
	Sleep(10)
	Send("{CTRLDOWN}")
	Sleep(10)
	Send("{CTRLUP}")
	Local $hWnd = WinGetHandle("[ACTIVE]")
	WinSetOnTop($hWnd, "", 1)
EndFunc   ;==>OnTop


Func OnTopOff()
	Sleep(10)
	Send("{CTRLDOWN}")
	Sleep(10)
	Send("{CTRLUP}")

	Local $hWnd = WinGetHandle("[ACTIVE]")
	WinSetOnTop($hWnd, "", 0)
EndFunc   ;==>OnTopOff
