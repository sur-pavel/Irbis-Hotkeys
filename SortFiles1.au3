#include <WinAPISys.au3>

_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)

HotKeySet("^o", "OpenFile")
HotKeySet("^m", "MoveOrRename")
HotKeySet("^p", "AddFile")
HotKeySet("^r", "RenameFile")

HotKeySet("^s", "IrbSave")
HotKeySet("^f", "Search")
HotKeySet("^b", "BriefView")
HotKeySet("^n", "OptimizeView")
HotKeySet("^{SPACE}", "ViewFocus")
HotKeySet("^{F12}", "ScrExit")


$IrbisTit = 'ИРБИС64 - АРМ "Каталогизатор"'
$TotalComTit = "Total Commander 9.0a"
$PdfViewer = "[CLASS:DSUI:PDFXCViewer]"
$GenPath = "d:\TestDir\2\"
$spec = ""


While 1
	Sleep(100)
WEnd


Func OpenFile()
	WinActivate("Total Commander 9.0a")
	If WinWaitActive("Total Commander 9.0a", "", 3) Then
		Send("!{F1}" & StringMid($GenPath, 1, 1))
		Send("{HOME}{F2}")
		MoveToRootFolder()

;~ 		Get file name and parse
		Sleep(100)
		Send("{F2}")
		Send("{HOME}{SHIFTDOWN}" & "^{RIGHT}" & "{SHIFTUP}")
		$clip = ""
;~ 		While (StringInStr($clip, ".pdf") == 0)
;~ 			Send("^c")
		While StringIsInt($clip) <> 1
			Send("^x")
			Sleep(100)
			$clip = ClipGet()
			$clip = StringReplace($clip, ' ', '')
		WEnd
;~ 		$clip = StringSplit($clip, "_")[3]

;~ 		MsgBox(0, "", $clip)
		Send("{ENTER}")
		Sleep(100)
		Send("{ENTER}")
	EndIf

	If WinWaitActive($PdfViewer, "", 3) Then
		WinActivate($IrbisTit)
		If WinWaitActive($IrbisTit, "", 3) Then
			Sleep(100)
			OptimizeView()
			Sleep(100)
			$searchTherm = ControlGetText($IrbisTit, "", "[CLASS:THSHintComboBox]")
			If StringInStr($searchTherm, "Инв. №") == 0 Then
				Srchfor("инв")
			Else
				Send("!d" & "!d")
			EndIf

			ClipPut($clip)
			Send("+{INSERT}")
			Sleep(100)
			Send("{ENTER}")
			Sleep(100)
			Send("{ENTER}")

			Sleep(500)
			GoToField("951")
			Sleep(500)
			Send("!d" & "!d{UP}")
		EndIf
	EndIf

EndFunc   ;==>OpenFile

Func MoveOrRename()
	If WinGetHandle("[ACTIVE]") <> WinGetHandle($TotalComTit) Then
		MoveFile()
	Else
		$panelText1 = ControlGetText($TotalComTit, "", "[CLASS:TMyPanel; INSTANCE:5]")
		$panelText2 = ControlGetText($TotalComTit, "", "[CLASS:TMyPanel; INSTANCE:8]")
		$leftPanel = StringInStr($panelText1, "файлов: 0 из", 0, 1)
		$rightPanel = StringInStr($panelText2, "файлов: 0 из", 0, 1)
;~ 		MsgBox(0, "", $panelText1 & @CRLF & $panelText2 & @CRLF & "Left = " & $leftPanel & @CRLF & "Right = " & $rightPanel)
		If ($leftPanel > 0 And $rightPanel > 0) Then
			MoveFile()
		Else
			HotKeySet("^m")
			Sleep(100)
			Send("^m")
			Sleep(100)
			HotKeySet("^m", "MoveOrRename")
		EndIf
	EndIf
EndFunc   ;==>MoveOrRename

Func MoveFile()
	$clip = ClipGet()
	_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
	$input = InputBox("", "Переместить в:", "", "", 190, 130)
	_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)

	Do
		$exit = 0
;~ 		$count = 1
		$sub = ""
		$spec = ""
		$startName = ""
		$endName = ""
		$cursorMove = 7
		Switch $input
			Case ""
				$folder = ""
				ExitLoop
			Case "5"
				$folder = "5"
				$count = 1
				$cursorMove = 7
				$startName = "5_"
			Case "6"
				$folder = "5"
				$cursorMove = 7
				$count = 1
				$startName = "6_"
			Case "7"
				$folder = "5"
				$cursorMove = 7
				$count = 1
				$spec = "аналог"
				$startName = "7_"
			Case "3"
				$folder = "3"
				$cursorMove = 6
				$count = 1
				$startName = "3_"
			Case "4"
				$folder = "3"
				$cursorMove = 6
				$count = 1
				$startName = "4_"
			Case "д"
				$folder = "ДУБЛЕТНЫЕ"
				$cursorMove = 6
				$count = 1
;~ 				$endName = " НПРК"
			Case "д3"
				$folder = "ДУБЛЕТНЫЕ"
				$cursorMove = 5
				$startName = "3_"
			Case "д4"
				$folder = "ДУБЛЕТНЫЕ"
				$cursorMove = 5
				$startName = "4_"
			Case "п"
				$folder = "Проб"
				$cursorMove = 6
				$count = 1
				$startName = ""
			Case "г"
				$folder = "ПРОБЛЕМНЫЕ"
				$sub = "Нет титул стр ГУГЛа"
				$cursorMove = 10
				$count = 1
				$startName = ""
			Case "к"
				$folder = "3"
				$cursorMove = 6
				$count = 1
				$startName = "3_"
				$spec = "конв"
			Case "с"
				$folder = "КОПИРАЙТ"
			Case "конв"
				$folder = "Конволют"
			Case "ин"
				$folder = "Иностранные"
			Case "ст"
				$folder = "Проблемные"
				$sub = "Статьи из периодики"
				$cursorMove = 10
				$count = 1
;~ 				$startName = "5_"
			Case "др"
				$folder = "Нет"
				$sub = "Другой год"
				$cursorMove = 9
			Case "нео"
				$folder = "Не открываются"
			Case "неом"
				$folder = "Не открываются"
				$sub = "Многотомники"
			Case "вопр"
				$folder = "Под вопросом"
			Case Else
				$exit = 1
				_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
				$input = InputBox("Внимание", "Неправильное значение." & @CRLF & "Повторите.", "", "", 190, 140)
				_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
		EndSwitch

	Until $exit = 0

	If $folder <> "" Then
		WinClose($PdfViewer)
		WinActivate("Total Commander 9.0a")
		If WinWaitActive("Total Commander 9.0a", "", 3) Then
;~ 		right panel to

;~ 			ClipPut($GenPath & $folder & "\")
			Send("!{F2}" & "z")
			Send("{BACKSPACE}")
			Send($folder)
			Send("{ENTER}")
			If $sub <> "" Then
				Send($sub)
				Send("{ENTER}")
			EndIf

;~ 			Send("+{Left}")
;~ 			Sleep(500)
;~ 			Send("+{INSERT}")
;~ 			Sleep(500)
;~ 			For $i = 1 To $count
;~ 				Sleep(100)
;~ 				Send("{ENTER}")
;~ 			Next

;~ 		left panel from
			Send("!{F1}" & "z")
			Send("{HOME}{F2}")
			MoveToRootFolder()
			Send("{F6}")
			If WinWaitActive("[CLASS:TInpComboDlg]", "", 3) Then
				Send("{HOME}")
				For $i = 1 To $cursorMove
					Send("^{RIGHT}")
				Next
;~ 		search underscore
				Send("+{LEFT}" & "^c")
				If (ClipGet() == "_") Then
					Send("{LEFT}{RIGHT}")
				Else
					Send("{RIGHT}")
				EndIf
				Send($startName)
				Send("^{LEFT}")
				If $input == 3 Or $input == 4 Or $input == 7 Or $input == "к" Then
					Send("{SHIFTDOWN}{END}{SHIFTUP}")
					Send("^C")
					Send("{HOME}")
					For $i = 1 To $cursorMove
						Send("^{RIGHT}")
					Next

				Else
					ClipPut($clip)
				EndIf

				If $input == "д" Then
					ClipPut($endName)
					Send("{END}" & "^{LEFT}{LEFT}+{INSERT}")
				EndIf

				If $input == "п" Then
					Send("{SHIFTDOWN}{END}{SHIFTUP}")
					Send("^C")
				EndIf
			EndIf

		EndIf
	EndIf
EndFunc   ;==>MoveFile

Func AddFile()

	WinActivate($IrbisTit)
	If WinWaitActive($IrbisTit, "", 3) Then
		ControlClick($IrbisTit, "", "[CLASS:TTntEdit.UnicodeClass; INSTANCE:1]")
		Send("!q")
		Sleep(100)
		Send(951 & "{ENTER}" & "{F2}")
		$hWnd = WinWaitActive('Элемент: ', "", 5)
		If $hWnd Then
			ControlClick($hWnd, "", "[CLASS:TTntRichEdit.UnicodeClass; INSTANCE:1]", "left", 1, 1, 1)
			Send("^a" & "+{INS}")
			Switch $spec
				Case ""
				Case "конв"
					ClipPut("См. электрон. копию приплет. кн.: ")
					Send("{ENTER 2}")
					Sleep(100)
					Send("+{INS}")
				Case "аналог"
					ClipPut("См. электрон. копию аналога:  г. изд.")
					Send("{ENTER 2}")
					Sleep(100)
					Send("+{INS}")
					For $i = 1 To 4
						Send("^{Left}")
					Next
					Send("{LEFT}")
			EndSwitch

		EndIf

	EndIf
EndFunc   ;==>AddFile

Func RenameFile()
	$clip = ClipGet()
	WinClose($PdfViewer)
	WinActivate("Total Commander 9.0a")
	If WinWaitActive("Total Commander 9.0a", "", 3) Then
		Send("!{F1}" & "z")
		Send("{HOME}{F2}")
		Sleep(500)
		MoveToRootFolder()
		Sleep(500)
		Send("{F2 2}")
		If StringLen($clip) > 8 And StringInStr($clip, "_") > 0 Then
			$clip = StringReplace($clip, '"', '')
			$clip = StringReplace($clip, ':', '.')
			$clip = StringReplace($clip, "_СПб.", "_СПб")
			$clip = StringReplace($clip, "_М.", "_М")
			$clip = StringReplace($clip, ".pdf", "")
			$endName = StringMid($clip, StringLen($clip) - 4, 5)
;~ 			$clip = StringReplace($clip, $endName, StringMid($endName, 1, 4))
;~ 			MsgBox(0, "", $clip & "!")
		EndIf
		ClipPut($clip)
	EndIf

EndFunc   ;==>RenameFile

Func IrbSave()
	If HotKeyOn("^s", "IrbSave") Then
		Send("+{ENTER}")
	EndIf
EndFunc   ;==>IrbSave

Func HotKeyOn($send, $func)
	If WinGetHandle("[ACTIVE]") == WinGetHandle($IrbisTit) Then
		ControlClick($IrbisTit, "", "[CLASS:TTntStringGrid.UnicodeClass; INSTANCE:3]")
		Return True
	Else
		HotKeySet($send)
		Sleep(100)
		Send($send)
		Sleep(100)
		HotKeySet($send, $func)
		Return False
	EndIf
EndFunc   ;==>HotKeyOn

Func ViewFocus()
	If HotKeyOn("^{SPACE}", "ViewFocus") Then
		Sleep(100)
		ControlClick($IrbisTit, "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "left", 1, 1034, 25)
	EndIf
EndFunc   ;==>ViewFocus

Func Search()
	If WinGetHandle("[ACTIVE]") <> WinGetHandle($TotalComTit) Then
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
		$input = InputBox("Выполнить", "Поиск по:", "", "", 190, 130)
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)

		WinActivate($IrbisTit)
		If WinWaitActive($IrbisTit, "", 5) Then
			Do
				$exit = 0
				$inputTest = TestInput($input)
				If IsArray($inputTest) Then
					$SPLIT = $inputTest
					$input = $SPLIT[1]
				Else
					$input = $inputTest
					$SPLIT = 0
				EndIf
				; 			1. Если введенная строка - число, открытие записи в Ирбисе по инв. номеру
				If StringIsInt($input) Then
					Srchfor("инв")
					ClipMan($input)
					Send("{ENTER}")
				Else

					; 			2. Поиск по виду основного словаря
					Switch $input
						Case "" ; пустое поле - закрытие окна
							ExitLoop
						Case "тех" ; "тех" - по технологии
							Srchfor("тех")
						Case "изд" ; "изд" - по изд-ву
							Srchfor("изд")
						Case "инв" ; "инв" - по инв. номеру
							Srchfor("инв")
						Case "хар" ; "хар" - по характеру документа
							Srchfor("хар")
						Case "год" ; "год" - по году издания
							Srchfor("год")
						Case "тек" ; "тек" - открыть текущую запись отдельно
							$invNum = GetInvNum()
							;**** открытие записи по инв. номеру
							Srchfor("инв")
							ClipMan($invNum)
							Send("{ENTER}")
						Case "мн" ; "мн" - все многотомники
							Srchfor("вид")
							Send("03" & "{ENTER}")
						Case "авт" ; "авт" - по автору
							SrchforEx("автор (э", $SPLIT)
						Case "под" ; "под" - по предметному подзаголовку
							SrchforEx("пред", $SPLIT)
						Case "руб" ; "руб" - по предметной рубрике
							SrchforEx("предметные руб", $SPLIT)
						Case "заг" ; "заг" - по заглавию
							SrchforEx($input, $SPLIT)
						Case "кл" ; "кл" - по ключевому слову
							SrchforEx($input, $SPLIT)
						Case "перс" ; "перс" - по персоналии
							SrchforEx($input, $SPLIT)
						Case Else ; неправильное значение заново открывает диалоговое окно ввода
							$exit = 1
							_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
							$hWnd1 = WinWaitActive($IrbisTit, "", 5)
							If $hWnd1 Then
								$input = InputBox("Внимание", "Неправильное значение." & @CRLF & "Повторите.", "", "", 190, 140)
								_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
							EndIf
					EndSwitch
				EndIf
			Until $exit = 0
		EndIf
	Else
		HotKeySet("^f")
		Sleep(100)
		Send("^f")
		Sleep(100)
		HotKeySet("^f", "Search")
	EndIf
EndFunc   ;==>Search

Func TestInput($input)
;~ 		Размер строки длиннее 100 символов
	If StringLen($input) > 100 Then
		$SPLIT = "error"
	ElseIf StringInStr($input, ' ') <> 0 Then
		$SPLIT = StringSplit($input, ' ')
;~ 			Введено больше 10 слов
		If $SPLIT[0] > 10 Then
			$SPLIT = 0
		Else
			For $i = 1 To $SPLIT[0]
;~ 				Одно из слов длиннее 20 символов
				If StringLen($SPLIT[$i]) > 20 Then
					$SPLIT = 0
					ExitLoop
				EndIf
			Next
		EndIf
	Else
		$SPLIT = $input
	EndIf
	Return $SPLIT
EndFunc   ;==>TestInput

Func Srchfor($srch)
	Send("!f")
	If WinWaitActive("Вид основного словаря", "", 5) Then
		ClipMan($srch)
		Send("{ENTER}")
		If WinWaitActive($IrbisTit, "", 5) Then
			EndOfSearch()
		EndIf
	EndIf

EndFunc   ;==>Srchfor

;~ Изменение вида словаря с вставкой последующих за командой строк
Func SrchforEx($srch, $SPLIT)
	Send("!f")
	If WinWaitActive("Вид основного словаря", "", 5) Then
		ClipMan($srch)
		Send("{ENTER}")
		If WinWaitActive($IrbisTit, "", 5) Then
			EndOfSearch()
			Sleep(100)

			If $SPLIT <> 0 Then
				$string = ""

				For $i = 2 To $SPLIT[0]
					If $i == $SPLIT[0] Then
						$string = $string & $SPLIT[$i]
					Else
						$string = $string & $SPLIT[$i] & ' '
					EndIf
				Next

				Sleep(100)
				ClipMan($string)
				Send("{ENTER}")
			EndIf
		EndIf
	EndIf

EndFunc   ;==>SrchforEx

Func EndOfSearch()
	Sleep(300)
	ControlSend($IrbisTit, "", "[CLASS:THSHintTntComboBox.UnicodeClass; INSTANCE:1]", "{HOME}" & "{DOWN}")
	Sleep(300)
	ControlSend($IrbisTit, "", "[CLASS:THSHintTntComboBox.UnicodeClass; INSTANCE:1]", "{DOWN}")
	$controlId = ""
	While (StringInStr($controlId, "TTntEdit.Unicode") == 0)
		Send("!d")
		$controlId = ControlGetFocus($IrbisTit)
	WEnd
	Send("!d")
EndFunc   ;==>EndOfSearch
;~ Использование буфера обмена с сохранением текущего состояния буфера
Func ClipMan($com)
	$clip = ClipGet()
	Sleep(100)
	ClipPut($com)
	Send("+{INS}")
	Sleep(100)
	ClipPut($clip)
EndFunc   ;==>ClipMan

Func GetInvNum()
	WinActivate($IrbisTit)
	$hWnd = WinWaitActive($IrbisTit, "", 5)
	If $hWnd Then
		GoToField(910)
		$text = WinGetText($IrbisTit, "")
		$pos1 = StringInStr($text, "^B", 0, 1) + 2
		$pos2 = StringInStr($text, "^C", 0, 1)
		$invNumLen = $pos2 - $pos1
		$invNum = (StringMid($text, $pos1, $invNumLen))
		Return $invNum
	EndIf
EndFunc   ;==>GetInvNum

Func GoToField($com)
	Send("!q")
	Sleep(100)
	Send($com & "{ENTER}")
EndFunc   ;==>GoToField

Func MoveToRootFolder()
	ClipPut($GenPath)
	Send("+{INSERT}{ENTER}" & "{END}")
EndFunc   ;==>MoveToRootFolder

Func ScrExit()
	Exit
EndFunc   ;==>ScrExit

Func BriefView()
	If HotKeyOn("^b", "BriefView") Then
		For $i = 1 To 3
			$controlText = ControlGetText($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:" & $i & "]")
			If StringInStr($controlText, "Оптимизированный") > 0 Then
				ControlSend($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:" & $i & "]", "{END}{UP 2}{ENTER}")
			EndIf
		Next
	EndIf
EndFunc   ;==>BriefView

Func OptimizeView()
	If HotKeyOn("^n", "OptimizeView") Then
		For $i = 1 To 3
			$controlText = ControlGetText($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:" & $i & "]")
			If StringInStr($controlText, "BRIEF_F - ДЛЯ ИМЕНИ ФАЙЛА") > 0 Then
				ControlSend($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:" & $i & "]", "{END}{ENTER}")
			EndIf
		Next
	EndIf
EndFunc   ;==>OptimizeView

