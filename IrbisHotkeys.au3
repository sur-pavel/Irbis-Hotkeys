#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=IrbisHotkeys.ico
#AutoIt3Wrapper_Outfile=IrbisHotkeys.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <WinAPISys.au3>
#include <Word.au3>
#include <Misc.au3>
#include <Date.au3>
#include <MsgBoxConstants.au3>

;
;
; 					Горячие клавиши для работы в АРМ "Каталогизатор" (ИРБИС64)
;


_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
HotKeySet("^v", "Vstavka")
HotKeySet("^s", "IrbSave")
HotKeySet("^z", "SearchNumbs")
HotKeySet("^q", "Field")
HotKeySet("^f", "Search")
HotKeySet("!h", "CohSearch")
HotKeySet("^w", "Osn")
HotKeySet("^k", "KK")
HotKeySet("^d", "Formular")
HotKeySet("^y", "Label")
HotKeySet("^m", "FormularLabel")
HotKeySet("^{F8}", "OnTop")
HotKeySet("^{F9}", "OnTopOff")
HotKeySet("^+g", "Obrzv")
HotKeySet("^{F12}", "ScrExit")
HotKeySet("^{SPACE}", "ViewFocus")

$IrbisTit = 'ИРБИС64 - АРМ "Каталогизатор"'


While 1
	Sleep(100)
	;**** Выход, если запущено больше одного экземпляра программы
	$countProc = ProcessList("IrbisHotkeys.exe")
	If $countProc[0][0] > 1 Then
		Exit
	EndIf
WEnd

;						CTRL+F12 Выход из скрипта
Func ScrExit()
	Exit
EndFunc   ;==>ScrExit


Func Obrzv()
	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^+g")
		Send("^+g")
		HotKeySet("^+g", "Obrzv")
	Else
		$glW = "Глобальная корректировка БД"
		ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:4]", "left", 1, 323, 11)
		$gWnd = WinWaitActive($glW, "", 5)
		If Not $gWnd Then
			MsgBox(4096, 'Сообщение', 'Окно не найдено, завершаем работу скрипта')
		EndIf
		If $gWnd Then
			Sleep(100)
			ControlClick($glW, "", "[CLASS:TCheckBox; INSTANCE:5]")
			ControlClick($glW, "", "[CLASS:TToolBar; INSTANCE:3]", "left", 1, 101, 11)
		EndIf

		$gWnd = WinWaitActive("Выбор", "", 5)
		If $gWnd Then
			ControlClick("Выбор", "", "[CLASS:TBitBtn; INSTANCE:2]")
		EndIf

		$gWnd = WinWaitActive("Открытие", "", 5)
		If $gWnd Then
			Send("{!}{!}КР-ФЛК" & "{DOWN}" & "{ENTER}")
			ControlClick($glW, "", "[CLASS:TButton; INSTANCE:5]")
		EndIf

		$gWnd = WinWaitActive("Глобальная корректировка БД MPDA", "", 5)
		If $gWnd Then
			Sleep(100)
			ControlClick("Глобальная корректировка БД MPDA", "", "[CLASS:TBitBtn; INSTANCE:3]")
		EndIf


	EndIf
EndFunc   ;==>Obrzv


;						CTRL+V Вставка. Вставляет в поле данные без раскрытия окна "Элемент"
Func Vstavka()
	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^v")
		Send("^v")
		HotKeySet("^v", "Vstavka")
	Else
		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
		Send("+{INS}")
	EndIf
EndFunc   ;==>Vstavka

;						CTRL+S Сохранение записи в Ирбисе
Func IrbSave()
	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^s")
		Send("^s")
		HotKeySet("^s", "IrbSave")
	Else
		Send("+{ENTER}")
		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
	EndIf
EndFunc   ;==>IrbSave

;						CTRL+Z Вывод нескольских инв. номеров. Ввести инв. номера и нажать TAB
Func SearchNumbs()


	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^z")
		Send("^z")
		HotKeySet("^z", "SearchNumbs")

	Else
		ControlFocus($IrbisTit, "", "[CLASS:TTntStringGrid.UnicodeClass; INSTANCE:3]")
		$bdName = ControlGetText($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:4]")
		$isPRBD = StringInStr($bdName, "PR - Периодические издания (с 2014 г.)")
		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
		Send("!z")


		$hWnd = WinWaitActive("Поиск по словарю/Рубрикатору", "", 5)
		If $hWnd Then
			If $isPRBD = 0 Then
				ControlSend("Поиск по словарю/Рубрикатору", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{HOME}")
				Sleep(100)
				ControlSend("Поиск по словарю/Рубрикатору", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{DOWN}")
			Else
				ControlSend("Поиск по словарю/Рубрикатору", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{HOME}")
				For $i = 4 To 1 Step -1
					Sleep(100)
					ControlSend("Поиск по словарю/Рубрикатору", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{DOWN}")
				Next
			EndIf
			Sleep(100)
			ControlClick("Поиск по словарю/Рубрикатору", "", "[CLASS:TGroupButton; INSTANCE:6]")
			ControlClick("Поиск по словарю/Рубрикатору", "", "[CLASS:TTntEdit.UnicodeClass; INSTANCE:1]")



			;**** Завершение отбора на кнопку TAB
			Local $hDLL = DllOpen("user32.dll")
			While 1
				If _IsPressed("12", $hDLL) Then
					Sleep(500)
				ElseIf _IsPressed("09", $hDLL) Then

					$hWnd1 = WinWaitActive("Поиск по словарю/Рубрикатору", "", 5)
					If $hWnd1 Then
						ControlClick("Поиск по словарю/Рубрикатору", "", "[CLASS:TBitBtn; INSTANCE:3]")
						Sleep(500)
						ControlClick("Поиск по словарю/Рубрикатору", "", "[CLASS:TBitBtn; INSTANCE:1]")
					EndIf
					$hWnd1 = WinWaitActive($IrbisTit, "", 5)
					If $hWnd1 Then
						ControlSend($IrbisTit, "", "[CLASS:THSHintTntComboBox.UnicodeClass; INSTANCE:1]", "{HOME}" & "{DOWN}")
					EndIf
					ExitLoop
				ElseIf _IsPressed("1B", $hDLL) Then
					ExitLoop
				EndIf
			WEnd

			DllClose($hDLL)
		EndIf
	EndIf
EndFunc   ;==>SearchNumbs
;

;						CTRL+Q Разные команды. Читает введенную строку и выполняет команды
Func Field()


	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^q")
		Send("^q")
		HotKeySet("^q", "Field")
	Else

		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
		$input = InputBox("Выполнить", "Название поля:", "", "", 190, 130)
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)

		ControlFocus($IrbisTit, "", "[CLASS:TTntStringGrid.UnicodeClass; INSTANCE:3]")


		;**** Пути до рубрик
		$sPath_ini = @ScriptDir & "\IrbisHotkeys.ini"
		$rubDir = IniRead($sPath_ini, "Sec1", "RubDir", "d:\Рубрики по уровням\")

		; 			1) Если введенная строка - число
		If StringIsInt($input) Then
			; Переход по номеру поля, если число меньше 4
			If StringLen($input) < 4 Then
				Send("!q" & $input & "{ENTER}")
			Else
				; Открытие записи в Ирбисе по инв. номеру, если число больше 3
				Send("!f" & "инв" & "{ENTER}")
				Sleep(100)
				Send("!d" & $input & "{ENTER}")
			EndIf
		Else
			; 			2) Переход по полям. Вид: введенная строка - выполняемая команда (открытие полей, файлов и проч.)

			Do
				$exit = 0
				WinActivate("Внимание")
				ControlFocus("Внимание", "", "[CLASS:Edit; INSTANCE:1]")
				Switch $input

					Case "" ; Пустая строка - закрытие окна
						ExitLoop

					Case "хар" ; "хар" - характер документа
						Send("!q" & 900 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "900', "", 5)
						If $hWnd Then
							Send("{DOWN 3}")
						EndIf
					Case "сер" ; "сер" - заглавие серии
						Send("!q" & 225 & "{ENTER}")
					Case "инд" ; "инд" - индекс МДА
						Send("!q" & 686 & "{ENTER}")
					Case "аз" ; "аз" - авторский знак
						Send("!q" & 908 & "{ENTER}" & "{F2}")
					Case "кат" ; "кат" - поле "Каталогизатор"
						Send("!q" & 907 & "{ENTER}")
					Case "анн" ; "анн" - аннотация
						Send("!q" & 331 & "{ENTER}" & "{F2}")
					Case "заг" ; "заг" - заглавие однотомника
						Send("!q" & 200 & "{ENTER}" & "{F2}")
					Case "авт" ; "авт" - автор однотомника
						Send("!q" & 700 & "{ENTER}" & "{F2}")
					Case "давт" ; "давт" - другие авторы однотомника
						Send("!q" & 701 & "{ENTER}" & "{F2}")
					Case "ред" ; "ред" - редакторы однотомника
						Send("!q" & 702 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "702', "", 5)
						If $hWnd Then
							Send("{ENTER}")
						EndIf
					Case "кол" ; "кол" - 700 поле, другие коллективы
						Send("!q" & 711 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "711', "", 5)
						If $hWnd Then
							Send("{ENTER}")
						EndIf
					Case "мес" ; "мес" - место издание однотомника
						Send("!q" & 210 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "210', "", 5)
						If $hWnd Then
							Send("{PGDN}" & "{UP 9}")
						EndIf
					Case "изде" ; "изде" - сведения об издании
						Send("!q" & 205 & "{ENTER}" & "{F2}")
					Case "изд" ; "изд" - издательство однотомника
						Send("!q" & 210 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "210', "", 5)
						If $hWnd Then
							Send("{DOWN 3}")
						EndIf
					Case "тип" ; "тип" - типография однотомника
						Send("!q" & 210 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "210', "", 5)
						If $hWnd Then
							Send("{PGDN}" & "{UP}")
						EndIf
					Case "год" ; "год" - год издания однотомника
						Send("!q" & 210 & "{ENTER}" & "{F2}")
					Case "отв" ; "отв" - сведения об ответственности однотомника
						Send("!q" & 200 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "200', "", 5)
						If $hWnd Then
							ControlFocus('Элемент: "200', "", "[CLASS:TTntRichEdit.UnicodeClass; INSTANCE:1")
							Send("{PGDN}")
							Sleep(200)
							Send("{UP}")
						EndIf
					Case "свед" ; "свед" - сведения к заглавию однотомника
						Send("!q" & 200 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "200', "", 5)
						If $hWnd Then
							Send("{PGDN}")
							Sleep(200)
							Send("{UP 2}")
						EndIf
					Case "прим" ; "прим" - общие примечания однотомника
						Send("!q" & 300 & "{ENTER}")
					Case "супо" ; "супо" - вставка в общ. примеч. "Изд. в суперобл."
						Send("!q" & 300 & "{ENTER}" & "Изд. в суперобл.")
					Case "разн" ; "разн" - разночтение заглавий
						Send("!q" & 517 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "517', "", 5)
						If $hWnd Then
							Send("{DOWN 2}")
						EndIf
					Case "руб" ; "руб" - предметная рубрика
						Send("!q" & 606 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "606', "", 5)
						If $hWnd Then
							Send("{DOWN}")
						EndIf
					Case "перс" ; "перс" - персоналия
						Send("!q" & 600 & "{ENTER}" & "{F2}")
					Case "стр" ; "стр" - количеств. хар-ки
						Send("!q" & 215 & "{ENTER}" & "{F2}")
					Case "мзаг" ; "мзаг" - заглавие многотомника
						Send("!q" & 461 & "{ENTER}" & "{F2}")
					Case "мавт" ; "мавт" - автор многотомника или первый редактор
						Send("!q" & 961 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "961', "", 5)
						If $hWnd Then
							Send("{DOWN 7}")
						EndIf
					Case "мизд" ; "мизд" - изд-во многотомника
						Send("!q" & 461 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "461', "", 5)
						If $hWnd Then
							Send("{ENTER 6}")
						EndIf
					Case "мгод" ; "мгод" - год начала издания многотомника
						Send("!q" & 461 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "461', "", 5)
						If $hWnd Then
							Send("{ENTER 12}")
						EndIf
					Case "мотв" ; "мотв" - сведения об ответственности многотомника
						Send("!q" & 461 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "461', "", 5)
						If $hWnd Then
							Sleep(100)
							ControlFocus('Элемент: "461', "", "[CLASS:TTntRichEdit.UnicodeClass; INSTANCE:1")
							Send("{ENTER 4}")
						EndIf
					Case "мсвед" ; "мсвед" - сведения к заглавию многотомника
						Send("!q" & 461 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "461', "", 5)
						If $hWnd Then
							Send("{ENTER 3}")
						EndIf
					Case "мприм" ; "мприм" - общие примечания многотомника
						Send("!q" & 46 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "46:', "", 5)
						If $hWnd Then
							Send("{ENTER 8}")
						EndIf
					Case "эб" ; "эб" - 830 поле
						Send("!q" & 830 & "{ENTER}")
					Case "инв" ; "инв" - первый инв. номер
						Send("!q" & 910 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "910', "", 5)
						If $hWnd Then
							Send("{DOWN}")
						EndIf
					Case "осо" ; "осо" - отключение сведений об ответственности
						Send("!q" & 905 & "{ENTER}")
						ClipPut("^11")
						Send("+{INS}")

					Case "бсз" ; "бсз" - сведения к заглавию с прописной
						Send("!q" & 905 & "{ENTER}")
						ClipPut("^23")
						Send("+{INS}")

					Case "уни" ; "уни" - описание переделывается в "ЗАПОЛНИТЬ НОВОЙ ЗАПИСЬЮ"
						Send("!q" & 503 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "503', "", 5)
						If $hWnd Then
							Send("{F2}")
							$hWnd1 = WinWaitActive("Подполе", "", 5)
							If $hWnd1 Then
								Send("{DOWN 4}" & "{ENTER}")
							EndIf
							$hWnd1 = WinWaitActive('Элемент: "503', "", 5)
							If $hWnd1 Then
								ControlClick('Элемент: "503', "", "[CLASS:TBitBtn; INSTANCE:2]")
							EndIf
							$hWnd1 = WinWaitActive($IrbisTit, "", 5)
							If $hWnd1 Then
								Send("!q" & 910 & "{ENTER}" & "^{ENTER}")
								ClipPut("^A0^B000-6")
								Send("+{INS}")
							EndIf
						EndIf
					Case "пп" ; "пп" - отметка о плохом переплете
						Send("!q" & 910 & "{ENTER}")
						$text = WinGetText($IrbisTit, "")
						$pos1 = StringInStr($text, "^B", 0, 1) + 2
						$pos2 = StringInStr($text, "^C", 0, 1)
						$invNumLen = $pos2 - $pos1
						$invNum = (StringMid($text, $pos1, $invNumLen))
						$hWnd1 = WinWaitActive($IrbisTit, "", 5)
						If $hWnd1 Then
							Send("!q" & 141 & "{ENTER}" & "{F2}")
						EndIf
						$hWnd = WinWaitActive('Элемент: "141', "", 5)
						If $hWnd Then
							Send("{PGDN}")
							Sleep(100)
							Send($invNum)
							Send("{UP 4}")
							Sleep(100)
							Send("{F2}")
							$hWnd1 = WinWaitActive("Подполе", "", 5)
							If $hWnd1 Then
								Send("{DOWN 3}" & "{ENTER}")
							EndIf
							$hWnd1 = WinWaitActive('Элемент: "141', "", 5)
							If $hWnd1 Then
								ControlClick('Элемент: "141', "", "[CLASS:TBitBtn; INSTANCE:2]")
							EndIf
						EndIf

					Case "библ" ; "библ" - примеч. о библиографии
						Send("!q" & 320 & "{ENTER}")
					Case "шифр" ; "шифр" - шифр док-та в БД
						Send("!q" & 903 & "{ENTER}")
					Case "сод" ; "сод" - открытие поля "Содержание" в виде таблицы и переход на заглавие
						Send("!q" & 330 & "{ENTER}" & "{F3}")
						$hWnd = WinWaitActive('Элемент:  "330: Содержание (оглавление)', "", 5)
						If $hWnd Then
							Send("{ENTER 33} {LEFT 4}")
						EndIf
					Case "пер" ; "пер" - заглавие оригинала переводного издания
						Send("!q" & 454 & "{ENTER}" & "{F2}")
					Case "пар" ; "пар" - параллельное заглавие
						Send("!q" & 510 & "{ENTER}" & "{F2}")
					Case "кл" ; "кл" - ключевые слова
						Send("!q" & 610 & "{ENTER}")
					Case "ис" ; "ис" - ISBN однотомника
						Send("!q" & 10 & "{ENTER}" & "{F2}")
					Case "мис" ; "мис" - ISBN многотомника
						Send("!q" & 461 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "461', "", 5)
						If $hWnd Then
							Send("{PGDN}" & "{UP 10}")
						EndIf
					Case "пазк" ; "пазк" - переделывает описание в PAZK
						Send("!q" & 900 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "900', "", 5)
						If $hWnd Then
							Send("{DOWN}" & "{F2}")
							$hWnd1 = WinWaitActive("Подполе", "", 5)
							If $hWnd1 Then
								Send("{HOME}" & "{DOWN 4}" & "{ENTER}")
							EndIf
							$hWnd1 = WinWaitActive('Элемент: "900', "", 5)
							If $hWnd1 Then
								Send("{TAB}" & "{ENTER}")
							EndIf
							$hWnd1 = WinWaitActive($IrbisTit, "", 5)
							If $hWnd1 Then
								Send("!q" & 920 & "{ENTER}" & "{F2}")
							EndIf
							$hWnd1 = WinWaitActive('Элемент: "920', "", 5)
							If $hWnd1 Then
								Send("{HOME}" & "{ENTER}")
							EndIf
							$hWnd1 = WinWaitActive($IrbisTit, "", 5)
							If $hWnd1 Then
								ControlSend($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:2]", "{HOME}")
							EndIf
						EndIf
					Case "пвк" ; "пвк" - переделывает описание в PVK
						Send("!q" & 900 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "900', "", 5)
						If $hWnd Then
							Send("{DOWN}" & "{F2}")
							$hWnd1 = WinWaitActive("Подполе", "", 5)
							If $hWnd1 Then
								Send("{HOME}" & "{DOWN 4}" & "{ENTER}")
							EndIf
							$hWnd1 = WinWaitActive('Элемент: "900', "", 5)
							If $hWnd1 Then
								Send("{TAB}" & "{ENTER}")
							EndIf
							$hWnd1 = WinWaitActive($IrbisTit, "", 5)
							If $hWnd1 Then
								Send("!q" & 920 & "{ENTER}" & "{F2}")
							EndIf
							$hWnd1 = WinWaitActive('Элемент: "920', "", 5)
							If $hWnd1 Then
								Send("{HOME}" & "{DOWN}" & "{ENTER}")
							EndIf
							$hWnd1 = WinWaitActive($IrbisTit, "", 5)
							If $hWnd1 Then
								ControlSend($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:2]", "{HOME}" & "{DOWN}")
							EndIf
						EndIf
					Case "спек" ; "спек" - переделывает описание в SPEC
						Send("!q" & 900 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "900', "", 5)
						If $hWnd Then
							Send("{DOWN}" & "{F2}")
							$hWnd = WinWaitActive("Подполе", "", 5)
							If $hWnd Then
								Send("{HOME}" & "{DOWN 2}" & "{ENTER}")
							EndIf
							$hWnd = WinWaitActive('Элемент: "900', "", 5)
							If $hWnd Then
								Send("{TAB}" & "{ENTER}")
							EndIf
							$hWnd = WinWaitActive($IrbisTit, "", 5)
							If $hWnd Then
								Send("!q" & 920 & "{ENTER}" & "{F2}")
							EndIf
							$hWnd = WinWaitActive('Элемент: "920', "", 5)
							If $hWnd Then
								Send("{HOME}" & "{DOWN 2}" & "{ENTER}")
							EndIf
							$hWnd = WinWaitActive($IrbisTit, "", 5)
							If $hWnd Then
								ControlSend($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:2]", "{HOME}" & "{DOWN 2}")
							EndIf
						EndIf




					Case "эбс" ; "эбс" - создание 830 поля и внесение сведений о копии в ЭБ
						;**** получение инв. номера и занесение строки в буфер
						WinActivate($IrbisTit)
						$hWnd = WinWaitActive($IrbisTit, "", 5)
						If $hWnd Then
							Send("!q")
							Sleep(100)
							Send(910 & "{ENTER}")
							$text = WinGetText($IrbisTit, "")
							$pos1 = StringInStr($text, "^B", 0, 1) + 2
							$pos2 = StringInStr($text, "^C", 0, 1)
							$invNumLen = $pos2 - $pos1
							$invNum = (StringMid($text, $pos1, $invNumLen))
							ClipPut("^0w=^!ЭБ^a" & $invNum)
						EndIf
						;**** добавление 830 поля
						WinActivate($IrbisTit)
						$hWnd = WinWaitActive($IrbisTit, "", 5)
						If $hWnd Then
							Send("!r")
							$hWnd1 = WinWaitActive("Добавить элемент в РЛ", "", 5)
							If $hWnd1 Then
								ControlFocus("Добавить элемент в РЛ", "", "[CLASS:TListBox; INSTANCE:1]")
								Send("{END}" & "{UP 2}" & "{ENTER}")
								$hWnd1 = WinWaitActive("Внимание", "", 1)
								If $hWnd1 Then
									WinClose("Добавить элемент в РЛ")
									WinClose("Внимание")
									Send("!q" & 830 & "{ENTER}" & "^{ENTER}")
								EndIf
								Send("+{INS}")
							EndIf
						EndIf

					Case "изм" ; "изм" - создание 830 поля и внесение сведений об изменении индекса МДА и авт. зн.
						;**** получение инв. номера и занесение строки в буфер
						WinActivate($IrbisTit)
						$hWnd = WinWaitActive($IrbisTit, "", 5)
						If $hWnd Then
							Send("!q")
							Sleep(100)
							Send(910 & "{ENTER}")
							$text = WinGetText($IrbisTit, "")
							$pos1 = StringInStr($text, "^B", 0, 1) + 2
							$pos2 = StringInStr($text, "^C", 0, 1)
							$invNumLen = $pos2 - $pos1
							$invNum = (StringMid($text, $pos1, $invNumLen))
							Send("!q")
							Sleep(100)
							Send(686 & "{ENTER}" & "+{END}" & "^c")
							$indMDA = ClipGet()
							Send("!q")
							Sleep(100)
							Send(908 & "{ENTER}" & "+{END}" & "^c")

							$avtZn = ClipGet()
							ClipPut("^007^!07^aИнв. " & $invNum & ": до " & _NowDate() & " индекс МДА " & $indMDA & ", авт. знак " & $avtZn)
						EndIf
						;**** добавление 830 поля
						WinActivate($IrbisTit)
						$hWnd = WinWaitActive($IrbisTit, "", 5)
						If $hWnd Then
							Send("!r")
						EndIf
						$hWnd1 = WinWaitActive("Добавить элемент в РЛ", "", 5)
						If $hWnd1 Then
							ControlFocus("Добавить элемент в РЛ", "", "[CLASS:TListBox; INSTANCE:1]")
							Send("{END}" & "{UP 2}" & "{ENTER}")
							$hWnd1 = WinWaitActive("Внимание", "", 1)
							If $hWnd1 Then
								WinClose("Добавить элемент в РЛ")
								WinClose("Внимание")
								WinActivate($IrbisTit)
								$hWnd = WinWaitActive($IrbisTit, "", 5)
								If $hWnd Then
									Send("!q" & 830 & "{ENTER}")
									Sleep(100)
									Send("^{ENTER}")
								EndIf
							EndIf
							$hWnd = WinWaitActive($IrbisTit, "", 5)
							If $hWnd1 Then
								Send("+{INS}")
							EndIf
						EndIf
					Case "конв" ; "конв" - создание 830 поля и внесение сведений о прежнем вхождении в конволют
						ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:4]", "left", 1, 230, 10)
						Send("{TAB 2}" & "{END}" & "{UP 2}" & "{ENTER}")
						$wTit = WinGetTitle("[ACTIVE]")
						$aTT = StringInStr($wTit, "Внимание")
						If $aTT = 1 Then
							WinClose("Внимание")
							WinClose("Добавить")
						Else
							$aTT = 0
						EndIf
						Sleep(300)
						Send("!q" & 910 & "{ENTER}" & "{F2}")
						$hWnd = WinWaitActive('Элемент: "910', "", 5)
						If $hWnd Then
							Send("{DOWN}")
							Sleep(500)
							$invNum = ControlGetText('Элемент: "910', "", "[CLASS:TTntRichEdit.UnicodeClass; INSTANCE:1]")
							WinClose('Элемент: "910')
							Send("!q" & 830 & "{ENTER}")
						EndIf
						If $aTT = 1 Then
							Sleep(300)
							Send("^{ENTER}")
						EndIf
						ClipPut("^007^!07^aИнв. " & $invNum & ": до " & _NowDate() & " входил в конволют (инв. )")
						Send("+{INS}")


						; 			3) Переход по базам
					Case "журн" ; "журн" - Периодические издания с 2014 г.
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
						Send("{ALT}" & "{ENTER 2}" & "PR" & "{ENTER}")
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
					Case "мда" ; "мда" - база МПДА
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
						Send("{ALT}" & "{ENTER 2}" & "MPDA" & "{ENTER}")
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
					Case "дис" ; "дис" - база диссертаций
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
						Send("{ALT}" & "{ENTER 2}" & "DST" & "{ENTER}")
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
					Case "инос" ; "инос" или "ифн" - фонд на иностранных языках
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
						Send("{ALT}" & "{ENTER 2}" & "IFN" & "{ENTER}")
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
					Case "ифн"
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
						Send("{ALT}" & "{ENTER 2}" & "IFN" & "{ENTER}")
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)

						; 			4) Справочники
					Case "рпк" ; "рпк" - Рос. правила кат-ции, файл должен быть по пути d:\РПК.pdf
						Run(@ProgramFilesDir & "\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe d:\РПК.pdf")
					Case "сокр"
						Run("c:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe d:\dESCTOP\ГОСТ_7.0.12-2011_Сокращ_слов.pdf")
					Case "инс"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, "d:\dESCTOP\Инструкции (запись диак. Павла).doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "форм"

						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, "d:\dESCTOP\2.Формуляры\Формуляр лежачий (история выдач).doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "скан"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, "d:\dESCTOP\СканКоп.rtf")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)


						; 			5) Рубрики. Открытие файлов с таблицами рубрик - индекс МДА без тире ("а0"). История России - "б8р", Рус. лит-ра - "г3р".
						; Путь до файлов - d:\Рубрики по уровням\. Название файлов - только индекс МДА (А0.doc)
						$sPath_ini = @ScriptDir & "\IrbisHotkeys.ini"
						$rubDir = IniRead($sPath_ini, "Sec1", "RubDir", "d:\Рубрики по уровням\")
					Case "а0"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "А0.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "а1"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "А1.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "а2"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "А2.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "а3"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "А3.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "а4"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "А4.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "а5"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "А5.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "а6"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "А6.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "а7"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "А7.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "а8"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "А8.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "а9"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "А9.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "а10"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "А9.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "б0"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Б0.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "б1"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Б1.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "б2"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Б2.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "б3"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Б3.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "б4"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Б4.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "б5"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Б5.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "б6"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Б6.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "б8"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Б8.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "б8р"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Б8р.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "б9"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Б9.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "б10"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Б10.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "в0"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "В0.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "в1"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "В1.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "в2"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "В2.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "в3"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "В3.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "в4"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "В4.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "в5"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "В5.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "г0"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Г0.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "г1"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Г1.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "г2"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Г1.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "г3"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Г3.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "г3р"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Г3р.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "г4"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Г4.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "д3"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Д3.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "д4"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Д4.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "д5"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Д5.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "е"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Е.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "ж"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Ж.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "з0"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "З0.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "з1"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "З1.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "з2"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "З2.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "з3"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "З3.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "з4"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "З4.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "з5"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "З5.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "к0"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "К0.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "к1"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "К1.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "к3"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "К3.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "п1"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "П1.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "п2"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "П2.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)
					Case "ц"
						Local $oWord = _Word_Create()
						_Word_DocOpen($oWord, $rubDir & "Ц.doc")
						Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
						WinActivate($hWnd)

					Case Else
						$exit = 1
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
						$hWnd1 = WinWaitActive($IrbisTit, "", 5)
						If $hWnd1 Then
							$input = InputBox("Внимание", "Неправильное значение. Повторите", "", "", 190, 140)
							_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
						EndIf
				EndSwitch
			Until $exit = 0
		EndIf
	EndIf
EndFunc   ;==>Field
;

;						CTRL+F Поиск по виду основного словаря. Читает введенную строку и выполняет поиск
Func Search()


	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^f")
		Send("^f")
		HotKeySet("^f", "Search")
	Else
		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
		$input = InputBox("Выполнить", "Поиск по:", "", "", 190, 130)
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)

		ControlFocus($IrbisTit, "", "[CLASS:TTntStringGrid.UnicodeClass; INSTANCE:3]")
		; если введенная строка - число, открытие записи в Ирбисе по инв. номеру

		If StringIsInt($input) Then
			Send("!f" & "инв" & "{ENTER}")
			Sleep(100)
			Send("!d" & $input & "{ENTER}")
		Else


			Do
				$exit = 0
				Switch $input
					Case "" ; пустое поле - закрытие окна
						ExitLoop
					Case "тех" ; "тех" - по технологии
						Send("!f" & "тех" & "{ENTER}")
						Sleep(100)
						Send("!d")
					Case "изд" ; "изд" - по изд-ву
						Send("!f" & "изд" & "{DOWN}" & "{ENTER}")
						Sleep(100)
						Send("!d")
					Case "инв" ; "инв" - по инв. номеру
						Send("!f" & "инв" & "{ENTER}")
						Sleep(100)
						Send("!d")
					Case "мн" ; "мн" - все многотомники
						Send("!f" & "вид" & "{ENTER}")
						Sleep(100)
						Send("!d" & "03" & "{ENTER}")
					Case "авт" ; "авт" - по автору
						Send("!f" & "авт" & "{DOWN 2}" & "{ENTER}")
						Sleep(100)
						Send("!d")
					Case "под" ; "под" - по предметному подзаголовку
						Send("!f" & "пред" & "{ENTER}")
						Sleep(100)
						Send("!d")
					Case "руб" ; "руб" - по предметной рубрике
						Send("!f" & "пред" & "{DOWN}" & "{ENTER}")
						Sleep(100)
						Send("!d")
					Case "заг" ; "заг" - по заглавию
						Send("!f" & "заг" & "{ENTER}")
						Sleep(100)
						Send("!d")
					Case "кл" ; "кл" - по ключевому слову
						Send("!f" & "кл" & "{ENTER}")
						Sleep(100)
						Send("!d")
					Case "перс" ; "перс" - по персоналии
						Send("!f" & "перс" & "{ENTER}")
						Sleep(100)
						Send("!d")
					Case Else
						$exit = 1
						_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
						$hWnd1 = WinWaitActive($IrbisTit, "", 5)
						If $hWnd1 Then
							$input = InputBox("Внимание", "Неправильное значение." & @CRLF & "Повторите.", "", "", 190, 140)
							_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
						EndIf




				EndSwitch

			Until $exit = 0
		EndIf
	EndIf

EndFunc   ;==>Search

;						ALT+H Последовательный поиск
; Ввести значение. Нажать TAB и ввести номер поля. Нажать еще раз TAB и выбрать уточняемый запрос. Нажать ENTER
Func CohSearch()
	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("!h")
		Send("!h")
		HotKeySet("!h", "CohSearch")
	Else
		ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:3]", "left", 1, 80, 11)
		Sleep(100)
		ControlClick("Последовательный", "", "[CLASS:TTabbedNotebook; INSTANCE:1]", "left", 2, 155, 12)
		Sleep(100)
		ControlFocus("Последовательный", "", "[CLASS:TTntEdit.UnicodeClass; INSTANCE:1]")

		Local $hDLL = DllOpen("user32.dll")
		$FirstTab = 0
		While 1
			If _IsPressed("0D", $hDLL) Then
				ControlClick("Последовательный", "", "[CLASS:TBitBtn; INSTANCE:3]", "left", 1)
				ExitLoop
			ElseIf _IsPressed("09", $hDLL) Then
				If $FirstTab = 1 Then
					While _IsPressed("09", $hDLL)
						Sleep(250)
					WEnd
					ControlClick("Последовательный", "", "[CLASS:THSHintTntComboBox.UnicodeClass; INSTANCE:1]", "left", 1)
				Else
					While _IsPressed("09", $hDLL)
						Sleep(250)
					WEnd
					$FirstTab = 1
				EndIf
			ElseIf _IsPressed("1B", $hDLL) Then
				ExitLoop
			EndIf

		WEnd

		DllClose($hDLL)
	EndIf
EndFunc   ;==>CohSearch

;						CTRL+W Печать основной карточки
Func Osn()


	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^w")
		Send("^w")
		HotKeySet("^w", "Osn")
	Else
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)

		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
		;**** Проверка на наличие автора
		$autorExM = 0
		$autorEx = 0
		;**** в многотомнике
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
		Send("!q" & 961 & "{ENTER}")
		$hWnd = WinWaitActive("ОШИБКА", "", 1)
		If $hWnd Then
			WinClose("ОШИБКА")
		Else
			$wText = WinGetText($IrbisTit)
			$autorExM = StringInStr($wText, "ДА")
			If $autorExM > 0 Then
				$autorExM = 1
			Else
				$autorExM = 2
			EndIf
		EndIf

		;**** в однотомнике

		Send("!q" & 700 & "{ENTER}")
		$wText = WinGetText($IrbisTit)
		$autorEx = StringInStr($wText, "A")

		If $autorEx > 0 Then
			$autorEx = 1
		Else
			$autorEx = 0
		EndIf


		Send("!q" & 910 & "{ENTER}")
		$text = WinGetText($IrbisTit, "")
		$pos1 = StringInStr($text, "^B", 0, 1) + 2
		$pos2 = StringInStr($text, "^C", 0, 1)
		$invNumLen = $pos2 - $pos1
		$invNum = (StringMid($text, $pos1, $invNumLen))

		ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 34, 12)
		$hWnd = WinWaitActive("Печать", "", 5)
		If $hWnd Then
			ControlSend("Печать", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{HOME}" & "{DOWN 7}")
			Sleep(200)
			ControlClick("Печать", "", "[CLASS:TBitBtn; INSTANCE:2]", "left", 2)
		EndIf
		$hWnd = WinWaitActive("Файл", "", 5)
		If $hWnd Then
			ControlSend("Файл", "", "[CLASS:Edit; INSTANCE:1]", $invNum)
			Send("{TAB 2}" & "{ENTER}")
		EndIf
		$hWnd = WinWaitActive("Внимание", "", 5)
		If $hWnd Then
			ControlClick("Внимание", "", "[CLASS:Button; INSTANCE:1]")
		EndIf
		$hWnd = WinWaitActive("[CLASS:OpusApp]", "", 5)
		If $hWnd Then
			WinClose("Печать текущего")
			$Object = ObjGet("", "Word.Application")

			If $autorExM = 2 And $autorEx = 1 Then
				$Object.Run("Макрос2")
			ElseIf $autorExM = 1 Or $autorEx = 1 Then
				$Object.Run("Макрос1")
			Else
				$Object.Run("Макрос2")
			EndIf

		Else
		EndIf

	EndIf
EndFunc   ;==>Osn

;						CTRL+K Печать контрольной карточки
Func KK()


	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^k")
		Send("^k")
		HotKeySet("^k", "KK")
	Else
		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")

		WinActivate($IrbisTit)
		ControlFocus($IrbisTit, "", "[CLASS:TTntStringGrid.UnicodeClass; INSTANCE:3]")
		ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 11, 12)
		$hWnd1 = WinWaitActive("Печать", "", 5)
		If $hWnd1 Then
			ControlSend("Печать", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{HOME}" & "{DOWN 10}")
			ControlClick("Печать", "", "[CLASS:TBitBtn; INSTANCE:2]")
		EndIf
		$hWnd1 = WinWaitActive("Файл", "", 5)
		If $hWnd1 Then
			$numDoc = 1
			If WinExists("1 [Режим") Then
				If WinExists("2 [Режим") Then
					$numDoc = 3
				Else
					$numDoc = 2
				EndIf
			EndIf
			ControlSend("Файл", "", "[CLASS:Edit; INSTANCE:1]", $numDoc)
			Send("{TAB 2}" & "{ENTER}")
		EndIf
		$hWnd1 = WinWaitActive("Подтвердить", "", 5)
		If $hWnd1 Then
			ControlClick("Подтвердить", "", "[CLASS:Button; INSTANCE:1]")
		EndIf

		$hWnd1 = WinWaitActive("Внимание", "", 5)
		If $hWnd1 Then
			ControlClick("Внимание", "", "[CLASS:Button; INSTANCE:1]")
		EndIf


		$hWnd1 = WinWaitActive("[CLASS:OpusApp]", "", 10)
		If $hWnd1 Then
			WinClose("Печать выходных форм - Результат поиска")
			$Object = ObjGet("", "Word.Application")
			$Object.Run("KK")

		EndIf
	EndIf
EndFunc   ;==>KK

;						CTRL+D Печать формуляра
Func Formular()


	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^d")
		Send("^d")
		HotKeySet("^d", "Formular")
	Else
		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
		WinActivate($IrbisTit)
		ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 11, 12)
		$hWnd1 = WinWaitActive("Печать", "", 5)
		If $hWnd1 Then
			ControlSend("Печать", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{HOME}" & "{DOWN 4}")
			Sleep(200)
			ControlClick("Печать", "", "[CLASS:TBitBtn; INSTANCE:2]")
		EndIf
		$hWnd1 = WinWaitActive("Файл", "", 5)
		If $hWnd1 Then
			$numDoc = 1
			If WinExists("1 [Режим") Then
				If WinExists("2 [Режим") Then
					$numDoc = 3
				Else
					$numDoc = 2
				EndIf
			EndIf
			ControlSend("Файл", "", "[CLASS:Edit; INSTANCE:1]", $numDoc)
			Send("{TAB 2}" & "{ENTER}")
		EndIf
		$hWnd1 = WinWaitActive("Подтвердить", "", 5)
		If $hWnd1 Then
			ControlClick("Подтвердить", "", "[CLASS:Button; INSTANCE:1]")
		EndIf
		$hWnd1 = WinWaitActive("Внимание", "", 5)
		If $hWnd1 Then
			ControlClick("Внимание", "", "[CLASS:Button; INSTANCE:1]")
		EndIf

		$hWnd1 = WinWaitActive("[CLASS:OpusApp]", "", 5)
		If $hWnd1 Then
			$Object = ObjGet("", "Word.Application")
			WinClose("Печать выходных форм - Результат поиска")
			$Object.Run("Formular")
		EndIf
	EndIf
EndFunc   ;==>Formular

;						CTRL+Y Печать ярлычка
Func Label()


	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^y")
		Send("^y")
		HotKeySet("^y", "Label")
	Else
		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
		WinActivate($IrbisTit)
		ControlFocus($IrbisTit, "", "[CLASS:TTntStringGrid.UnicodeClass; INSTANCE:3]")
		ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 11, 12)
		$hWnd1 = WinWaitActive("Печать", "", 5)
		If $hWnd1 Then
			ControlSend("Печать", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{HOME}" & "{DOWN 3}")
			Sleep(200)
			ControlClick("Печать", "", "[CLASS:TBitBtn; INSTANCE:2]")
		EndIf
		$hWnd1 = WinWaitActive("Файл", "", 5)
		If $hWnd1 Then
			$numDoc = 1
			If WinExists("1 [Режим") Then
				If WinExists("2 [Режим") Then
					$numDoc = 3
				Else
					$numDoc = 2
				EndIf
			EndIf
			ControlSend("Файл", "", "[CLASS:Edit; INSTANCE:1]", $numDoc)
			Send("{TAB 2}" & "{ENTER}")
		EndIf
		$hWnd1 = WinWaitActive("Подтвердить", "", 5)
		If $hWnd1 Then
			ControlClick("Подтвердить", "", "[CLASS:Button; INSTANCE:1]")
		EndIf
		$hWnd1 = WinWaitActive("Внимание", "", 5)
		If $hWnd1 Then
			ControlClick("Внимание", "", "[CLASS:Button; INSTANCE:1]")
		EndIf

		$hWnd1 = WinWaitActive("[CLASS:OpusApp]", "", 5)
		If $hWnd1 Then
			$Object = ObjGet("", "Word.Application")
			WinClose("Печать выходных форм - Результат поиска")
			$Object.Run("Label_2")
		EndIf
	EndIf
EndFunc   ;==>Label

;						CTRL+F8 Закрепить окно поверх всех
Func OnTop()
	Sleep(10)
	Send("{CTRLDOWN}")
	Sleep(10)
	Send("{CTRLUP}")


	Local $hWnd = WinGetHandle("[ACTIVE]")
	WinSetOnTop($hWnd, "", 1)
EndFunc   ;==>OnTop

;						CTRL+F9 Отменить закрепление окна поверх всех
Func OnTopOff()
	Sleep(10)
	Send("{CTRLDOWN}")
	Sleep(10)
	Send("{CTRLUP}")

	Local $hWnd = WinGetHandle("[ACTIVE]")
	WinSetOnTop($hWnd, "", 0)
EndFunc   ;==>OnTopOff

;						CTRL+M Печать ярлычка для мягких формуляров
Func FormularLabel()


	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^+m")
		Send("^+m")
		HotKeySet("^+m", "FormularLabel")
	Else
		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
		WinActivate($IrbisTit)
		ControlFocus($IrbisTit, "", "[CLASS:TTntStringGrid.UnicodeClass; INSTANCE:3]")
		ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 11, 12)
		$hWnd1 = WinWaitActive("Печать", "", 5)
		If $hWnd1 Then
			ControlSend("Печать", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{HOME}" & "{DOWN 5}")
			Sleep(200)
			ControlClick("Печать", "", "[CLASS:TBitBtn; INSTANCE:2]")
		EndIf
		$hWnd1 = WinWaitActive("Файл", "", 5)
		If $hWnd1 Then
			$numDoc = 1
			If WinExists("1 [Режим") Then
				If WinExists("2 [Режим") Then
					$numDoc = 3
				Else
					$numDoc = 2
				EndIf
			EndIf
			ControlSend("Файл", "", "[CLASS:Edit; INSTANCE:1]", $numDoc)
			Send("{TAB 2}" & "{ENTER}")
		EndIf
		$hWnd1 = WinWaitActive("Подтвердить", "", 5)
		If $hWnd1 Then
			ControlClick("Подтвердить", "", "[CLASS:Button; INSTANCE:1]")
		EndIf
		$hWnd1 = WinWaitActive("Внимание", "", 5)
		If $hWnd1 Then
			ControlClick("Внимание", "", "[CLASS:Button; INSTANCE:1]")
		EndIf
		$hWnd1 = WinWaitActive("[CLASS:OpusApp]", "", 5)
		If $hWnd1 Then
			WinClose("Печать выходных форм - Результат поиска")
		EndIf
	EndIf
EndFunc   ;==>FormularLabel

;						CTRL+SPACE Фокус на окне полного описания
Func ViewFocus()
	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^{SPACE}")
		Send("^{SPACE}")
		HotKeySet("^{SPACE}", "ViewFocus")
	Else
		ControlClick($IrbisTit, "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "left", 1, 1034, 25)
	EndIf
EndFunc   ;==>ViewFocus
