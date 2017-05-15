#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=IrbisHotkeys.ico
#AutoIt3Wrapper_Outfile=IrbisHotkeys.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <WinAPISys.au3>
#include <Word.au3>
#include <Misc.au3>
#include <Date.au3>
#include <MsgBoxConstants.au3>
#include <Excel.au3>

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

HotKeySet("^w", "Print")
HotKeySet("^k", "Print")
HotKeySet("^d", "Print")
HotKeySet("^y", "Print")
HotKeySet("^m", "Print")

HotKeySet("^{F8}", "OnTop")
HotKeySet("^{F9}", "OnTopOff")
HotKeySet("^+g", "Obrzv")
HotKeySet("^{F12}", "ScrExit")
HotKeySet("^{SPACE}", "ViewFocus")
HotKeySet("^+k", "CopySelected")

$clip = ''
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
	If HotKeyOn("^+g", "Obrzv") Then
		;~ 		Предотвращение залипания SHIFT'a
		Sleep(10)
		Send("{SHIFTDOWN}")
		Sleep(10)
		Send("{SHIFTUP}")
		$wTit = WinGetTitle("[ACTIVE]")
		$SPA = StringInStr($wTit, "СПА")

		If $SPA <> 0 Then
;~ 							Установка этапа работы
			ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 176, 13)
			$hWnd = WinWaitActive("Установка личных параметров", "", 5)
			If $hWnd Then
				WinMove("Установка личных параметров", "", 594, 295, 472, 310)
				ControlClick("Установка личных параметров", "", "[CLASS:TStringGrid; INSTANCE:1]", "left", 1, 347, 30)
				_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
				Send("{HOME}" & "+{END}" & 'obrzv' & "{TAB}{ENTER}")
			EndIf
			;**** корректировка
			$glW = "Глобальная корректировка БД"
			$hWnd = WinWaitActive($IrbisTit, "", 3)
			If $hWnd Then
				Sleep(100)
				ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:4]", "left", 1, 323, 11)
				$gWnd = WinWaitActive($glW, "", 3)
				If $gWnd Then
					Sleep(300)
					ControlClick($glW, "", "[CLASS:TCheckBox; INSTANCE:5]")
					ControlClick($glW, "", "[CLASS:TToolBar; INSTANCE:3]", "left", 1, 101, 11)
				EndIf

				$gWnd = WinWaitActive("Выбор", "", 3)
				If $gWnd Then
					ControlClick("Выбор", "", "[CLASS:TBitBtn; INSTANCE:2]")
				EndIf

				$gWnd = WinWaitActive("Открытие", "", 3)
				If $gWnd Then
					ClipMan("!!КР-ФЛК")
					Send("{DOWN}" & "{ENTER}")
					ControlClick($glW, "", "[CLASS:TButton; INSTANCE:5]")
				EndIf
;~ 				Перезапуск скрипта - временное решение бага с отключением CTRL
				Run(FileGetShortName(@ScriptFullPath))
			EndIf
		EndIf

	EndIf
EndFunc   ;==>Obrzv

;						CTRL+V Вставка. Вставляет в поле данные без раскрытия окна "Элемент"
Func Vstavka()
	If HotKeyOn("^v", "Vstavka") Then
		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
		Send("+{INS}")
	EndIf
EndFunc   ;==>Vstavka

;						CTRL+S Сохранение записи в Ирбисе
Func IrbSave()
	If HotKeyOn("^s", "IrbSave") Then
		Send("+{ENTER}")
	EndIf
EndFunc   ;==>IrbSave

;						CTRL+Z Вывод нескольких инв. номеров. Ввести инв. номера и нажать TAB
Func SearchNumbs()
	If HotKeyOn("^z", "SearchNumbs") Then

;~ 		Фокус на рабочем листе
		ControlFocus($IrbisTit, "", "[CLASS:TTntStringGrid.UnicodeClass; INSTANCE:3]")
;~ 		Получение названия базы и сравнение
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
;~ 			Если запрос по базе период. изданий, то отбор по заглавию
			If $isPRBD = 0 Then
				ControlSend("Поиск по словарю/Рубрикатору", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{HOME}")
				Sleep(100)
				ControlSend("Поиск по словарю/Рубрикатору", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{DOWN}")
			Else
;~ 			Иначе - отбор по инв. номеру
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

;						CTRL+Q Разные команды. Читает введенную строку и выполняет команды
Func Field()

	If HotKeyOn("^q", "Field") Then
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
		$input = InputBox("Выполнить", "Название поля:", "", "", 190, 130)
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
		If StringLen($input) > 100 Then
			$input = "error"
		EndIf
		If WinWaitActive($IrbisTit, "", 5) Then
			ControlFocus($IrbisTit, "", "[CLASS:TTntStringGrid.UnicodeClass; INSTANCE:3]")

			;**** Пути до рубрик
			$sPath_ini = @ScriptDir & "\IrbisHotkeys.ini"
			$rubDir = IniRead($sPath_ini, "Sec1", "RubDir", "d:\Рубрики по уровням\")

			; 			1. Если введенная строка - число
			If StringIsInt($input) Then
				; Переход по номеру поля, если цифр в числе меньше 4
				If StringLen($input) < 4 Then
					GoToField($input)
				Else
					; Открытие записи в Ирбисе по инв. номеру, если цифр в числе больше 3
					Srchfor("инв")
					ClipMan($input)
					Send("{ENTER}")
				EndIf
			Else
				; 			2. Переход по полям. Вид: введенная строка - выполняемая команда (открытие полей, файлов и проч.)

				Do
					$exit = 0
					WinActivate("Внимание")
					ControlFocus("Внимание", "", "[CLASS:Edit; INSTANCE:1]")
					Switch $input

						Case "" ; Пустая строка - закрытие окна
							ExitLoop
						Case "уо"
							$wTit = WinGetTitle("[ACTIVE]")
							$SPA = StringInStr($wTit, "СПА")
							If $SPA <> 0 Then
								If OpenElement(910) Then
									Send("{DOWN 3}")
									Sleep(100)
									Send("+{END}" & "У0-к" & "{TAB}" & "{ENTER}")
									Sleep(300)
;~ 								Send("+{ENTER}")
								EndIf
							EndIf
						Case "прк" ; "прк" - этап работы ПРК
;~ 							Получение названия текущей базы и сравнение
							$bdName = ControlGetText($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:4]")
							$isDSTBD = StringInStr($bdName, "DST - Диссертационная база МДА")

;~ 							Вызов окна личных параметров
							ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 176, 13)
							$hWnd = WinWaitActive("Установка личных параметров", "", 5)
							If $hWnd Then
								WinMove("Установка личных параметров", "", 594, 295, 472, 310)
								ControlClick("Установка личных параметров", "", "[CLASS:TStringGrid; INSTANCE:1]", "left", 1, 347, 30)
								Send("{F2}")
								$hWnd1 = WinWaitActive('"Этап работы"', "", 5)
								If $hWnd1 Then
;~ 									Если это база диссертаций, установка этапа работы КР
									If $isDSTBD Then
										Send("{HOME}{PGDN}{DOWN 3}{ENTER}")
									Else
										Send("{HOME}{PGDN}{DOWN 2}{ENTER}")
									EndIf
								EndIf
								If $hWnd Then
									Send("{TAB}{ENTER}")
								EndIf
							EndIf

						Case "кр" ; "кр" - этап работы КР
;~ 							Получение названия текущей базы и сравнение
							$bdName = ControlGetText($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:4]")
							$isDSTBD = StringInStr($bdName, "DST - Диссертационная база МДА")

;~ 							Вызов окна личных параметров
							ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 176, 13)
							$hWnd = WinWaitActive("Установка личных параметров", "", 5)
							If $hWnd Then
								WinMove("Установка личных параметров", "", 594, 295, 472, 310)
								ControlClick("Установка личных параметров", "", "[CLASS:TStringGrid; INSTANCE:1]", "left", 1, 347, 30)
								Send("{F2}")
								$hWnd1 = WinWaitActive('"Этап работы"', "", 5)
								If $hWnd1 Then
;~ 									Если это база диссертаций, установка этапа работы КР
									If $isDSTBD Then
										Send("{HOME}{PGDN}{DOWN 3}{ENTER}")
									Else
										Send("{HOME}{PGDN}{DOWN}{ENTER}")
									EndIf
								EndIf
								If $hWnd Then
									Send("{TAB}{ENTER}")
								EndIf
							EndIf

						Case "прф" ; "прф" - этап работы ПРФ
							ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 176, 13)
							$hWnd = WinWaitActive("Установка личных параметров", "", 5)
							If $hWnd Then
								WinMove("Установка личных параметров", "", 594, 295, 472, 310)
								ControlClick("Установка личных параметров", "", "[CLASS:TStringGrid; INSTANCE:1]", "left", 1, 347, 30)
								Send("{F2}")
								$hWnd1 = WinWaitActive('"Этап работы"', "", 5)
								If $hWnd1 Then
									Send("{END}{PGUP}{UP 2}{ENTER}")
								EndIf
								If $hWnd Then
									Send("{TAB}{ENTER}")
								EndIf
							EndIf

						Case "сис" ; "сис" - этап работы С
							ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 176, 13)
							$hWnd = WinWaitActive("Установка личных параметров", "", 5)
							If $hWnd Then
								WinMove("Установка личных параметров", "", 594, 295, 472, 310)
								ControlClick("Установка личных параметров", "", "[CLASS:TStringGrid; INSTANCE:1]", "left", 1, 347, 30)
								Send("{F2}")
								$hWnd1 = WinWaitActive('"Этап работы"', "", 5)
								If $hWnd1 Then
									Send("{HOME}{DOWN 2}{ENTER}")
								EndIf
								If $hWnd Then
									Send("{TAB}{ENTER}")
								EndIf
							EndIf

						Case "пкт" ; "пкт" - этап работы ОБРНЗ, ПКТ
							ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 176, 13)
							$hWnd = WinWaitActive("Установка личных параметров", "", 5)
							If $hWnd Then
								WinMove("Установка личных параметров", "", 594, 295, 472, 310)
								ControlClick("Установка личных параметров", "", "[CLASS:TStringGrid; INSTANCE:1]", "left", 1, 347, 30)
								Send("{F2}")
								$hWnd1 = WinWaitActive('"Этап работы"', "", 5)
								If $hWnd1 Then
									Send("{HOME}{ENTER}")
								EndIf
								If $hWnd Then
									Send("{TAB}{ENTER}")
								EndIf
							EndIf

						Case "хар" ; "хар" - характер документа
							If OpenElement(900) Then
								Send("{DOWN 3}")
							EndIf
						Case "сер" ; "сер" - заглавие серии
							If OpenElement(225) Then
								Send("{DOWN 2}")
							EndIf
						Case "кат" ; "кат" - поле "Каталогизатор"
							GoToField(907)
						Case "анн" ; "анн" - аннотация
							OpenElement(331)
						Case "заг" ; "заг" - заглавие однотомника
							OpenElement(200)
						Case "авт" ; "авт" - автор однотомника
							OpenElement(700)
						Case "давт" ; "давт" - другие авторы однотомника
							OpenElement(701)
						Case "ред" ; "ред" - редакторы однотомника
							If OpenElement(702) Then
								Send("{ENTER}")
							EndIf
						Case "кол" ; "кол" - другие коллективы, не входящие в заголовок описания
							If OpenElement(711) Then
								Send("{ENTER}")
							EndIf
						Case "мес" ; "мес" - место издание однотомника
							If OpenElement(210) Then
								Send("{PGDN}" & "{UP 9}")
							EndIf
						Case "изде" ; "изде" - сведения об издании
							OpenElement(205)
						Case "изд" ; "изд" - издательство однотомника
							If OpenElement(210) Then
								Send("{DOWN 3}")
							EndIf
						Case "тип" ; "тип" - типография однотомника
							If OpenElement(210) Then
								Send("{PGDN}" & "{UP}")
							EndIf
						Case "год" ; "год" - год издания однотомника
							OpenElement(210)
						Case "отв" ; "отв" - сведения об ответственности однотомника
							If OpenElement(200) Then
								Send("{PGDN}")
								Sleep(100)
								Send("{UP}")
							EndIf
						Case "свед" ; "свед" - сведения к заглавию однотомника
							If OpenElement(200) Then
								Send("{PGDN}")
								Sleep(200)
								Send("{UP 2}")
							EndIf
						Case "прим" ; "прим" - общие примечания однотомника
							GoToField(300)
						Case "супо" ; "супо" - вставка в общ. примеч. "Изд. в суперобл."
							GoToField(300)
							Send("Изд. в суперобл.")
						Case "разн" ; "разн" - разночтение заглавий
							If OpenElement(517) Then
								Send("{DOWN 2}")
							EndIf
						Case "инд" ; "инд" - индекс МДА
							OpenElement(686)
						Case "аз" ; "аз" - авторский знак
							OpenElement(908)
						Case "руб" ; "руб" - предметная рубрика
							If OpenElement(606) Then
								Send("{DOWN}")
							EndIf
						Case "рубр" ; "рубр" - предметная рубрика, с раскытием первого подзаголовка на весь экран
							If OpenElement(606) Then
								Send("{DOWN}" & "{F2}")
								$hWnd = WinWaitActive("Подполе")
								If $hWnd Then
									Send("#{UP}")
								EndIf
							EndIf
						Case "перс" ; "перс" - персоналия
							OpenElement(600)
						Case "стр" ; "стр" - количеств. хар-ки
							OpenElement(215)
						Case "мзаг" ; "мзаг" - заглавие многотомника
							OpenElement(461)
						Case "мавт" ; "мавт" - автор многотомника или первый редактор
							If OpenElement(961) Then
								Send("{DOWN 7}")
							EndIf
						Case "мизд" ; "мизд" - изд-во многотомника
							If OpenElement(461) Then
								Send("{ENTER 6}")
							EndIf
						Case "мгод" ; "мгод" - год начала издания многотомника
							If OpenElement(461) Then
								Send("{ENTER 12}")
							EndIf
						Case "мотв" ; "мотв" - сведения об ответственности многотомника
							If OpenElement(461) Then
								Sleep(100)
								ControlFocus('Элемент: "461', "", "[CLASS:TTntRichEdit.UnicodeClass; INSTANCE:1")
								Send("{ENTER 4}")
							EndIf
						Case "мсвед" ; "мсвед" - сведения к заглавию многотомника
							If OpenElement(461) Then
								Send("{ENTER 3}")
							EndIf
						Case "мприм" ; "мприм" - общие примечания многотомника
							If OpenElement(46) Then
								Send("{ENTER 8}")
							EndIf
						Case "эб" ; "эб" - 830 поле
							GoToField(830)
						Case "инв" ; "инв" - первый инв. номер
							If OpenElement(910) Then
								Send("{DOWN}")
							EndIf
						Case "осо" ; "осо" - отключение сведений об ответственности
							GoToField(905)
							ClipMan("^11")
						Case "бсз" ; "бсз" - сведения к заглавию с прописной
							GoToField(905)
							ClipMan("^23")
						Case "уни" ; "уни" - описание переделывается в "ЗАПОЛНИТЬ НОВОЙ ЗАПИСЬЮ"
							If OpenElement(503) Then
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
									GoToField(910)
									Send("^{ENTER}")
									ClipMan("^A0^B000-6")
								EndIf
							EndIf
						Case "пп" ; "пп" - отметка о плохом переплете
							GoToField(910)
							$text = WinGetText($IrbisTit, "")
							$pos1 = StringInStr($text, "^B", 0, 1) + 2
							$pos2 = StringInStr($text, "^C", 0, 1)
							$invNumLen = $pos2 - $pos1
							$invNum = (StringMid($text, $pos1, $invNumLen))
							$hWnd1 = WinWaitActive($IrbisTit, "", 5)
							If $hWnd1 Then
								If OpenElement(141) Then
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
							EndIf


						Case "библ" ; "библ" - примеч. о библиографии
							GoToField(320)
						Case "шифр" ; "шифр" - шифр док-та в БД
							GoToField(903)
						Case "сод" ; "сод" - открытие поля "Содержание" в виде таблицы и переход на заглавие
							If OpenElementF3(330) Then
								Send("{ENTER 33} {LEFT 4}")
							EndIf
						Case "пер" ; "пер" - заглавие оригинала переводного издания
							OpenElement(454)
						Case "пар" ; "пар" - параллельное заглавие
							OpenElement(510)
						Case "кл" ; "кл" - ключевые слова
							GoToField(610)
						Case "ис" ; "ис" - ISBN однотомника
							OpenElement(10)
						Case "мис" ; "мис" - ISBN многотомника
							If OpenElement(461) Then
								Send("{PGDN}" & "{UP 10}")
							EndIf
						Case "пазк" ; "пазк" - переделывает описание в PAZK
							If OpenElement(900) Then
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
									If OpenElement(920) Then
										Send("{HOME}" & "{ENTER}")
									EndIf
								EndIf
								$hWnd1 = WinWaitActive($IrbisTit, "", 5)
								If $hWnd1 Then
									ControlSend($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:2]", "{HOME}")
								EndIf
							EndIf
						Case "пвк" ; "пвк" - переделывает описание в PVK
							If OpenElement(900) Then
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
									If OpenElement(920) Then
										Send("{HOME}" & "{DOWN}" & "{ENTER}")
									EndIf
								EndIf
								$hWnd1 = WinWaitActive($IrbisTit, "", 5)
								If $hWnd1 Then
									ControlSend($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:2]", "{HOME}" & "{DOWN}")
								EndIf
							EndIf

						Case "пвкср"
							;**** открытие текущего номера и записи для сравнения
							$invNum = GetInvNum()
							Send("!z")
							$hWnd = WinWaitActive("Поиск по словарю/Рубрикатору", "", 5)
							If $hWnd Then
								ControlSend("Поиск по словарю/Рубрикатору", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{HOME}")
								Sleep(100)
								ControlSend("Поиск по словарю/Рубрикатору", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{DOWN}")
								Sleep(100)
								ControlClick("Поиск по словарю/Рубрикатору", "", "[CLASS:TGroupButton; INSTANCE:6]")
								ControlFocus("Поиск по словарю/Рубрикатору", "", "[CLASS:TTntEdit.UnicodeClass; INSTANCE:1]")
								ClipMan($invNum)
								Send("{ENTER}")
								Sleep(100)
								Send("284411" & "{ENTER}")
								Sleep(200)
								ControlClick("Поиск по словарю/Рубрикатору", "", "[CLASS:TBitBtn; INSTANCE:3]")
								Sleep(500)
								ControlClick("Поиск по словарю/Рубрикатору", "", "[CLASS:TBitBtn; INSTANCE:1]")
							EndIf
							$hWnd1 = WinWaitActive($IrbisTit, "", 5)
							If $hWnd1 Then
								ControlSend($IrbisTit, "", "[CLASS:THSHintTntComboBox.UnicodeClass; INSTANCE:1]", "{HOME}" & "{DOWN}")
							EndIf



						Case "спек" ; "спек" - переделывает описание в SPEC
							If OpenElement(900) Then
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
									If OpenElement(920) Then
										Send("{HOME}" & "{DOWN 2}" & "{ENTER}")
									EndIf
								EndIf
								$hWnd = WinWaitActive($IrbisTit, "", 5)
								If $hWnd Then
									ControlSend($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:2]", "{HOME}" & "{DOWN 2}")
								EndIf
							EndIf




						Case "эбс" ; "эбс" - создание 830 поля и внесение сведений о копии в ЭБ
							$invNum = GetInvNum()
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
										GoToField(830)
										Send("^{ENTER}")
									EndIf
									ClipMan("^0w=^!ЭБ^a" & $invNum)
								EndIf
							EndIf

						Case "изм" ; "изм" - создание 830 поля и внесение сведений об изменении индекса МДА и авт. зн.
							GoToField(686)
							_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
							$indMDA = WinGetText($IrbisTit, "")
							$pos1 = StringInStr($indMDA, "Page7", 0, 1) + 6
							$indMDA = (StringMid($indMDA, $pos1, 3))

							GoToField(908)
							$avtZn = WinGetText($IrbisTit, "")
							$pos1 = StringInStr($avtZn, "Page7", 0, 1) + 6
							$text = StringMid($avtZn, $pos1, 3)
							If StringInStr($text, ' ') = 0 Then
								$avtZnLen = 3
							Else
								$avtZnLen = 4
							EndIf
							$avtZn = StringMid($avtZn, $pos1, $avtZnLen)
							$invNum = GetInvNum()
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
										GoToField(830)
										Send("^{ENTER}")
									EndIf
								EndIf
								$hWnd = WinWaitActive($IrbisTit, "", 5)
								If $hWnd1 Then
									ClipMan("^007^!07^aИнв. " & $invNum & ": до " & _NowDate() & " индекс МДА " & $indMDA & ", авт. знак " & $avtZn)
								EndIf
							EndIf
						Case "конв" ; "конв" - создание 830 поля и внесение сведений о прежнем вхождении в конволют
							$invNum = GetInvNum()
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
										GoToField(830)
										Send("^{ENTER}")
									EndIf
								EndIf
								$hWnd = WinWaitActive($IrbisTit, "", 5)
								If $hWnd1 Then
									ClipMan("^007^!07^aИнв. " & $invNum & ": до " & _NowDate() & " входил в конволют (инв. )")
								EndIf
							EndIf


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
							Run ("RunDLL32.EXE shell32.dll,ShellExec_RunDLL d:\РПК.pdf")
						Case "сокр"
							Run ("RunDLL32.EXE shell32.dll,ShellExec_RunDLL d:\dESCTOP\ГОСТ_7.0.12-2011_Сокращ_слов.pdf")
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
						Case "рег"
							Local $oExcel = _Excel_Open()
							_Excel_BookOpen(_Excel_Open(), "d:\dESCTOP\Регистрация книг.xls")

						Case "скан"
							Local $oWord = _Word_Create()
							_Word_DocOpen($oWord, "d:\dESCTOP\СканКоп.rtf")
							Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
							WinActivate($hWnd)
						Case "веб" ; "веб" - Электронный каталог IRBIS 2.0
							Opt("WinTitleMatchMode", 2)

							$path = @ProgramFilesDir & "\Mozilla Firefox\firefox.exe"
							$url = "http://194.169.10.3/jirbis2/index.php?option=com_irbis&view=irbis&Itemid=108"
							Run($path & " " & $url, "", @SW_MAXIMIZE)


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
	EndIf
EndFunc   ;==>Field

;						CTRL+F Поиск по виду основного словаря. Читает введенную строку и выполняет поиск
Func Search()
	If HotKeyOn("^f", "Search") Then

		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
		$input = InputBox("Выполнить", "Поиск по:", "", "", 190, 130)
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)

		If WinWaitActive($IrbisTit, "", 5) Then
			Do
				$exit = 0
				$inputTest = TestInput($input)
				If IsArray($inputTest) Then
					$SPLIT = $inputTest
					$input = $SPLIT[1]
				Else
					$input = $inputTest
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
	EndIf
EndFunc   ;==>Search

;						ALT+H Последовательный поиск
; Ввести значение. Нажать TAB и ввести номер поля. Нажать еще раз TAB и выбрать уточняемый запрос. Нажать ENTER
Func CohSearch()
	If HotKeyOn("!h", "CohSearch") Then
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

;						CTRL+W Печать основной карточки. Открытие существующего файла (путь: c:\irbiswrk\) или создание нового.

;						CTRL+K Печать контрольной карточки

;						CTRL+D Печать формуляра

;						CTRL+Y Печать ярлычка

;						CTRL+M Печать ярлычка для мягких формуляров
Func FormularLabel()
	$wTit = WinGetTitle("[ACTIVE]")
	$isIrbis = StringInStr($wTit, $IrbisTit)
	If $isIrbis = 0 Then
		HotKeySet("^+m")
		Send("^+m")
		HotKeySet("^+m", "FormularLabel")
	Else
;~ 		Print("formularLabel")
	EndIf
EndFunc   ;==>FormularLabel


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

;						CTRL+SPACE Фокус на окне полного описания
Func ViewFocus()
	If HotKeyOn("^{SPACE}", "ViewFocus") Then
		ControlClick($IrbisTit, "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "left", 1, 1034, 25)
	EndIf
EndFunc   ;==>ViewFocus

;						CTRL+SHIFT+K Копировать отмеченные поля в буферную запись
Func CopySelected()
	If HotKeyOn("^+k", "CopySelected") Then
		;~ 		Предотвращение залипания SHIFT'a
		Sleep(10)
		Send("{SHIFTDOWN}")
		Sleep(10)
		Send("{SHIFTUP}")
		ControlFocus($IrbisTit, "", "[CLASS:TTntStringGrid.UnicodeClass; INSTANCE:3]")
		Sleep(100)
		Send("{APPSKEY}{UP 8}{ENTER}")
	EndIf
EndFunc   ;==>CopySelected

;~ Переход на поле по номеру
Func GoToField($com)
	Send("!q")
	Sleep(100)
	Send($com & "{ENTER}")
EndFunc   ;==>GoToField

;~ Раскрытие поля при нажатии F2
Func OpenElement($com)
	Send("!q")
	Sleep(100)
	Send($com & "{ENTER}" & "{F2}")
	$hWnd = WinWaitActive('Элемент: "' & $com, "", 5)
	If $hWnd Then
		ControlClick($hWnd, "", "[CLASS:TTntRichEdit.UnicodeClass; INSTANCE:1]", "left", 1, 1, 1)
		Return $hWnd
	EndIf
EndFunc   ;==>OpenElement

;~ Раскрытие поля при нажатии F3
Func OpenElementF3($com)
	Send("!q")
	Sleep(100)
	Send($com & "{ENTER}" & "{F3}")
	$hWnd = WinWaitActive('Элемент:  "' & $com, "", 5)
	If $hWnd Then
		ControlClick($hWnd, "", "[CLASS:TTntRichEdit.UnicodeClass; INSTANCE:1]", "left", 1, 12, 12)
		Return $hWnd
	EndIf
EndFunc   ;==>OpenElementF3

;~ Получение инв. номера
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

;~ Изменение вида словаря
Func Srchfor($srch)
	Send("!f")
	If WinWaitActive("Вид основного словаря", "", 5) Then
		ClipMan($srch)
		Send("{ENTER}")
		If WinWaitActive($IrbisTit, "", 5) Then
			Sleep(100)
			Send("!d")
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
			Sleep(100)
			Send("!d")
			Sleep(100)
			$string = ""
			_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0419)
			For $i = 2 To $SPLIT[0]
				If $i == $SPLIT[0] Then
					$string = $string & $SPLIT[$i]
				Else
					$string = $string & $SPLIT[$i] & ' '
				EndIf
			Next
			_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
			Sleep(100)
			ClipMan($string)
			Send("{ENTER}")
		EndIf
	EndIf

EndFunc   ;==>SrchforEx

;~ Использование буфера обмена с сохранением текущего состояния буфера
Func ClipMan($com)
	$clip = ClipGet()
	ClipPut($com)
	Send("+{INS}")
	ClipPut($clip)
EndFunc   ;==>ClipMan

;~ Проверка введенной строки
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

;~ Вывод на печать
Func Print()
	If HotKeyOn(@HotKeyPressed, "Print") Then
		_WinAPI_SetKeyboardLayout(WinGetHandle(AutoItWinGetTitle()), 0x0409)
;~ 		Получение названия базы и сравнение
		$bdName = ControlGetText($IrbisTit, "", "[CLASS:THSHintComboBox; INSTANCE:4]")
		$isMPDABD = StringInStr($bdName, "MPDA - Московская Православная Духовная Академия")
		$isDSTBD = StringInStr($bdName, "DST - Диссертационная база МДА")
		$isPRBD = StringInStr($bdName, "PR - Периодические издания (с 2014 г.)")
		$cancel = 1
		Switch @HotKeyPressed
			Case "^w"
				If $isMPDABD Then
					$count = 9
				ElseIf $isDSTBD Then
					$count = 5
				ElseIf $isPRBD Then
					$count = 1
				Else
					$cancel = 0
				EndIf
				$invNum = GetInvNum()
				$filePath = "c:\irbiswrk\" & $invNum & ".RTF"
				$SecondfilePath = "c:\irbiswrk\Сделаны\" & $invNum & ".RTF"
				If FileExists($filePath) Or FileExists($SecondfilePath) Then
					$ans = MsgBox(67, "Внимание", "Файл существует. Открыть его?")
					If $ans = 6 Then
						$cancel = 0
						If FileExists($filePath) Then
							Local $oWord = _Word_Create()
							_Word_DocOpen($oWord, $filePath)
							Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
							WinActivate($hWnd)

						ElseIf FileExists($SecondfilePath) Then
							Local $oWord = _Word_Create()
							_Word_DocOpen($oWord, $SecondfilePath)
							Local $hWnd = WinWait("[CLASS:OpusApp]", "", 10)
							WinActivate($hWnd)
						EndIf
					ElseIf $ans = 7 Then
						$cancel = 1
					Else
						$cancel = 0
					EndIf
				EndIf
				If $isPRBD Then
					$macro = ''
				Else
					$macro = "Макрос1"
				EndIf

			Case "^k"
				If $isMPDABD Then
					$count = 12
				ElseIf $isDSTBD Then
					$count = 10
				Else
					$cancel = 0
				EndIf
				$macro = "KK"

			Case "^d"
				$count = 6
				If $isMPDABD Then
					$count = 6
				ElseIf $isDSTBD Then
					$count = 4
				Else
					$cancel = 0
				EndIf
				$macro = "Formular"

			Case "^y"
				GoToField(920)
				$text = WinGetText($IrbisTit)
				If $isMPDABD Then
					If StringInStr($text, "SPEC") Then
						$ans = MsgBox(67, "Внимание", "Печатать номера томов?")
						If $ans = 6 Then
							$count = 3
						ElseIf $ans = 7 Then
							$count = 4
						Else
							$cancel = 0
						EndIf
					Else
						$count = 3
					EndIf
				ElseIf $isDSTBD Then
					$count = 3
				Else
					$cancel = 0
				EndIf
				$macro = "Label_2"

			Case "^m"
				If $isMPDABD Then
					$count = 5
					$macro = ''
				Else
					$cancel = 0
				EndIf
		EndSwitch

		If $cancel Then
			WinActivate($IrbisTit)
			ControlFocus($IrbisTit, "", "[CLASS:TTntStringGrid.UnicodeClass; INSTANCE:3]")
			If @HotKeyPressed == "^w" And $isPRBD = 0 Then
				ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 34, 12)
			Else
				ControlClick($IrbisTit, "", "[CLASS:TToolBar; INSTANCE:1]", "left", 1, 11, 12)
			EndIf
			$hWnd1 = WinWaitActive("Печать", "", 5)
			If $hWnd1 Then
				ControlSend("Печать", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{HOME}")
				For $i = 1 To $count
					ControlSend("Печать", "", "[CLASS:THSHintComboBox; INSTANCE:1]", "{DOWN}")
				Next
				ControlClick("Печать", "", "[CLASS:TBitBtn; INSTANCE:2]")
			EndIf
			$hWnd1 = WinWaitActive("Файл", "", 5)
			If $hWnd1 Then
				If @HotKeyPressed == "^w" And $isPRBD = 0 Then
					ControlSend("Файл", "", "[CLASS:Edit; INSTANCE:1]", $invNum)
				Else
					$numDoc = 1
					If WinExists("1 [Режим") Then
						If WinExists("2 [Режим") Then
							$numDoc = 3
						Else
							$numDoc = 2
						EndIf
					EndIf
					ControlSend("Файл", "", "[CLASS:Edit; INSTANCE:1]", $numDoc)
				EndIf
				Send("{TAB 2}" & "{ENTER}")
			EndIf
			$hWnd1 = WinWaitActive("Подтвердить", "", 1)
			If $hWnd1 Then
				ControlClick("Подтвердить", "", "[CLASS:Button; INSTANCE:1]")
			EndIf

			$hWnd1 = WinWaitActive("Внимание", "", 5)
			If $hWnd1 Then
				ControlClick("Внимание", "", "[CLASS:Button; INSTANCE:1]")
			EndIf
			If $macro <> '' Then
				$hWnd1 = WinWaitActive("[CLASS:OpusApp]", "", 10)
				If $hWnd1 Then
					WinClose("Печать")
					$Object = ObjGet("", "Word.Application")
					$Object.Run($macro)
				EndIf
			Else
				WinClose("Печать")
			EndIf
		EndIf
	EndIf
EndFunc   ;==>Print

;~ Отключение сочетаний клавиш в других приложениях
Func HotKeyOn($send, $func)
	If WinGetHandle("[ACTIVE]") == WinGetHandle($IrbisTit) Then

;~ 		Предотвращение залипания CTRL'a
		Sleep(10)
		Send("{CTRLDOWN}")
		Sleep(10)
		Send("{CTRLUP}")
		Return True
	Else
		HotKeySet($send)
		Send($send)
		HotKeySet($send, $func)
		Return False
	EndIf
EndFunc   ;==>HotKeyOn
