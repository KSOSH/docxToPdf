Option Explicit

Const PDF = 17
Const fTitle = "РАСПИСАНИЕ ЗАНЯТИЙ "
Const cTitle = " КЛАССА НА "
Const tGboy = "ГБОУ СОШ п. Комсомольский м. р. Кинельский Самарской обл."
Const assetsFolder = "assets/files/0000/do/"
Const assetsType = "rs"
Dim docTitle
Dim objWord
Dim objDocument
Dim strSourceFolder
Dim objFSO
Dim objFile
Dim customProp
Dim prop
Dim fCount
Dim csvFile
Dim csvText
Dim rsDate

' Функция транслита
Function Rus2Lat(strRus)
	Dim i
	Dim strTemp
	Dim strLat
	For i = 1 To Len(strRus)
		strTemp = Mid(strRus, i, 1)			 
		Select Case strTemp
			Case "а"
				strLat = strLat & "a"
			Case "А"
				strLat = strLat & "a"
			Case "б"
				strLat = strLat & "b"
			Case "Б"
				strLat = strLat & "b"
			Case "в"
				strLat = strLat & "v"
			Case "В"
				strLat = strLat & "v"
			Case "г"
				strLat = strLat & "g"
			Case "Г"
				strLat = strLat & "g"
			Case "д"
				strLat = strLat & "d"
			Case "Д"
				strLat = strLat & "d"
			Case "е"
				strLat = strLat & "e"
			Case "Е"
				strLat = strLat & "e"
			Case "ё"
				strLat = strLat & "e"
			Case "Ё"
				strLat = strLat & "e"
			Case "ж"
				strLat = strLat & "zh"
			Case "Ж"
				strLat = strLat & "zh"
			Case "з"
				strLat = strLat & "z"
			Case "З"
				strLat = strLat & "z"
			Case "и"
				strLat = strLat & "i"
			Case "И"
				strLat = strLat & "i"
			Case "й"
				strLat = strLat & "i"
			Case "Й"
				strLat = strLat & "i"
			Case "к"
				strLat = strLat & "k"
			Case "К"
				strLat = strLat & "k"
			Case "л"
				strLat = strLat & "l"
			Case "Л"
				strLat = strLat & "l"
			Case "м"
				strLat = strLat & "m"
			Case "М"
				strLat = strLat & "m"
			Case "н"
				strLat = strLat & "n"
			Case "Н"
				strLat = strLat & "n"
			Case "о"
				strLat = strLat & "o"
			Case "О"
				strLat = strLat & "o"
			Case "п"
				strLat = strLat & "p"
			Case "П"
				strLat = strLat & "p"
			Case "р"
				strLat = strLat & "r"
			Case "Р"
				strLat = strLat & "r"
			Case "с"
				strLat = strLat & "s"
			Case "С"
				strLat = strLat & "s"
			Case "т"
				strLat = strLat & "t"
			Case "Т"
				strLat = strLat & "t"
			Case "у"
				strLat = strLat & "u"
			Case "У"
				strLat = strLat & "u"
			Case "ф"
				strLat = strLat & "f"
			Case "Ф"
				strLat = strLat & "f"
			Case "х"
				strLat = strLat & "kh"
			Case "Х"
				strLat = strLat & "kh"
			Case "ц"
				strLat = strLat & "ts"
			Case "Ц"
				strLat = strLat & "ts"
			Case "ч"
				strLat = strLat & "ch"
			Case "Ч"
				strLat = strLat & "ch"
			Case "ш"
				strLat = strLat & "sh"
			Case "Ш"
				strLat = strLat & "sh"
			Case "щ"
				strLat = strLat & "sch"
			Case "Щ"
				strLat = strLat & "sch"
			Case "ъ"
				strLat = strLat & ""
			Case "Ъ"
				strLat = strLat & ""
			Case "ы"
				strLat = strLat & "y"
			Case "Ы"
				strLat = strLat & "y"
			Case "ь"
				strLat = strLat & ""
			Case "Ь"
				strLat = strLat & ""
			Case "э"
				strLat = strLat & "e"
			Case "Э"
				strLat = strLat & "e"
			Case "ю"
				strLat = strLat & "yu"
			Case "Ю"
				strLat = strLat & "yu"
			Case "я"
				strLat = strLat & "ya"
			Case "Я"
				strLat = strLat & "ya"
			case "«"
				strLat = strLat & ""
			case "»"
				strLat = strLat & ""
			case " "
				strLat = strLat & "-"
			Case Else
				strLat = strLat & strTemp
		End Select
	Next
	Rus2Lat = strLat
End Function

' Если у скрипта есть аргументы
If WScript.Arguments.Count = 1 Then
	rsDate = ""
	' Первый аргумент должен быть папкой, которую будем обрабатывать.
	strSourceFolder = WScript.Arguments.Item(0)
	' Создаём объект для работы с файловой системой
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	' Если папка существует
	If objFSO.FolderExists(strSourceFolder) Then
		Set objWord = Nothing
		fCount = 0
		Set csvFile = objFSO.CreateTextFile(strSourceFolder & "\csv.csv", True)
		For Each objFile In objFSO.GetFolder(strSourceFolder).Files
			If StrComp(objFSO.GetExtensionName(objFile.Name), "docx", vbTextCompare) = 0 Then
				' Запускаем Word если он ещё не запущен
				If objWord Is Nothing Then
					Set objWord = WScript.CreateObject("Word.Application")
				End If
				' Пустой заголовок
				docTitle = ""
				' Открываем документ
				Set objDocument = objWord.Documents.Open(objFile.Path)
				' Получаем объект свойст документа
				Set customProp = objDocument.BuiltinDocumentProperties
				' Получаем дату
				rsDate = objFSO.GetBaseName(strSourceFolder) & "." & objFSO.GetExtensionName(strSourceFolder)
				' Собираем заголовок
				docTitle = fTitle & objFSO.GetBaseName(objFile.Name) & cTitle & rsDate
				
				' Перебираем свойства документа
				For Each prop in customProp
					' Устанавливаем нужные свойства документа
					Select case prop.Name
						' Заголовок документа
						case "Title"
							prop.Value = docTitle & " " & tGboy
						' Тема документа
						case "Subject"
							prop.Value = docTitle & " " & tGboy
						' Автор документа
						case "Author"
							prop.Value = tGboy
						' Компания
						case "Company"
							prop.Value = tGboy
					End Select
				Next
				' Сохраняем документ как PDF. Транслит имени файла для сохранения
				' Так же сначало сохраниться сам документ перед конвертацией.
				objDocument.SaveAs2 objFSO.BuildPath(objFile.ParentFolder.Path, Rus2Lat(objFSO.GetBaseName(objFile.Name)) & ".pdf"), PDF
				' Записываем данные в csv файл
				csvText = """" & docTitle & """;""" & assetsFolder & assetsType & "/" & rsDate & "/" & Rus2Lat(objFSO.GetBaseName(objFile.Name)) & ".pdf"""
				csvFile.WriteLine(csvText)
				' Закрываем документ
				objDocument.Close
				' Обнуляем переменную
				' Set objDocument = Nothing
				fCount = fCount + 1
			End If
		Next
		' Если Word запущен - закроем его
		If Not objWord Is Nothing Then
			objWord.Quit
		End If
		' Обнуляем переменную
		Set objWord = Nothing
		' Закрываем csv файл
		csvFile.Close
		' Вывод сообщения о количестве обработанных файлов
		MsgBox "Обработано " & fCount & " файлов"
	End If
	' Обнуляем переменную
	Set objFSO = Nothing
Else
	MsgBox "Not found parametrs"
End If
' Выходим из выполнения сценария.
WScript.Quit 0	
