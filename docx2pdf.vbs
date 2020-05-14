Option Explicit

Const PDF = 17
Const fTitle = "���������� ������� "
Const cTitle = " ������ �� "
Const tGboy = "���� ��� �. ������������� �. �. ���������� ��������� ���."
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

' ������� ���������
Function Rus2Lat(strRus)
	Dim i
	Dim strTemp
	Dim strLat
	For i = 1 To Len(strRus)
		strTemp = Mid(strRus, i, 1)			 
		Select Case strTemp
			Case "�"
				strLat = strLat & "a"
			Case "�"
				strLat = strLat & "a"
			Case "�"
				strLat = strLat & "b"
			Case "�"
				strLat = strLat & "b"
			Case "�"
				strLat = strLat & "v"
			Case "�"
				strLat = strLat & "v"
			Case "�"
				strLat = strLat & "g"
			Case "�"
				strLat = strLat & "g"
			Case "�"
				strLat = strLat & "d"
			Case "�"
				strLat = strLat & "d"
			Case "�"
				strLat = strLat & "e"
			Case "�"
				strLat = strLat & "e"
			Case "�"
				strLat = strLat & "e"
			Case "�"
				strLat = strLat & "e"
			Case "�"
				strLat = strLat & "zh"
			Case "�"
				strLat = strLat & "zh"
			Case "�"
				strLat = strLat & "z"
			Case "�"
				strLat = strLat & "z"
			Case "�"
				strLat = strLat & "i"
			Case "�"
				strLat = strLat & "i"
			Case "�"
				strLat = strLat & "i"
			Case "�"
				strLat = strLat & "i"
			Case "�"
				strLat = strLat & "k"
			Case "�"
				strLat = strLat & "k"
			Case "�"
				strLat = strLat & "l"
			Case "�"
				strLat = strLat & "l"
			Case "�"
				strLat = strLat & "m"
			Case "�"
				strLat = strLat & "m"
			Case "�"
				strLat = strLat & "n"
			Case "�"
				strLat = strLat & "n"
			Case "�"
				strLat = strLat & "o"
			Case "�"
				strLat = strLat & "o"
			Case "�"
				strLat = strLat & "p"
			Case "�"
				strLat = strLat & "p"
			Case "�"
				strLat = strLat & "r"
			Case "�"
				strLat = strLat & "r"
			Case "�"
				strLat = strLat & "s"
			Case "�"
				strLat = strLat & "s"
			Case "�"
				strLat = strLat & "t"
			Case "�"
				strLat = strLat & "t"
			Case "�"
				strLat = strLat & "u"
			Case "�"
				strLat = strLat & "u"
			Case "�"
				strLat = strLat & "f"
			Case "�"
				strLat = strLat & "f"
			Case "�"
				strLat = strLat & "kh"
			Case "�"
				strLat = strLat & "kh"
			Case "�"
				strLat = strLat & "ts"
			Case "�"
				strLat = strLat & "ts"
			Case "�"
				strLat = strLat & "ch"
			Case "�"
				strLat = strLat & "ch"
			Case "�"
				strLat = strLat & "sh"
			Case "�"
				strLat = strLat & "sh"
			Case "�"
				strLat = strLat & "sch"
			Case "�"
				strLat = strLat & "sch"
			Case "�"
				strLat = strLat & ""
			Case "�"
				strLat = strLat & ""
			Case "�"
				strLat = strLat & "y"
			Case "�"
				strLat = strLat & "y"
			Case "�"
				strLat = strLat & ""
			Case "�"
				strLat = strLat & ""
			Case "�"
				strLat = strLat & "e"
			Case "�"
				strLat = strLat & "e"
			Case "�"
				strLat = strLat & "yu"
			Case "�"
				strLat = strLat & "yu"
			Case "�"
				strLat = strLat & "ya"
			Case "�"
				strLat = strLat & "ya"
			case "�"
				strLat = strLat & ""
			case "�"
				strLat = strLat & ""
			case " "
				strLat = strLat & "-"
			Case Else
				strLat = strLat & strTemp
		End Select
	Next
	Rus2Lat = strLat
End Function

' ���� � ������� ���� ���������
If WScript.Arguments.Count = 1 Then
	rsDate = ""
	' ������ �������� ������ ���� ������, ������� ����� ������������.
	strSourceFolder = WScript.Arguments.Item(0)
	' ������ ������ ��� ������ � �������� ��������
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	' ���� ����� ����������
	If objFSO.FolderExists(strSourceFolder) Then
		Set objWord = Nothing
		fCount = 0
		Set csvFile = objFSO.CreateTextFile(strSourceFolder & "\csv.csv", True)
		For Each objFile In objFSO.GetFolder(strSourceFolder).Files
			If StrComp(objFSO.GetExtensionName(objFile.Name), "docx", vbTextCompare) = 0 Then
				' ��������� Word ���� �� ��� �� �������
				If objWord Is Nothing Then
					Set objWord = WScript.CreateObject("Word.Application")
				End If
				' ������ ���������
				docTitle = ""
				' ��������� ��������
				Set objDocument = objWord.Documents.Open(objFile.Path)
				' �������� ������ ������ ���������
				Set customProp = objDocument.BuiltinDocumentProperties
				' �������� ����
				rsDate = objFSO.GetBaseName(strSourceFolder) & "." & objFSO.GetExtensionName(strSourceFolder)
				' �������� ���������
				docTitle = fTitle & objFSO.GetBaseName(objFile.Name) & cTitle & rsDate
				
				' ���������� �������� ���������
				For Each prop in customProp
					' ������������� ������ �������� ���������
					Select case prop.Name
						' ��������� ���������
						case "Title"
							prop.Value = docTitle & " " & tGboy
						' ���� ���������
						case "Subject"
							prop.Value = docTitle & " " & tGboy
						' ����� ���������
						case "Author"
							prop.Value = tGboy
						' ��������
						case "Company"
							prop.Value = tGboy
					End Select
				Next
				' ��������� �������� ��� PDF. �������� ����� ����� ��� ����������
				' ��� �� ������� ����������� ��� �������� ����� ������������.
				objDocument.SaveAs2 objFSO.BuildPath(objFile.ParentFolder.Path, Rus2Lat(objFSO.GetBaseName(objFile.Name)) & ".pdf"), PDF
				' ���������� ������ � csv ����
				csvText = """" & docTitle & """;""" & assetsFolder & assetsType & "/" & rsDate & "/" & Rus2Lat(objFSO.GetBaseName(objFile.Name)) & ".pdf"""
				csvFile.WriteLine(csvText)
				' ��������� ��������
				objDocument.Close
				' �������� ����������
				' Set objDocument = Nothing
				fCount = fCount + 1
			End If
		Next
		' ���� Word ������� - ������� ���
		If Not objWord Is Nothing Then
			objWord.Quit
		End If
		' �������� ����������
		Set objWord = Nothing
		' ��������� csv ����
		csvFile.Close
		' ����� ��������� � ���������� ������������ ������
		MsgBox "���������� " & fCount & " ������"
	End If
	' �������� ����������
	Set objFSO = Nothing
Else
	MsgBox "Not found parametrs"
End If
' ������� �� ���������� ��������.
WScript.Quit 0	
