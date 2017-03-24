'Option Explicit

Sub Main
	Dim oBook, oDoc
	Dim Path$
	Dim oTxt, oVcurs, oTcurs 
	Dim GroupNum%
	Dim counter%
	
	GlobalScope.BasicLibraries.loadLibrary("Tools")

	oBook = ThisComponent
	Path = DirectoryNameOutOfPath(oBook.getURL(),"/") & "/group.odt"
	oDoc = OpenDocument(Path, NoArgs())

	oDoc.Text.String = ""
	
	GroupNum = 1

	For counter = 1 To oBook.Sheets.count - 1
	
		Dim sName$, Manuf$, BodyMat$, Dn$, Pn$, Env$, T1$, T2$, Tw$, Pw$, myTable, oColSeps
		Dim LastColumn%

		oTxt = oDoc.text 'getText()	
		LastColumn  = Func.GetLastUsedColumn(oBook.Sheets.getByIndex(counter), 14)
		
		sName = oBook.Sheets(counter).getCellByPosition(1, 4).getString()
		Manuf = oBook.Sheets(counter).getCellByPosition(1, 5).getString()
		BodyMat = oBook.Sheets(counter).getCellByPosition(1, 6).getString()		
		Dn = oBook.Sheets(counter).getCellByPosition(1, 7).getString()
		Pn = oBook.Sheets(counter).getCellByPosition(1, 8).getString()
		Env = oBook.Sheets(counter).getCellByPosition(1, 9).getString()		
		T1 = oBook.Sheets(counter).getCellByPosition(1, 10).getString()
		T2 = oBook.Sheets(counter).getCellByPosition(2, 10).getString()
		Pw = oBook.Sheets(counter).getCellByPosition(1, 11).getString()
		Pw = Pw & " кгс/см" & Chr(178)
		
		If oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsInch").State = 1 Then
			Dn = Dn & Chr(34)
		else
			Dn = Dn & " мм"
		End If
	
		If  oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsAnsi").State = 1 Then
			Pn = Pn & "#"
		else
			Pn = Pn & " кгс/см"
		End If
		
		If oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsAvg").State = 1 Then
			Tw = "Окружающей среды"
		else
			If T2 <> "" Then
				Tw = T1 & " " & Chr(247) & " " & T2 & " " & Chr(176) & "C"
			else
				Tw = T1 & " " & Chr(176) & "C"
			End If
		End If

		oVcurs = oDoc.CurrentController.getViewCursor()
		'oTcurs = oVcurs.getText().createTextCursorByRange(oVcurs)
		'oTxt.createTextCursorByRange(oVcurs.getEnd())
		
		oVcurs.ParaStyleName = "Header1"
		oTxt.insertString(oTxt.End, "ГРУППА " & GroupNum & Chr(13), FALSE)	

		InsertText(oDoc, "Наименование арматуры: ", sName, ";" & Chr(13))
		InsertText(oDoc, "Изготовитель: ", Manuf, ";" & Chr(13))
		InsertText(oDoc, "Материал корпуса: ", BodyMat, ";" & Chr(13))
		InsertText(oDoc, "Номинальный проход: ", Dn, ";" & Chr(13))
		InsertText(oDoc, "Номиналное давление: ", Pn, ";" & Chr(13))
		InsertText(oDoc, "Рабочая среда: ", Env, ";" & Chr(13))
		InsertText(oDoc, "Температура рабочей среды: ", Tw, ";" & Chr(13))		
		InsertText(oDoc, "Давление рабочее: ", Pw, "." & Chr(13))		
		
		oVcurs.ParaStyleName = "Header2"
		oTxt.insertString(oTxt.End, "Технологические параметры эксплуатации", FALSE)
			
		myTable = Func.CreateTable(oDoc, LastColumn + 1, 5)
		oTxt.insertTextContent(oTxt.End, myTable, false)
		oColSeps = MyTable.TableColumnSeparators
		oColSeps(0).Position = 700
		oColSeps(1).Position = 2600
		oColSeps(2).Position = 4300
		oColSeps(3).Position = 8600
		MyTable.TableColumnSeparators = oColSeps
		myTable.getCellByName("A1").setString("№" & Chr(13) & "п/п")
		myTable.getCellByName("B1").setString("Заводской" & Chr(13) & "(идентификационный) номер")
		myTable.getCellByName("C1").setString("Дата изготовления, год")
		myTable.getCellByName("D1").setString("Место установки(наименование, индекс трубопровода)")	
		myTable.getCellByName("E1").setString("Дата установки, год")	
		
		For i = 1 To LastColumn
			myTable.getCellByPosition(0, i).setString(i)
			myTable.getCellByPosition(1, i).setString(oBook.Sheets(counter).getCellByPosition(i, 15).getString())
			myTable.getCellByPosition(2, i).setString(oBook.Sheets(counter).getCellByPosition(i, 16).getString())
			myTable.getCellByPosition(3, i).setString(oBook.Sheets(counter).getCellByPosition(i, 18).getString())
			myTable.getCellByPosition(4, i).setString(oBook.Sheets(counter).getCellByPosition(i, 17).getString())
		Next
		
		GroupNum = GroupNum + 1
	
	Next counter
		
End Sub

Function InsertText(oDocument, Text1$, Text2$, Text3$)
	Dim oCursor
	
	If Text2 = "" Then Exit Function
	oCursor = oDocument.text.End
	oCursor.ParaStyleName = "Text1"
	oDocument.text.insertString(oCursor, Text1, False)
	oCursor.CharStyleName = "Text2"
	oDocument.text.insertString(oCursor, Text2, False)
	oCursor.CharStyleName = "Базовый"
	oDocument.text.insertString(oCursor, Text3, False)	
End Function

