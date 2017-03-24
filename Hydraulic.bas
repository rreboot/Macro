'Option Explicit

Sub Main
	Dim oBook, oDoc
	Dim Path$
	Dim counter%, id%
	Dim ustanovka$, obj$
	Dim a(0, 9), Nrow
	Dim MyFormat$
	
	GlobalScope.BasicLibraries.loadLibrary("Tools")

	oBook = ThisComponent
	Path = DirectoryNameOutOfPath(oBook.getURL(),"/") & "/gi.odt"
	oDoc = OpenDocument(Path, NoArgs())
	
	oTable = oDoc.TextTables(0)
		
	oTable.Rows.removeByIndex(3, oTable.Rows.Count - 3)
	oTable.getCellRangeByName("A3:J3").SetDataArray(Array(Array("", "", "", "", "", "", "", "", "", ""))
	
	ustanovka = oBook.Sheets(1).getCellByPosition(1, 2).getString()
	obj = oBook.Sheets(1).getCellByPosition(1, 1).getString()	
	oDoc.Bookmarks.getByName("установка").Anchor.setString(ustanovka)
	oDoc.Bookmarks.getByName("объект").Anchor.setString(obj)		
	
	id = 1
	Nrow = 0
	MyFormat = "0.0"
	
	For counter = 1 To oBook.Sheets.count - 1
	
		Dim sName$, Dn$, Pn$, Manuf$, Status$, Pw, P1, P2
		Dim LastColumn%

		LastColumn  = Func.GetLastUsedColumn(oBook.Sheets.getByIndex(counter), 14)
		
		sName = oBook.Sheets(counter).getCellByPosition(1, 4).getString()
		Dn = oBook.Sheets(counter).getCellByPosition(1, 7).getString()
		Pn = oBook.Sheets(counter).getCellByPosition(1, 8).getString()
		Manuf = oBook.Sheets(counter).getCellByPosition(1, 5).getString()
		Status = "Работоспособна"
		Pw = oBook.Sheets(counter).getCellByPosition(1, 11).getString()

		If oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsInch").State = 1 Then
			Dn = Dn & Chr(34)
		End If
	
		If  oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsAnsi").State = 1 Then
			Pn = Pn & "#"
		End If	

		Pw = CDbl(Pw)
		If Pw < 2 Then
			P1 = "2,0"			
		Else
			P1 = Format(Pw * 1.25, MyFormat)
			
		End If
		P2 = Format(Pw, MyFormat)			

		For i = 1 To LastColumn
			
			Nmanuf = oBook.Sheets(counter).getCellByPosition(i, 15).getString()
			InstPlace = oBook.Sheets(counter).getCellByPosition(i, 18).getString()

			ReDim Preserve a(Nrow, 9)
						
			a(Nrow, 0) = id
			a(Nrow, 1) = sName
			a(Nrow, 2) = Dn
			a(Nrow, 3) = Pn
			a(Nrow, 4) = Manuf
			a(Nrow, 5) = Nmanuf
			a(Nrow, 6) = InstPlace
			a(Nrow, 7) = Status
			a(Nrow, 8) = P1
			a(Nrow, 9) = P2
			
			Nrow = Nrow + 1
			id = id + 1			
			
		Next i	
	Next counter
	
	oTable.Rows.insertByIndex(oTable.Rows.Count, Nrow - 1)
	oTable.getCellRangeByPosition(0, 2, 9, oTable.Rows.Count - 1).SetDataArray(a())
End Sub

