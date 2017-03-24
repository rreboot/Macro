Function GetLastUsedColumn(oSheet, RowNumber) As Integer		
	Dim oRow, oEmptyRanges, Count%		
	oRow = oSheet.getRows().getByIndex(RowNumber)
	oEmptyRanges = oRow.queryEmptyCells
	Count = oEmptyRanges.Count		
	GetLastUsedColumn = oEmptyRanges.RangeAddresses(Count-1).StartColumn - 1
End Function

Function CreateTable(document, rows%, cols%) As Object
	Dim oTextTable	
	oTextTable = document.createInstance("com.sun.star.text.TextTable")
	oTextTable.initialize(rows, cols)  
	oTextTable.HoriOrient = 0 'com.sun.star.text.HoriOrientation::NONE
	oTextTable.LeftMargin = 0
	oTextTable.RightMargin = 0	
	CreateTable = oTextTable
End Function

Sub CopySheet
	Dim oBook, Count%, sName$
	oBook = ThisComponent
	Count = oBook.Sheets.Count
	sName = oBook.CurrentController.ActiveSheet.Name
	oBook.Sheets.copyByName(sName, "Группа " & Count, Count)
	oBook.CurrentController.setActiveSheet(oBook.Sheets(Count))
End Sub

