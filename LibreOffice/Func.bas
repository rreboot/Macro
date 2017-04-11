Function GetLastUsedColumn(oSheet, RowNumber) As Integer		
Dim oRow, oEmptyRanges, Count%		
	oRow = oSheet.getRows().getByIndex(RowNumber)
	oEmptyRanges = oRow.queryEmptyCells
	Count = oEmptyRanges.Count		
	GetLastUsedColumn = oEmptyRanges.RangeAddresses(Count-1).StartColumn - 1
End Function

Function GetLastUsedRow(oSheet, ColNumber) As Integer		
Dim oCol, oEmptyRanges, Count%		
	oCol = oSheet.getColumns().getByIndex(ColNumber)
	oEmptyRanges = oCol.queryEmptyCells
	Count = oEmptyRanges.Count		
	GetLastUsedRow = oEmptyRanges.RangeAddresses(Count-1).StartRow - 1
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

Sub subSaveAs(oDoc, sURL, Optional sType)
	If isMissing(sType) then
		oDoc.storeAsURL(sURL, array())
	Else
		Dim mFileType(0)
		mFileType(0) = createUnoStruct("com.sun.star.beans.PropertyValue")
		mFileType(0).Name = "FilterName"
		mFileType(0).Value = sType
'		oDoc.storeAsURL(sURL, mFileType())
		oDoc.storeToURL(sURL, mFileType())		
	End If
End Sub

Function OpenAsTemplate(sURL, Optional Hidden as Boolean) as Object
Dim a(1) As New com.sun.star.beans.PropertyValue
	a(0).Name = "AsTemplate"
	a(0).Value = true
	If Hidden = True Then
		a(1).Name = "Hidden"
		a(1).Value = True
	End If
	OpenAsTemplate = StarDesktop.LoadComponentFromUrl(sURL, "_blank" , 0, a())
End Function

Function ShowDocument(oDoc, Vis as Boolean)
	Controller = oDoc.CurrentController
	Frame = Controller.Frame
	ContainerWin = Frame.ContainerWindow
	ContainerWin.Visible = Vis
End Function

Function RemoveRowsByCellString(oTable as Object, iCol%, iRow%, sVal$)
Dim Count%
	Count = oTable.Rows.Count - 1
	For i = iRow To Count
		If oTable.getCellByPosition(iCol, i).String = sVal Then
			oTable.Rows.RemoveByIndex(i, 1)
			Count = Count - 1
			i = i - 1
		End If
	Next i
End Function

Function setOffsetValByCellString(oTable as Object, iCol%, iRow%, HorOffset%, sVal$, sVal2$)
Dim Count%
	Count = oTable.Rows.Count - 1
	For i = iRow To Count
		If oTable.getCellByPosition(iCol, i).String = sVal Then
			oTable.getCellByPosition(iCol + HorOffset, i).String = sVal2
		End If
	Next i
End Function


Function ConvertLength(Value, sFrom$, sTo$, Places)
Dim svc As Object
	svc = createUnoService("com.sun.star.sheet.FunctionAccess")
	Convert = svc.callFunction("CONVERT_ADD", Array(Value, sFrom, sTo))
	Rr = svc.callFunction("ROUND", Array(Convert, Places))
	ConvertLength = Rr
End Function

Function Min(Arr)
Dim svc As Object
Dim Fn
	svc = createUnoService("com.sun.star.sheet.FunctionAccess")
	Fn = svc.callFunction("MIN", Arr)
	Min = Fn
End Function

Function Max(Arr)
Dim svc As Object
Dim Fn
	svc = createUnoService("com.sun.star.sheet.FunctionAccess")
	Fn = svc.callFunction("MAX", Arr)
	Max = Fn
End Function

Function ConvertFromAnsi(ansi%)
	Select Case ansi:
		Case 150:
			ConvertFromAnsi = 20
		Case 300:
			ConvertFromAnsi = 50		
		Case 400:
			ConvertFromAnsi = 68
		Case 600:
			ConvertFromAnsi = 100
		Case 900:
			ConvertFromAnsi = 150
		Case 1500:
			ConvertFromAnsi = 250
		Case 2500:
			ConvertFromAnsi = 420
		Case Else: 
			ConvertFromAnsi = 0
	End Select
End Function

Function Replace_symbols(ByVal txt As String) As String
    St$ = "~!?@/\#$%^&*=|`"""
    For i% = 1 To Len(St$)
        txt = Replace(txt, Mid(St$, i, 1), "_")
    Next
    Replace_symbols = txt
End Function

Sub PrintFiles
	Dim AllFiles$, NextFile$
	Dim oBook, oDoc
	Dim a(0) As New com.sun.star.beans.PropertyValue
	Dim counter%

	GlobalScope.BasicLibraries.loadLibrary("Tools")

	a(0).Name = "Hidden"
	a(0).Value = True
	
	oBook = ThisComponent

	DestDir = DirectoryNameOutOfPath(oBook.getURL(),"/") & "/Расчеты/"
	NextFile = Dir(ConvertFromUrl(DestDir), 0)
	counter = 0

	Dim oProgressBar as Object, oProgressBarModel As Object, oDialog as Object
	Dim oLabel as Object
	Dim ProgressValue As Long, ProgressValueMin, ProgressValueMax, ProgressStep
	DialogLibraries.loadLibrary("Standard")
	oDialog = CreateUnoDialog(DialogLibraries.Standard.Dialog1)
	ProgressValueMin = 0
	ProgressValueMax = 1
	ProgressStep = 1
	oProgressBarModel = oDialog.getModel().getByName("ProgressBar1")
	oProgressBarModel.setPropertyValue("ProgressValueMin", ProgressValueMin)
	oProgressBarModel.setPropertyValue("ProgressValueMax", ProgressValueMax)
	oDialog.getModel().getByName("Label1").Label = "Печатаем расчеты. Подождите..."
	oLabel = oDialog.getModel().getByName("Label2")
	oDialog.setVisible(True)
	
	While NextFile <> "" 
		If InStr(1, NextFile, ".odt") > 0 and InStr(1, NextFile, "~") <= 0 Then	
			oLabel.Label = NextFile
			oProgressBarModel.setPropertyValue("ProgressValue", 1)
			oDoc = StarDesktop.LoadComponentFromUrl(DestDir & NextFile, "_parent" , 0, a()) 
			oDoc.Print(NoArgs())
		End If
		NextFile = Dir
		wait 3000	
	Wend
	
End Sub

