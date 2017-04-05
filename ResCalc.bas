'Option Explicit
GLOBAL LogReport as String

Sub Main(NumSheets)
	Dim oBook, oDoc1, oDoc2
	Dim CurDir$
	Dim counter%
	Dim MyFormat$

	GlobalScope.BasicLibraries.loadLibrary("Tools")

	oBook = ThisComponent
	LogReport = ""
	CurDir = DirectoryNameOutOfPath(oBook.getURL(),"/") 

	If CreateFolder(CurDir & "/Расчеты") = False Then
		Exit Sub
	End if

	MyFormat = "+0;-0;0"

	Dim oProgressBar as Object, oProgressBarModel As Object, oDialog as Object
	Dim oLabel as Object
	Dim ProgressValue As Long, ProgressValueMin, ProgressValueMax, ProgressStep
	DialogLibraries.loadLibrary("Standard")
	oDialog = CreateUnoDialog(DialogLibraries.Standard.Dialog1)
	REM progress bar settings
	ProgressValueMin = 0
	ProgressValueMax = Ubound(NumSheets()) - Lbound(NumSheets()) + 1
	ProgressStep = 1
	REM set minimum and maximum progress value
	oProgressBarModel = oDialog.getModel().getByName( "ProgressBar1" )
	oProgressBarModel.setPropertyValue( "ProgressValueMin", ProgressValueMin)
	oProgressBarModel.setPropertyValue( "ProgressValueMax", ProgressValueMax)
	oLabel = oDialog.getModel().getByName("Label2")
	REM show progress bar
	oDialog.setVisible( True )
	
	For counter = Lbound(NumSheets()) to Ubound(NumSheets())
		
		oLabel.Label = oBook.Sheets(counter).getName
	
		Dim Customer$, Obj$, Place$, sName$, sDn$, Dmm#, Dinch, Pn$
		Dim Manuf$, Material$, Env$, MinFlange#, MinBody#
		Dim TempRange, Temperature#, Pr#, Pansi
		Dim LastColumn%	
		Dim Header$, Fname1$, Fname2$
		Dim oTnormative, oTlist1, oTlist2, oTcalc
		Dim DataArr()
		Dim isHs as Boolean

		oDoc1 = Func.OpenAsTemplate(CurDir & "/rc1.odt", True)
		oDoc2 = Func.OpenAsTemplate(CurDir & "/rc2.odt", True)

		Fname1 = CurDir & "/Расчеты/РП_" & counter & ".odt"
		Fname2 = CurDir & "/Расчеты/РО_" & counter & ".odt"
		
		oTnormative = oDoc1.TextTables.getByName("normative")
		oTlist1 = oDoc1.TextTables.getByName("list1")
		oTlist2 = oDoc1.TextTables.getByName("list2")
		oTcalc = oDoc2.TextTables.getByName("calc_data")
	
		LastColumn  = Func.GetLastUsedColumn(oBook.Sheets.getByIndex(counter), 14)
		DataArr = oBook.Sheets(counter).getCellRangeByPosition(1, 14, LastColumn, 20).getDataArray()		

		Customer = oBook.Sheets(counter).getCellByPosition(1, 0).getString()
		Obj = oBook.Sheets(counter).getCellByPosition(1, 1).getString()
		Place = oBook.Sheets(counter).getCellByPosition(1, 2).getString() & Chr(13)
		sName = oBook.Sheets(counter).getCellByPosition(1, 4).getString()
		sDn = oBook.Sheets(counter).getCellByPosition(1, 7).getString()
		Pn = oBook.Sheets(counter).getCellByPosition(1, 8).getString()
		Manuf = oBook.Sheets(counter).getCellByPosition(1, 5).getString()
		Material = oBook.Sheets(counter).getCellByPosition(1, 6).getString()
		Env = oBook.Sheets(counter).getCellByPosition(1, 9).getString()
		MinFlange = Func.Min(DataArr(5))
		MinBody = Func.Min(DataArr(6))
		TempRange = oBook.Sheets(counter).getCellRangeByPosition(1, 10, 2, 10).getData()
		Temperature = Func.Max(TempRange(0))
		
		If oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsAvg").State = 1 _
		Or Temperature < 20 Then
			Temperature = 20
		End If
		
		If oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsInch").State = 1 Then
			Dinch = sDn
			Dmm = ConvertLength(Dinch, "in", "mm", -1)
			sDn = sDn & Chr(34)
		Else
			Dmm = sDn
			Dinch = ConvertLength(Dmm, "mm", "in", 0)
			sDn = "Ду" & sDn			
		End If
	
		If  oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsAnsi").State = 1 Then
			Pansi = Pn
			Pr = Func.ConvertFromAnsi(Pansi)
			Pn = Pn & "#"
			setOffsetValByCellString(oTlist2, 0, 0, 2, "C", Pansi & " (" & Pr & ")")
		Else
			Pr = Pn
			Pansi = 0
			Pn = "Ру" & Pn
			RemoveRowsByCellString(oTnormative, 0, 0, "B")
			RemoveRowsByCellString(oTlist2, 0, 0, "C")
		End If	

		If oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsHs").State = 1 Then
			oTcalc.getCellByName("E3").String = Chr(8805) & "5"	
			isHs = True
		Else
			RemoveRowsByCellString(oTnormative, 0, 0, "A")
			oTcalc.getCellByName("E3").String = Chr(8805) & "8"
			isHs = False
		End If		

		If Trim(Manuf) <> "" Then
			Manuf = "фирмы-изготовителя: " & Manuf
		End If
		
		If Material = "" Then
			RemoveRowsByCellString(oTlist2, 0, 0, "B")
		Else
			setOffsetValByCellString(oTlist2, 0, 0, 2, "B", Material)
		End If

		
		Header = Customer & ", " & Obj & ", " & Place &_
				 sName & " (" & sDn & ", " & Pn & ") " & Manuf

		oDoc1.TextFrames.getByName("header").String = Header
		oDoc2.TextFrames.getByName("header").String = Header
		oDoc1.TextFrames.getByName("ngroup").String = counter
		oDoc2.TextFrames.getByName("ngroup").String = counter

		'Заполнение таблицы 1 - обозначение...
		oTlist1.Rows.insertByIndex(oTlist1.Rows.Count, LastColumn - 1)
		
		For i = 0 To UBound(DataArr(1))
			oTlist1.getCellByPosition(0, i + 1).String = DataArr(1)(i)
			oTlist1.getCellByPosition(1, i + 1).String = DataArr(0)(i)
		Next i
		
		setOffsetValByCellString(oTlist2, 0, 0, 2, "A", Dmm & " (" & Dinch & Chr(34) & ")")
		setOffsetValByCellString(oTlist2, 0, 0, 2, "D", Format(Temperature, MyFormat))
		setOffsetValByCellString(oTlist2, 0, 0, 2, "E", Pr)
		setOffsetValByCellString(oTlist2, 0, 0, 2, "F", Env)
		

		oTcalc.getCellByPosition(0, 2).String = Join(DataArr(1), ", ")
		oTcalc.getCellByPosition(2, 2).String = Dmm & " (" & Dinch & Chr(34) & ")"
		oTcalc.getCellByPosition(3, 2).String = Format(MinFlange, "0.0")
		oTcalc.getCellByPosition(3, 3).String = Format(MinBody, "0.0")
		
		LogReport = LogReport & oBook.Sheets(counter).Name & ": " & _
					Calculation(Dmm, Pr, Temperature, MinFlange, MinBody, isHs) & Chr(13)		

		If FileExists(Fname1) Then
			Kill(Fname1)
		End If
		
		If FileExists(Fname2) Then
			Kill(Fname2)
		End If

		Func.subSaveAs(oDoc1, Fname1)
		Func.subSaveAs(oDoc2, Fname2)
		oDoc1.close(False)
		oDoc2.close(False)		

		oProgressBarModel.setPropertyValue( "ProgressValue", Counter)

	Next counter
		
	FileNo = FreeFile
	Open CurDir & "/LogReport.txt" For Output As #FileNo
	Print #FileNo, LogReport
	Close #FileNo	
		
	Shell "notepad " & ConvertFromUrl(CurDir & "/LogReport.txt")
	
End Sub

Function Calculation(Dmm, Pr, Temp, MinFlange, MinBody, IsHs as Boolean)
	Dim Sr1, Sr2
	Dim Sigma
	
	If IsHs = True Then 
        Select Case Temp
            Case 20 To 49.9: Sigma = 960
            Case 50 To 74.9: Sigma = 950
            Case 75 To 99.9: Sigma = 940
            Case 100 To 124.9: Sigma = 920
            Case 125 To 149.9: Sigma = 910
            Case 150 To 174.9: Sigma = 890
            Case 175 To 199.9: Sigma = 880
            Case 200 To 224.9: Sigma = 860
            Case 225 To 249.9: Sigma = 810
            Case 250 To 274.9: Sigma = 790
            Case 275 To 299.9: Sigma = 750
            Case 300 To 349.9: Sigma = 660
            Case Is > 350: Sigma = 610
            Case Else: Sigma = 970
        End Select
	Else
        Select Case Temp
            Case 20 To 49.9: Sigma = 1280
            Case 50 To 74.9: Sigma = 1260
            Case 75 To 99.9: Sigma = 1250
            Case 100 To 124.9: Sigma = 1230
            Case 125 To 149.9: Sigma = 1220
            Case 150 To 174.9: Sigma = 1200
            Case 175 To 199.9: Sigma = 1180
            Case 200 To 224.9: Sigma = 1150
            Case 225 To 249.9: Sigma = 1120
            Case 250 To 274.9: Sigma = 1060
            Case 275 To 299.9: Sigma = 1000
            Case 300 To 349.9: Sigma = 880
            Case 350 To 374.9: Sigma = 820
            Case 375 To 399.9: Sigma = 770
            Case 400 To 409.9: Sigma = 750
            Case 410 To 419.9: Sigma = 720
            Case 420 To 429.9: Sigma = 680
            Case 430 To 439.9: Sigma = 600
            Case Is > 440: Sigma = 530
            Case Else: Sigma = 1300
        End Select		
	End If
	
    Sr1 = Round((Dmm * Pr) / (2 * Sigma - Pr), 2)
    Sr2 = Round((1.5 * Dmm * Pr) / (2 * Sigma - Pr), 2)

    If (Sr1 < MinFlange) And (Sr2 < MinBody) Then
		Calculation = "Обеспечено"
    Else
		Calculation = "Не проходит: Smin Фланца = " & Sr1 & " Smin Корпуса = " & Sr2
    End If
	
End Function


Sub CalcAllSheets
Dim oBook
	oBook = ThisComponent
	Dim a(1 To oBook.Sheets.Count - 1)	
	Main(a())
End Sub

Sub CalcOneSheet
	oBook = ThisComponent
	Current = oBook.CurrentController.ActiveSheet.getRangeAddress().sheet
	Dim a(Current To Current)
	Main(a())
End Sub

