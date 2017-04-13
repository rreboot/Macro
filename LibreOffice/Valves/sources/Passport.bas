'Option Explicit

Sub Main
	Dim oBook, oDoc
	Dim CurDir$
	Dim counter%
	Dim MyFormat$

	GlobalScope.BasicLibraries.loadLibrary("Tools")

	oBook = ThisComponent
	CurDir = DirectoryNameOutOfPath(oBook.getURL(),"/") 

	If CreateFolder(CurDir & "/Паспорта") = False Then
		Exit Sub
	End if

	Dim oProgressBar as Object, oProgressBarModel As Object, oDialog as Object
	Dim oLabel as Object
	DialogLibraries.loadLibrary("Standard")
	oDialog = CreateUnoDialog(DialogLibraries.Standard.Dialog1)
	oProgressBarModel = oDialog.getModel().getByName( "ProgressBar1" )
	oProgressBarModel.EnableVisible(False)
	oLabel = oDialog.getModel().getByName("Label2")
	oDialog.getModel().getByName("Label1").Label = "Формируем паспорта. Подождите..."
	oDialog.setVisible( True )
	
	For counter = 1 to oBook.Sheets.Count - 1
		
'		oLabel.Label = oBook.Sheets(counter).getName
	
		Dim Customer$, sName$, Dn$, Dmm#, Pn$, Pw#, Pt#, Env$, sY$, sPurp$
		Dim Manuf$, OKP$, sYear$, Nmanuf$, Obj$, Place$, PasspNum$, PrepDate$
		Dim DensityTest#, EnduranceTest#
		Dim LastColumn%, LastRow%
		Dim oTable, Fname$
		Dim DataArr(), MaterialArr(), Purpose(2)

		LastColumn  = Func.GetLastUsedColumn(oBook.Sheets.getByIndex(counter), 14)
		LastRow = Func.GetLastUsedRow(oBook.Sheets.getByindex(counter), 0)

		DataArr = oBook.Sheets(counter).getCellRangeByPosition(1, 14, LastColumn, 20).getDataArray()		
		MaterialArr = oBook.Sheets(counter).getCellRangeByPosition(0, 22, 1, LastRow).getDataArray()


		Customer = oBook.Sheets(counter).getCellByPosition(1, 0).getString()
		Obj = oBook.Sheets(counter).getCellByPosition(1, 1).getString()
		Place = oBook.Sheets(counter).getCellByPosition(1, 2).getString()
		sName = oBook.Sheets(counter).getCellByPosition(1, 4).getString()
		Dn = oBook.Sheets(counter).getCellByPosition(1, 7).getString()
		Pn = oBook.Sheets(counter).getCellByPosition(1, 8).getString()
		Manuf = oBook.Sheets(counter).getCellByPosition(1, 5).getString()
		Pw = CDbl(oBook.Sheets(counter).getCellByPosition(1, 11).getString())
		sY = Right(DataArr(0)(0), 2)
		PasspNum = Place & "-" & sY & "-"
		PrepDate = oBook.Sheets(0).getCellByPosition(3, 2).getString()
		
		Purpose(0) = "Запорная арматура для перекрытия потока рабочей среды с определенной герметичностью"
		Purpose(1) = "Регулирующая арматура для регулирования параметров рабочей среды посредством изменения расхода"
		Purpose(2) = "Обратная арматура для автоматического предотвращения обратного потока рабочей среды"
		
		If oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsInch").State = 1 Then
			Dn = Dn & Chr(34)
		End If
	
		If oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsAnsi").State = 1 Then
			Pt = Func.ConvertFromAnsi(Pn)
			Pn = Pn & "#"
		Else
			Pt = Pn
		End If	

		If oBook.Sheets(counter).DrawPage.Forms("Standard").getByName("IsHs").State = 1 Then
			Env = "Природный газ с содержанием H2S до 6% объема, жидкие углеводороды"
		Else
			Env = "Жидкие и газообразные неагрессивные среды"
		End If		

		If Trim(Manuf) <> "" Then
			Manuf = "-"
		End If
		
		Select Case True
			Case (InStr(1, LCase(sName), "задвижка") > 0)
				OKP = "374120"
				sPurp = Purpose(0)
			Case (InStr(1, LCase(sName), "обратн") > 0)
				OKP = "374230"
				sPurp = Purpose(2)
			Case (InStr(1, LCase(sName), "регулирующ") > 0)
				OKP = "374250"
				sPurp = Purpose(1)
			Case (InStr(1, LCase(sName), "кран") > 0)
				OKP = "374220"
				sPurp = Purpose(0)
			Case (InStr(1, LCase(sName), "запорн") > 0)
				OKP = "374230"
				sPurp = Purpose(0)
			Case Else
				OKP = "374200"
				sPurp = Purpose(0)
		End Select

        If InStr(1, Obj, "ГПЗ") <> 0 Then
        	EnduranceTest = 1.25 * Pw
        	DensityTest = 1.1 * Pw
        Else
        	EnduranceTest = 1.25 * Pt
        	DensityTest = 1.1 * Pt
		End If		

	
		For i = 0 To UBound(DataArr(0))
		Dim iRows%
			Nmanuf = DataArr(1)(i)
			sYear = DataArr(2)(i)
			oLabel.Label = oBook.Sheets(counter).getName & " " & Nmanuf			
			Fname = CurDir & "/Паспорта/Паспорт Гр. " & counter & "_" & i & " " & Replace_symbols(Nmanuf) & ".odt"
			PasspNum = PasspNum & Nmanuf
			oDoc = Func.OpenAsTemplate(CurDir & "/passport.odt", True)		
			oTable = oDoc.TextTables(0)
		
			oDoc.Bookmarks.getByName("passp_number").Anchor.setString(PasspNum)
			oDoc.Bookmarks.getByName("prep_date").Anchor.setString(PrepDate)
			
			oTable.getCellByPosition(1, 1).setString(sName)
			oTable.getCellByPosition(1, 2).setString(OKP)
			oTable.getCellByPosition(1, 3).setString(sPurp)
			oTable.getCellByPosition(1, 4).setString(Manuf)
			oTable.getCellByPosition(1, 5).setString(Dn)
			oTable.getCellByPosition(3, 5).setString(sYear)			
			oTable.getCellByPosition(1, 6).setString(Pn)
			oTable.getCellByPosition(3, 6).setString(Nmanuf)
			oTable.getCellByPosition(3, 7).setString(Env)
			oTable.getCellByPosition(1, 9).setString(Format(EnduranceTest, "0.0#"))
			oTable.getCellByPosition(1, 10).setString(Format(DensityTest, "0.0#"))
			
			iRows = oTable.Rows.Count
			oTable.Rows.insertByIndex(iRows, UBound(MaterialArr()) + 1)
			
			oTable.getCellRangeByPosition(0, iRows, 1, oTable.Rows.Count - 1).SetDataArray(MaterialArr())
			
			If FileExists(Fname) Then
				Kill(Fname)
			End If

			Func.subSaveAs(oDoc, Fname)		
			oDoc.close(False)
			
		Next i
	Next counter
	
End Sub
