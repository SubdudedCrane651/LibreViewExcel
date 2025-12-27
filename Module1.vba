Sub CreateDiabetesChart()
	Dim ws As Worksheet
	Set ws = ThisWorkbook.Sheets("Glycèmie De Moi")

	' Delete existing charts
	Dim co As ChartObject
	For Each co In ws.ChartObjects
		co.Delete
	Next co

	' Create new chart
	Dim newChartObj As ChartObject
	Set newChartObj = ws.ChartObjects.Add( _
		Left : = ws.Range("K5").Left, _
		Top : = ws.Range("K5").Top, _
		Width : = 500, _
		Height : = 300)

	With newChartObj.Chart
		.ChartType = xlLine
		.DisplayBlanksAs = xlInterpolated

		Dim lastRow As Integer
		lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

		' Get unique dates while limiting to 20 days
		Dim uniqueDates As Object
		Set uniqueDates = CreateObject("Scripting.Dictionary")
		
		Dim i As Integer
		For i = 5 To lastRow
			If Not uniqueDates.exists(ws.Cells(i, 1).Value) Then
				uniqueDates.Add ws.Cells(i, 1).Value, ws.Cells(i, 1).Value
			End If
			If uniqueDates.Count >= 20 Then Exit For
		Next i

		' Define X-axis range (filtered dates)
		Dim xValuesRange As String
		xValuesRange = "A5:A" & lastRow
		
		' Ensure each series is plotted
		Dim seriesCount As Integer
		seriesCount = 0

		' SERIES 1: Before Breakfast
		If Application.WorksheetFunction.Count(ws.Range("B5:B" & lastRow)) > 0 Then
			seriesCount = seriesCount + 1
			With .SeriesCollection.NewSeries
				.Name = "Glycémie à jeun"
				.XValues = ws.Range(xValuesRange)
				.Values = ws.Range("B5:B" & lastRow)
			End With
		End If

		' SERIES 2: Before Dinner
		If Application.WorksheetFunction.Count(ws.Range("D5:D" & lastRow)) > 0 Then
			seriesCount = seriesCount + 1
			With .SeriesCollection.NewSeries
				.Name = "Glycémie avant diner"
				.XValues = ws.Range(xValuesRange)
				.Values = ws.Range("C5:C" & lastRow)
			End With
		End If

		' SERIES 3: Before Supper
		If Application.WorksheetFunction.Count(ws.Range("E5:E" & lastRow)) > 0 Then
			seriesCount = seriesCount + 1
			With .SeriesCollection.NewSeries
				.Name = "Glycémie avant souper"
				.XValues = ws.Range(xValuesRange)
				.Values = ws.Range("D5:D" & lastRow)
			End With
		End If

		' SERIES 4: Before Sleeping
		If Application.WorksheetFunction.Count(ws.Range("F5:F" & lastRow)) > 0 Then
			seriesCount = seriesCount + 1
			With .SeriesCollection.NewSeries
				.Name = "Glycémie la nuit"
				.XValues = ws.Range(xValuesRange)
				.Values = ws.Range("E5:E" & lastRow)
			End With
		End If

		' Format axes
		With .Axes(xlCategory)
			.TickLabels.Orientation = 45
			.HasTitle = True
			.AxisTitle.Text = "Date"
		End With

		With .Axes(xlValue)
			.HasTitle = True
			.AxisTitle.Text = "Glucose"
		End With

		.HasLegend = True

		' Assign colors to each series
		If seriesCount >= 1 Then.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(255, 0, 0) ' Red
		If seriesCount >= 2 Then.SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(0, 255, 0) ' Green
		If seriesCount >= 3.9 Then.SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(0, 0, 255) ' Blue
		If seriesCount >= 4 Then.SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(255, 165, 0) ' Orange
	End With
End Sub

Sub GlucoseColorIndex()
	Dim ws As Worksheet
	Set ws = ThisWorkbook.Sheets("Glycèmie De Moi") ' Ensure it runs on Sheet2

	' High value (Red)
	For Each cell In ws.Range("B5:B1000")
		If cell.Value > 10 Then cell.Font.Color = vbRed
	Next cell
	
	For Each cell In ws.Range("C5:C1000")
		If cell.Value > 10 Then cell.Font.Color = vbRed
	Next cell
	
	For Each cell In ws.Range("D5:D1000")
		If cell.Value > 10 Then cell.Font.Color = vbRed
	Next cell
	
	For Each cell In ws.Range("E5:E1000")
		If cell.Value > 10 Then cell.Font.Color = vbRed
	Next cell
	
	For Each cell In ws.Range("F5:F1000")
		If cell.Value > 10 Then cell.Font.Color = vbRed
	Next cell
	
	For Each cell In ws.Range("a2:F2")
		If cell.Value > 10 Then cell.Font.Color = vbRed
	Next cell

	' Normal value (Green)
	For Each cell In ws.Range("B5:F1000")
		If cell.Value <= 10 And cell.Value >= 3.9 Then cell.Font.Color = RGB(0, 128, 0) ' Green
	Next cell
	
	' Normal value (Green)
	For Each cell In ws.Range("a2:F2")
		If cell.Value <= 10 And cell.Value >= 3.9 Then cell.Font.Color = RGB(0, 128, 0) ' Green
	Next cell

	' Low value (Blue)
	For Each cell In ws.Range("B5:F1000")
		If cell.Value < 3.9 Then cell.Font.Color = vbBlue
	Next cell
	
	For Each cell In ws.Range("a2:f2")
		If cell.Value < 3.9 Then cell.Font.Color = vbBlue
	Next cell


	'MsgBox "Glucose values colored successfully!", vbInformation, "Success"
End Sub

Sub CalculateDailyAverages()
	Dim ws As Worksheet
	Dim lastRow As Long
	Dim dict As Object
	Dim cell As Range
	Dim DateValue As Variant
	Dim avgValue As Double
	Dim sumValues As Double
	Dim countValues As Integer
	Dim rng As Range
	Dim outputRow As Long

	' Set worksheet
	Set ws = ThisWorkbook.Sheets("Glycèmie De Moi") ' Change "Sheet1" to your actual sheet name

	' Find last row in column B
	lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

	' Create a dictionary to store sum and count for each date
	Set dict = CreateObject("Scripting.Dictionary")

	' Loop through column B starting from row 5
	For Each cell In ws.Range("A5:A" & lastRow)
		DateValue = cell.Value
		If Not dict.exists(DateValue) Then
			dict.Add DateValue, Array(0, 0) ' Store sum and count as an array
		End If

		' Get current sum and count
		sumValues = dict(DateValue)(0)
		countValues = dict(DateValue)(1)

		' Calculate sum for columns B to E
		Set rng = ws.Range(cell.Offset(0, 1), cell.Offset(0, 4)) ' Columns B to E
		sumValues = sumValues + Application.WorksheetFunction.Sum(rng)
		countValues = countValues + Application.WorksheetFunction.Count(rng)

		' Update dictionary
		dict(DateValue) = Array(sumValues, countValues)
	Next cell

	' Output averages in column H starting from row 5
	outputRow = 5
	For Each DateValue In dict.keys
		avgValue = dict(DateValue)(0) / dict(DateValue)(1) ' Calculate average
		'ws.Cells(outputRow, 8).Value = dateValue ' Place date in column H
		ws.Cells(outputRow, 6).Value = Round(avgValue, 1) ' Place average in column H
		outputRow = outputRow + 1
	Next DateValue

	' Cleanup
	Set dict = Nothing
	
	Call GlucoseColorIndex
	
	'MsgBox "Averages calculated successfully!", vbInformation
End Sub

Sub CopyDataByDateRange(msg)
	Dim wsSource As Worksheet, wsDest As Worksheet
	Dim lastRow As Long, destRow As Long
	Dim startDate As Date, endDate As Date
	Dim cell As Range

	' Set worksheets
	Set wsSource = ThisWorkbook.Sheets("Page D'accueil") ' Source data
	Set wsDest = ThisWorkbook.Sheets("Glycèmie De Moi") ' Destination data

	' Clear previous data on Sheet2 (A5:F1000)
	wsDest.Range("A5:F1000").ClearContents

	' Ensure date values are properly extracted
	startDate = CDate(wsSource.Range("I2").Value)
	endDate = CDate(wsSource.Range("J2").Value)

	' Find the last row in column A of Page D'accueil
	lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
	destRow = 5 ' Start copying to Sheet2 at row 5

	' Convert Column A values to actual dates
	For Each cell In wsSource.Range("B2:B" & lastRow)
		cell.Value = CDate(cell.Value)
	Next cell

	' Loop through rows in Column A and copy data based on date range
	For Each cell In wsSource.Range("B2:B" & lastRow)
		Debug.Print "Checking row:", cell.Row, " Date:", cell.Value, " Start:", startDate, " End:", endDate ' Debugging

		If IsDate(cell.Value) Then
			If CDate(cell.Value) >= startDate And CDate(cell.Value) <= endDate Then
				wsSource.Range(cell.Offset(0, 0), cell.Offset(0, 5)).Copy ' Copy B to F
				wsDest.Cells(destRow, 1).PasteSpecial Paste : = xlPasteValues
				destRow = destRow + 1
			End If
		End If
	Next cell

	' Cleanup
	Application.CutCopyMode = False
	Call CalculateDailyAverages
	Call GlucoseColorIndex
	If msg = True Then
		MsgBox "Data copied successfully!", vbInformation, "Success"
	End If
End Sub

Function GetWindowsVersion() As String
	Dim objOS As Object
	Set objOS = CreateObject("WScript.Shell")
	GetWindowsVersion = objOS.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
End Function


Sub RemoveImageAndRunPython()
	Dim ws As Worksheet
	Dim shp As Shape
	Dim pythonScript As String
	
	' Set worksheet
	Set ws = ThisWorkbook.Sheets("Glycèmie De Moi") ' Change sheet name if needed
	
	' Loop through all shapes to find and delete the image
	For Each shp In ws.Shapes
		If shp.TopLeftCell.Address = "$K$27" Then
			shp.Delete
		End If
	Next shp
	
	' Set the Python script path
	pythonScript = """f:/Python/LibreViewExcel/Glucose_Chart.py""" ' Update with your actual Python script
	
	' Run the Python script
	If GetWindowsVersion() = "Windows 10 Home" Then

		Shell "F:\Python\LibreViewExcel\.venv\Scripts\python.exe " & pythonScript, vbNormalFocus


	Else


		'Shell "C:/Users/rchrd/AppData/Local/Programs/Python/Python312/python.exe " & pythonScript, vbNormalFocus
		Shell "F:\Python\Blood_Pressure_Tracker\BOSGAME\.venv\Scripts\python.exe " & pythonScript, vbNormalFocus


	End If
	
	MsgBox "Image deleted & Python script executed!", vbInformation, "Success"
End Sub
Sub RunExeInWorkbookDirectory()
	Dim exePath As String
	exePath = ThisWorkbook.Path & "\glucose_chart.exe"
	
	If Dir(exePath) <> "" Then
		Shell exePath, vbNormalFocus
	Else
		MsgBox "Executable not found in workbook directory.", vbExclamation
	End If
End Sub

Sub CopyChartAndPictureToZ1AndPrint()
	
	Dim ws As Worksheet
	Dim targetCell As Range
	Dim targetLeft As Double, targetTop As Double
	Dim shp As Shape
	Dim chartObj As ChartObject
	Dim imgShape As Shape
	Dim pastedChart As ChartObject
	Dim pastedImage As Shape
	Dim i As Long
	
	Set ws = ActiveSheet
	Set targetCell = ws.Range("Z1")
	
	targetLeft = targetCell.Left
	targetTop = targetCell.Top
	
	'===========================================================
	' 1. DELETE ANY SHAPES ALREADY IN THE Z1 AREA
	'===========================================================
	Range("Y1").Select
	
	For Each shp In ws.Shapes
		If shp.Left >= targetLeft And shp.Left <= targetLeft + 500 _
				 And shp.Top >= targetTop And shp.Top <= targetTop + 800 Then
			shp.Delete
		End If
	Next shp
	
	'===========================================================
	' 2. FIND THE ONLY CHART ON THE SHEET
	'===========================================================
	If ws.ChartObjects.Count = 0 Then
		MsgBox "No chart found on sheet."
		Exit Sub
	End If
	
	Set chartObj = ws.ChartObjects(1) ' assume only one chart
	
	'===========================================================
	' 3. FIND THE LAST NON-CHART SHAPE (ASSUMED PICTURE)
	'===========================================================
	Set imgShape = Nothing
	
	For i = ws.Shapes.Count To 1 Step - 1
		Set shp = ws.Shapes(i)
		
		' skip shapes that are charts
		If shp.Type <> msoChart Then
			' also skip if this shape is the chart's container shape (name check)
			If shp.Name <> chartObj.Name Then
				Set imgShape = shp
				Exit For ' last non-chart found (loop is backwards)
			End If
		End If
	Next i
	
	If imgShape Is Nothing Then
		MsgBox "No suitable non-chart shape (picture) found on sheet."
		Exit Sub
	End If
	
	'===========================================================
	' 4. COPY CHART TO Z1
	'===========================================================
	
	' Make sure NOTHING is selected so Excel does a real copy
	Application.CutCopyMode = False
	ws.Cells(1, 1).Select
	
	chartObj.Copy
	
	On Error Resume Next
	ws.Paste
	On Error GoTo 0
	
	Set pastedChart = ws.ChartObjects(ws.ChartObjects.Count)
	
	With pastedChart
		.Left = targetLeft
		.Top = targetTop
		.Width = 400
		.Height = 300
	End With
	
	'===========================================================
	' 5. DUPLICATE "PICTURE" BELOW THE CHART (NO CLIPBOARD)
	'===========================================================
	' No Copy/Paste here – Duplicate avoids the ghost
	Set pastedImage = imgShape.Duplicate
	
	With pastedImage
		On Error Resume Next
		.LockAspectRatio = msoTrue
		.Width = 400
		On Error GoTo 0
		.Left = targetLeft
		.Top = pastedChart.Top + pastedChart.Height + 20
	End With
	
	'===========================================================
	' 6. PRINT SETUP
	'===========================================================
	Call SetupPrintAreaAndHeadersForChartAndPicture

	Call CreateDiabetesChart
	
End Sub
Sub ListChartAndImageNames()
	Dim sh As Shape
	Dim co As ChartObject
	Dim ws As Worksheet
	Dim i As Integer
	
	Set ws = ActiveSheet
	i = 1
	
	' Create headers
	ws.Cells(1, 1).Value = "Shape Type"
	ws.Cells(1, 2).Value = "Name"
	
	' Loop through all shapes
	For Each sh In ws.Shapes
		ws.Cells(i + 1, 1).Value = sh.Type
		ws.Cells(i + 1, 2).Value = sh.Name
		i = i + 1
	Next sh
	
	' Loop through all embedded charts
	For Each co In ws.ChartObjects
		ws.Cells(i + 1, 1).Value = "ChartObject"
		ws.Cells(i + 1, 2).Value = co.Name
		i = i + 1
	Next co
	
	MsgBox "Done! Check column A and B for names."
End Sub

Sub SetupPrintAreaAndHeaders()
	Dim ws As Worksheet
	Set ws = ActiveSheet
	Dim currentDateTime As Date
	
	With ws.PageSetup
		' Set print area
		.PrintArea = "A1:F30"
		
		' Set orientation
		.Orientation = xlLandscape

		currentDateTime = Now
		
		CenterHeader = ""
		.RightHeader = ""
		.RightFooter = ""
		
		' Set headers
		.CenterHeader = "Les dernières lectures de glycémie de Richard Perreault pour 14 jours"
		.RightFooter = "Date du rapport " + Format(currentDateTime, "dd/mm/yyyy hh:mm:ss")
		
		' Optional: Fit to one page
		.Zoom = False
		.FitToPagesWide = 1
		.FitToPagesTall = 1
		
		' Optional: Margins and alignment
		.CenterHorizontally = True
		.CenterVertically = False
	End With
	
	' Preview before printing
	ws.PrintPreview
End Sub

Sub SetupPrintAreaAndHeadersForChartAndPicture()
	Dim ws As Worksheet
	Set ws = ActiveSheet
	Dim currentDateTime As Date
	
	With ws.PageSetup
		' Set print area
		'.PrintArea = "Z:AH"
		Call SetPrintArea
		
		' Set orientation
		.Orientation = xlPortrait

		currentDateTime = Now

		.CenterHeader = ""
		.RightHeader = ""
		.RightFooter = ""
		
		' Set headers
		.CenterHeader = "Les dernières lectures de glycémie de Richard Perreault pour 14 jours"
		.RightFooter = "Date du rapport " + Format(currentDateTime, "dd/mm/yyyy hh:mm:ss")
		
		' Optional: Fit to one page
		.Zoom = False
		.FitToPagesWide = 1
		.FitToPagesTall = 1
		
		' Optional: Margins and alignment
		.CenterHorizontally = True
		.CenterVertically = False
	End With
	
	' Preview before printing
	ws.PrintPreview
End Sub