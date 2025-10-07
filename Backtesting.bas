Option Explicit



Private Function GetValidationItems(ByVal c As Range) As Variant
	' Returns a 1-D variant array of list items from a data validation list
	Dim f As String, src As String
	Dim rng As Range, arr As Variant
	Dim i As Long, tmp As Variant
	
	If c.Validation.Type <> xlValidateList Then
		Exit Function
	End If
	
	f = c.Validation.Formula1
	If Len(f) = 0 Then Exit Function
	
	If Left$(f, 1) = "=" Then
		src = Mid$(f, 2)
		On Error Resume Next
		Set rng = Nothing
		' Try as a direct range
		Set rng = c.Parent.Range(src)
		' Try as a workbook-level name
		If rng Is Nothing Then
			If ThisWorkbook.Names.Count > 0 Then
				Dim nm As Name
				For Each nm In ThisWorkbook.Names
					If LCase$(nm.Name) = LCase$(src) Then
						If Not nm.RefersToRange Is Nothing Then
							Set rng = nm.RefersToRange
						End If
						Exit For
					End If
				Next nm
			End If
		End If
		On Error GoTo 0
		
		If Not rng Is Nothing Then
			arr = rng.Value
			' Normalize to 1-D array
			If IsArray(arr) Then
				ReDim tmp(1 To rng.Cells.Count)
				For i = 1 To rng.Cells.Count
					tmp(i) = rng.Cells(i).Value
				Next i
				GetValidationItems = tmp
			Else
				ReDim tmp(1 To 1): tmp(1) = arr
				GetValidationItems = tmp
			End If
		Else
			' Fallback: try evaluating to array (for dynamic arrays)
			arr = Application.Evaluate(f)
			If IsArray(arr) Then
				ReDim tmp(1 To UBound(arr, 1) * UBound(arr, 2))
				Dim r As Long, c2 As Long, k As Long: k = 1
				For r = LBound(arr, 1) To UBound(arr, 1)
					For c2 = LBound(arr, 2) To UBound(arr, 2)
						tmp(k) = arr(r, c2): k = k + 1
					Next c2
				Next r
				GetValidationItems = tmp
			Else
				' Comma-separated literal list without leading '='
			End If
		End If
	Else
		' Comma-separated list
		GetValidationItems = Split(f, ",")
	End If
End Function

Public Sub CheckBM22xBM23Combinations()
	' Run on the active sheet by default
	Dim wsName As String: wsName = ActiveSheet.Name
	RunBM22xBM23OnSheet ThisWorkbook.Worksheets(wsName)
End Sub

Public Sub CheckBM22xBM23_AllSheetsExceptMaster()
	Dim ws As Worksheet
	Dim totalSheets As Long, processedSheets As Long
	Dim startTime As Double, elapsed As Double, avgPerSheet As Double
	Dim remaining As Long, estSec As Long

	' Count eligible sheets
	For Each ws In ThisWorkbook.Worksheets
		If LCase(ws.Name) <> "master" Then totalSheets = totalSheets + 1
	Next ws
	processedSheets = 0
	startTime = Timer

	For Each ws In ThisWorkbook.Worksheets
		If LCase(ws.Name) <> "master" Then
			Application.StatusBar = "Starting sheet " & (processedSheets + 1) & "/" & totalSheets & ": " & ws.Name
			RunBM22xBM23OnSheet ws
			processedSheets = processedSheets + 1
			elapsed = Timer - startTime
			If processedSheets > 0 Then
				avgPerSheet = elapsed / processedSheets
				remaining = Application.Max(0, totalSheets - processedSheets)
				estSec = CLng(avgPerSheet * remaining)
				Application.StatusBar = "Completed " & processedSheets & "/" & totalSheets & _
					" sheets. Overall ETA: " & FormatETA(estSec)
			End If
		End If
	Next ws

	Application.StatusBar = False
End Sub

Private Sub RunBM22xBM23OnSheet(ByVal ws As Worksheet)
	Dim cBM22 As Range, cBM23 As Range, rngCheck As Range, resultsStart As Range
	Dim items22 As Variant, items23 As Variant
	Dim item22 As Variant, item23 As Variant
	Dim negCount As Long, firstNegCell As Range, cell As Range
	Dim nextRow As Long
	Dim wasScreenUpdating As Boolean, wasEnableEvents As Boolean, oldCalc As XlCalculation
	Dim totalCombos As Long
	Dim count22 As Long, count23 As Long
	Dim startSheetTime As Double, processedCombos As Long

	On Error GoTo CleanFail

	Set cBM22 = ws.Range("BM22")
	Set cBM23 = ws.Range("BM23")
	Set rngCheck = ws.Range("BN22:BN54")
	Set resultsStart = ws.Range("CI22")

	items22 = GetValidationItems(cBM22)
	items23 = GetValidationItems(cBM23)
	If IsEmpty(items22) Then Exit Sub
	If IsEmpty(items23) Then Exit Sub

	ws.Range(resultsStart, resultsStart.Offset(1000, 5)).Clear
	resultsStart.Resize(1, 5).Value = Array("BM22_Value", "BM23_Value", "HasNegative", "FirstNegativeCell", "FirstNegativeValue")
	nextRow = resultsStart.Row + 1

	wasScreenUpdating = Application.ScreenUpdating
	wasEnableEvents = Application.EnableEvents
	oldCalc = Application.Calculation
	Application.ScreenUpdating = False
	Application.EnableEvents = False
	Application.Calculation = xlCalculationManual

	On Error Resume Next
	count22 = UBound(items22) - LBound(items22) + 1
	count23 = UBound(items23) - LBound(items23) + 1
	On Error GoTo 0
	If count22 < 1 Or count23 < 1 Then GoTo CleanOk
	totalCombos = count22 * count23
	startSheetTime = Timer
	processedCombos = 0

	Dim i As Long, j As Long
	For i = LBound(items22) To UBound(items22)
		item22 = items22(i)
		If Len(CStr(item22)) > 0 Then
			cBM22.Value = item22
			For j = LBound(items23) To UBound(items23)
				item23 = items23(j)
				If Len(CStr(item23)) > 0 Then
					cBM23.Value = item23
					ws.Calculate
					negCount = Application.WorksheetFunction.CountIf(rngCheck, "<0")
					Set firstNegCell = Nothing
					If negCount > 0 Then
						For Each cell In rngCheck.Cells
							If Val(cell.Value) < 0 Then
								Set firstNegCell = cell
								Exit For
							End If
						Next cell
					End If
					ws.Cells(nextRow, resultsStart.Column).Value = item22
					ws.Cells(nextRow, resultsStart.Column + 1).Value = item23
					ws.Cells(nextRow, resultsStart.Column + 2).Value = (negCount > 0)
					If Not firstNegCell Is Nothing Then
						ws.Cells(nextRow, resultsStart.Column + 3).Value = firstNegCell.Address(0, 0)
					ws.Cells(nextRow, resultsStart.Column + 4).Value = firstNegCell.Value
					Else
						ws.Cells(nextRow, resultsStart.Column + 3).Value = ""
						ws.Cells(nextRow, resultsStart.Column + 4).Value = ""
					End If
					nextRow = nextRow + 1
					processedCombos = processedCombos + 1
					' Throttle status updates to avoid UI overhead
					If (processedCombos Mod 25) = 0 Or processedCombos = totalCombos Then
						Dim elapsedSheet As Double, rate As Double, remCombos As Long, estSheetSec As Long
						elapsedSheet = Timer - startSheetTime
						If elapsedSheet > 0 Then
							rate = elapsedSheet / processedCombos
							remCombos = Application.Max(0, totalCombos - processedCombos)
							estSheetSec = CLng(rate * remCombos)
							Application.StatusBar = "Sheet " & ws.Name & ": " & processedCombos & "/" & totalCombos & _
								" (" & Format(processedCombos / totalCombos, "0%") & ") ETA: " & FormatETA(estSheetSec)
						End If
					End If
				End If
			Next j
		End If
	Next i

	StyleResults ws, resultsStart, 5

CleanOk:
	Application.ScreenUpdating = wasScreenUpdating
	Application.EnableEvents = wasEnableEvents
	Application.Calculation = oldCalc
	Exit Sub

CleanFail:
	Application.ScreenUpdating = wasScreenUpdating
	Application.EnableEvents = wasEnableEvents
	Application.Calculation = oldCalc
	Exit Sub
End Sub

Private Function FormatETA(ByVal seconds As Long) As String
	If seconds < 60 Then
		FormatETA = seconds & "s"
	Else
		FormatETA = Int(seconds / 60) & "m " & (seconds Mod 60) & "s"
	End If
End Function

Public Sub ConsolidateResultsToMaster()
	Dim master As Worksheet, ws As Worksheet
	Dim outStart As Range
	Dim nextRow As Long, srcLastRow As Long
	Dim srcStart As Range
	Dim startTime As Double, processed As Long, totalSheets As Long
	Dim elapsed As Double, avg As Double, remaining As Long, estSec As Long
	Dim lo As ListObject, rngOut As Range

    Set master = ThisWorkbook.Worksheets("master")
    Set outStart = master.Range("E1")

	' Count eligible sheets
	For Each ws In ThisWorkbook.Worksheets
		If LCase(ws.Name) <> "master" Then totalSheets = totalSheets + 1
	Next ws

    ' Clear previous consolidated table and headers only from column E onward
    Dim clearRange As Range
    Set clearRange = master.Range(outStart, master.Cells(master.Rows.Count, master.Columns.Count))
    On Error Resume Next
    For Each lo In master.ListObjects
        If Not Intersect(lo.Range, clearRange) Is Nothing Then lo.Unlist
    Next lo
    clearRange.FormatConditions.Delete
    On Error GoTo 0
    clearRange.Clear

	outStart.Resize(1, 6).Value = Array("Sheet", "BM22_Value", "BM23_Value", "HasNegative", "FirstNegativeCell", "FirstNegativeValue")
	nextRow = outStart.Row + 1

	startTime = Timer
	processed = 0

	' Gather from each sheet
	For Each ws In ThisWorkbook.Worksheets
		If LCase(ws.Name) <> "master" Then
			Application.StatusBar = "Reading: " & ws.Name & " (" & (processed + 1) & "/" & totalSheets & ")"
			' Source table starts at CI22 with 5 columns
			Set srcStart = ws.Range("CI22")
			' Find last row with data in first column of source table
			srcLastRow = ws.Cells(ws.Rows.Count, srcStart.Column).End(xlUp).Row
			If srcLastRow >= srcStart.Row + 1 Then
				Dim srcRows As Long: srcRows = srcLastRow - srcStart.Row
				' Write Sheet name then 5 columns of data
				master.Cells(nextRow, outStart.Column).Resize(srcRows, 1).Value = ws.Name
				master.Cells(nextRow, outStart.Column + 1).Resize(srcRows, 5).Value = _
					ws.Range(srcStart.Offset(1, 0), ws.Cells(srcLastRow, srcStart.Column + 4)).Value
				nextRow = nextRow + srcRows
			End If

			processed = processed + 1
			elapsed = Timer - startTime
			If processed > 0 Then
				avg = elapsed / processed
				remaining = Application.Max(0, totalSheets - processed)
				estSec = CLng(avg * remaining)
				Application.StatusBar = "Consolidated " & processed & "/" & totalSheets & _
					". ETA: " & FormatETA(estSec)
			End If
		End If
	Next ws

    ' Style as table with filters and conditional highlights
    If nextRow > outStart.Row + 1 Then
        Set rngOut = master.Range(outStart, master.Cells(nextRow - 1, outStart.Column + 5))
        StyleMaster rngOut
    End If

	Application.StatusBar = False
	MsgBox "Consolidation complete. Rows: " & (nextRow - outStart.Row - 1), vbInformation
End Sub

Private Sub StyleMaster(ByVal rng As Range)
	Dim lo As ListObject, ws As Worksheet
	Dim hasNegCol As Long
	Dim lastRow As Long

	Set ws = rng.Worksheet

	' Clear previous tables/CF inside range
	On Error Resume Next
	For Each lo In ws.ListObjects
		If Not Intersect(lo.Range, rng) Is Nothing Then lo.Unlist
	Next lo
	rng.FormatConditions.Delete
	On Error GoTo 0

	' Create table
	Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
	lo.Name = "MasterResults"
	lo.TableStyle = "TableStyleMedium9"

	' Autofit
	rng.EntireColumn.AutoFit

	' Conditional format HasNegative TRUE
	hasNegCol = rng.Column + 3 ' Sheet=0, BM22=1, BM23=2, HasNegative=3
	With ws.Range(ws.Cells(rng.Row + 1, hasNegCol), ws.Cells(rng.Row + rng.Rows.Count - 1, hasNegCol))
		.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=TRUE"
		With .FormatConditions(.FormatConditions.Count)
			.SetFirstPriority
			With .Interior: .Color = RGB(255, 235, 156): .TintAndShade = 0: End With
			With .Font: .Color = RGB(156, 101, 0): .Bold = True: End With
		End With
	End With

    ' Negative values in last column (FirstNegativeValue)
    With ws.Range(ws.Cells(rng.Row + 1, rng.Column + 5), ws.Cells(rng.Row + rng.Rows.Count - 1, rng.Column + 5))
		.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
		With .FormatConditions(.FormatConditions.Count)
			.SetFirstPriority
			With .Font: .Color = RGB(192, 0, 0): .Bold = True: End With
		End With
        ' Apply numeric format with thousands and 2 decimals
        .NumberFormat = "#,##0.00"
	End With
End Sub
Private Sub StyleResults(ByVal ws As Worksheet, ByVal resultsStart As Range, ByVal numMetricCols As Long)
	Dim lastRow As Long, lastCol As Long
	Dim rng As Range
	Dim lo As ListObject
	Dim hasTable As Boolean
	
	If ws Is Nothing Or resultsStart Is Nothing Then Exit Sub
	
	lastRow = ws.Cells(ws.Rows.Count, resultsStart.Column).End(xlUp).Row
	lastCol = resultsStart.Column + numMetricCols - 1 ' Adjust based on actual number of columns
	If lastRow < resultsStart.Row Or lastCol < resultsStart.Column Then Exit Sub
	Set rng = ws.Range(resultsStart, ws.Cells(lastRow, lastCol))
	
	' Clear old tables/formatting for a clean look (but keep data)
	On Error Resume Next
	For Each lo In ws.ListObjects
		If Not Intersect(lo.Range, rng) Is Nothing Then
			lo.Unlist
		End If
	Next lo
	rng.FormatConditions.Delete
	On Error GoTo 0
	
	' Create table with a nice built-in style
	Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
	lo.Name = "ResultsTable"
	lo.TableStyle = "TableStyleMedium9"
	
	' Autofit columns and set readable widths
	On Error Resume Next
	ws.Range(ws.Cells(resultsStart.Row, resultsStart.Column), ws.Cells(resultsStart.Row, lastCol)).EntireColumn.AutoFit
	On Error GoTo 0
	
	' Number format last column (FirstNegativeValue) with thousands and 2 decimals
	On Error Resume Next
	ws.Columns(lastCol).NumberFormat = "#,##0.00"
	On Error GoTo 0
	
	' Skip freeze panes - user doesn't want it
	
	' Conditional format: highlight HasNegative = TRUE
	Dim hasNegCol As Long: hasNegCol = resultsStart.Column + 2 ' HasNegative is 3rd column
	With ws.Range(ws.Cells(resultsStart.Row + 1, hasNegCol), ws.Cells(lastRow, hasNegCol))
		.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=TRUE"
		.FormatConditions(.FormatConditions.Count).SetFirstPriority
		With .FormatConditions(1).Interior
			.Color = RGB(255, 235, 156) ' soft amber
			.TintAndShade = 0
		End With
		With .FormatConditions(1).Font
			.Color = RGB(156, 101, 0)
			.Bold = True
		End With
	End With
	
	' Conditional format: highlight negative numeric values in metrics columns (last numMetricCols)
	Dim startMetricCol As Long
	startMetricCol = Application.Max(resultsStart.Column, lastCol - numMetricCols + 1)
	With ws.Range(ws.Cells(resultsStart.Row + 1, startMetricCol), ws.Cells(lastRow, lastCol))
		.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
		.FormatConditions(.FormatConditions.Count).SetFirstPriority
		With .FormatConditions(1).Font
			.Color = RGB(192, 0, 0)
			.Bold = True
		End With
	End With
	
	' Add filters already provided by table; ensure header bold
	ws.Rows(resultsStart.Row).Font.Bold = True
End Sub
