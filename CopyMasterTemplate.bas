Option Explicit

Public Sub CopyFromMasterTemplateToAllSheets()
	' Copies BL21:CG52 from external "Master Template" workbook's "Master" sheet
	' into the same range on all sheets of this workbook except the sheet named "master".
	' Copies values, formulas, and formats.

	Dim sourceWb As Workbook
	Dim sourceWs As Worksheet
	Dim destWs As Worksheet
	Dim templatePath As Variant
	Dim weOpened As Boolean
	Dim originalCalc As XlCalculation

	On Error GoTo CleanFail

	Application.ScreenUpdating = False
	Application.EnableEvents = False
	originalCalc = Application.Calculation
	Application.Calculation = xlCalculationManual

	' Try to get an already-open workbook named "Master Template"
	On Error Resume Next
	Set sourceWb = Workbooks("Master Template.xlsx")
	If sourceWb Is Nothing Then Set sourceWb = Workbooks("Master Template.xlsm")
	If sourceWb Is Nothing Then Set sourceWb = Workbooks("Master Template.xls")
	On Error GoTo CleanFail

	' If not open, prompt user to pick the file
	If sourceWb Is Nothing Then
		templatePath = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , "Select 'Master Template' workbook")
		If VarType(templatePath) = vbBoolean And templatePath = False Then GoTo CleanExit ' user cancelled
		Set sourceWb = Workbooks.Open(CStr(templatePath), ReadOnly:=True)
		weOpened = True
	End If

	' Find the source sheet: prefer a sheet named "Master" (case-insensitive), else use the first sheet
	Set sourceWs = Nothing
	For Each destWs In sourceWb.Worksheets
		If LCase(destWs.Name) = "master" Then
			Set sourceWs = destWs
			Exit For
		End If
	Next destWs
	If sourceWs Is Nothing Then Set sourceWs = sourceWb.Worksheets(1)

	' Copy to each destination worksheet except the master sheet
	For Each destWs In ThisWorkbook.Worksheets
		If LCase(destWs.Name) <> "master" Then
			sourceWs.Range("BL21:CG52").Copy Destination:=destWs.Range("BL21:CG52")
		End If
	Next destWs

CleanExit:
	' Restore app state
	Application.Calculation = originalCalc
	Application.EnableEvents = True
	Application.ScreenUpdating = True

	' If we opened the source workbook, close it without saving
	If weOpened Then
		On Error Resume Next
		sourceWb.Close SaveChanges:=False
		On Error GoTo 0
	End If
	Exit Sub

CleanFail:
	' Attempt to restore state on error, then rethrow
	On Error Resume Next
	Application.Calculation = originalCalc
	Application.EnableEvents = True
	Application.ScreenUpdating = True
	If weOpened And Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
	On Error GoTo 0
	Err.Raise Err.Number, "CopyFromMasterTemplateToAllSheets", Err.Description
End Sub
