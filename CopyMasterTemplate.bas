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

	' Append: set labels and validations on master, and link cells across sheets
	Dim masterWs As Worksheet
	Dim sht As Worksheet
	Set masterWs = ThisWorkbook.Worksheets("master")

	' Labels on master
	masterWs.Range("A21").Value = "Percentage"
	masterWs.Range("A22").Value = "Capital"

	' Data validation on master B21 (percentage list)
	With masterWs.Range("B21").Validation
		.Delete
		.Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, _
			Formula1:=".1%,.2%,.5%,1%,2%,3%,4%,5%,6%"
		.IgnoreBlank = True
		.InCellDropdown = True
	End With

	' Data validation on master B22 (capital list)
	With masterWs.Range("B22").Validation
		.Delete
		.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, _
			Formula1:="500,1000,2000,3000, 5000, 10000,12500,15000, 20000, 30000,50000, 100000, 500000, 1000000"
		.IgnoreBlank = True
		.InCellDropdown = True
	End With

	' Set default selections
	masterWs.Range("B21").Value = "1%"
	masterWs.Range("B22").Value = 5000

	' Link BS21 and BV21 on all non-master sheets to master selections
	For Each sht In ThisWorkbook.Worksheets
		If LCase(sht.Name) <> "master" Then
			With sht.Range("BS21")
				.Clear
				.Formula = "='master'!B21"
			End With
			With sht.Range("BV21")
				.Clear
				.Formula = "='master'!B22"
			End With
		End If
	Next sht

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
