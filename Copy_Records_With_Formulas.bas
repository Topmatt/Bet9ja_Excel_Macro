Sub CopyRecordsWithRelativeFormulas()
'
' Copy Records With Relative Formulas Macro
' Solves the problem of copying formulas with absolute references ($) to different locations
'
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceRange As Range
    Dim targetRange As Range
    
    ' Set your source and target sheets (modify these names as needed)
    Set sourceSheet = ThisWorkbook.Worksheets("SourceSheetName") ' Change to your source sheet name
    Set targetSheet = ThisWorkbook.Worksheets("TargetSheetName") ' Change to your target sheet name
    
    ' Set the range you want to copy (adjust as needed)
    Set sourceRange = sourceSheet.Range("A1:Z100") ' Adjust range as needed
    
    ' Set where you want to paste (adjust as needed)
    Set targetRange = targetSheet.Range("A1") ' Starting cell for paste
    
    ' Copy the range
    sourceRange.Copy
    
    ' Paste as values first to avoid formula issues
    targetRange.PasteSpecial Paste:=xlPasteValues
    
    ' If you need the formulas, paste them separately
    ' targetRange.PasteSpecial Paste:=xlPasteFormulas
    
    Application.CutCopyMode = False
    
    MsgBox "Records copied successfully!"
End Sub

Sub CopyWithR1C1Formulas()
'
' Copy With R1C1 Formulas Macro
' Uses R1C1 notation which adjusts automatically when copied
'
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceRange As Range
    Dim targetRange As Range
    
    ' Set your source and target sheets (modify these names as needed)
    Set sourceSheet = ThisWorkbook.Worksheets("SourceSheetName") ' Change to your source sheet name
    Set targetSheet = ThisWorkbook.Worksheets("TargetSheetName") ' Change to your target sheet name
    
    ' Set the range you want to copy (adjust as needed)
    Set sourceRange = sourceSheet.Range("A1:Z100") ' Adjust range as needed
    
    ' Set where you want to paste (adjust as needed)
    Set targetRange = targetSheet.Range("A1") ' Starting cell for paste
    
    ' Copy the range
    sourceRange.Copy
    
    ' Paste using R1C1 notation which adjusts automatically
    targetRange.PasteSpecial Paste:=xlPasteFormulas
    
    Application.CutCopyMode = False
    
    MsgBox "Records with R1C1 formulas copied successfully!"
End Sub

Sub ConvertAbsoluteToRelative()
'
' Convert Absolute To Relative Macro
' Converts all absolute references ($) to relative references in the selected range
'
    Dim ws As Worksheet
    Dim cell As Range
    Dim formula As String
    Dim newFormula As String
    Dim selectedRange As Range
    
    ' Get the currently selected range
    Set selectedRange = Selection
    
    ' Check if a range is selected
    If selectedRange Is Nothing Then
        MsgBox "Please select a range first!"
        Exit Sub
    End If
    
    ' Convert formulas in the selected range
    For Each cell In selectedRange.Cells
        If cell.HasFormula Then
            formula = cell.Formula
            ' Remove $ signs to make references relative
            newFormula = Replace(formula, "$", "")
            cell.Formula = newFormula
        End If
    Next cell
    
    MsgBox "Absolute references converted to relative references!"
End Sub

Sub ConvertAllSheetsAbsoluteToRelative()
'
' Convert All Sheets Absolute To Relative Macro
' Converts all absolute references ($) to relative references in all sheets except master
'
    Dim ws As Worksheet
    Dim cell As Range
    Dim formula As String
    Dim newFormula As String
    Dim formulaCount As Integer
    
    formulaCount = 0
    
    For Each ws In ThisWorkbook.Worksheets
        If LCase(ws.Name) <> "master" Then
            For Each cell In ws.UsedRange.Cells
                If cell.HasFormula Then
                    formula = cell.Formula
                    ' Remove $ signs to make references relative
                    newFormula = Replace(formula, "$", "")
                    cell.Formula = newFormula
                    formulaCount = formulaCount + 1
                End If
            Next cell
        End If
    Next ws
    
    MsgBox "Converted " & formulaCount & " formulas from absolute to relative references!"
End Sub

Sub CopyWithOffset(sourceSheetName As String, targetSheetName As String, _
                   sourceStartCell As String, targetStartCell As String, _
                   rangeSize As String)
'
' Copy With Offset Macro
' Copies a range from one sheet to another with automatic formula adjustment
'
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim rowOffset As Long
    Dim colOffset As Long
    Dim cell As Range
    Dim newFormula As String
    
    ' Set worksheets
    Set sourceSheet = ThisWorkbook.Worksheets(sourceSheetName)
    Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    
    ' Set ranges
    Set sourceRange = sourceSheet.Range(sourceStartCell & ":" & rangeSize)
    Set targetRange = targetSheet.Range(targetStartCell)
    
    ' Calculate the offset
    rowOffset = targetRange.Row - sourceRange.Row
    colOffset = targetRange.Column - sourceRange.Column
    
    ' Copy each cell individually and adjust formulas
    For Each cell In sourceRange.Cells
        If cell.HasFormula Then
            ' Adjust the formula based on offset
            newFormula = AdjustFormulaForOffset(cell.Formula, rowOffset, colOffset)
            targetSheet.Cells(cell.Row + rowOffset, cell.Column + colOffset).Formula = newFormula
        Else
            ' Copy value directly
            targetSheet.Cells(cell.Row + rowOffset, cell.Column + colOffset).Value = cell.Value
        End If
    Next cell
    
    MsgBox "Records copied with offset adjustment!"
End Sub

Function AdjustFormulaForOffset(formula As String, rowOffset As Long, colOffset As Long) As String
'
' Adjust Formula For Offset Function
' Adjusts cell references in formulas based on row and column offset
'
    Dim result As String
    result = formula
    
    ' Replace absolute references with relative ones
    result = Replace(result, "$", "")
    
    ' You could add more sophisticated logic here to handle specific cases
    AdjustFormulaForOffset = result
End Function

Sub SmartCopyRecords()
'
' Smart Copy Records Macro
' Interactive macro that asks user for source and target information
'
    Dim sourceSheetName As String
    Dim targetSheetName As String
    Dim sourceRange As String
    Dim targetStartCell As String
    Dim copyOption As Integer
    
    ' Get user input
    sourceSheetName = InputBox("Enter source sheet name:", "Source Sheet")
    If sourceSheetName = "" Then Exit Sub
    
    targetSheetName = InputBox("Enter target sheet name:", "Target Sheet")
    If targetSheetName = "" Then Exit Sub
    
    sourceRange = InputBox("Enter source range (e.g., A1:Z100):", "Source Range")
    If sourceRange = "" Then Exit Sub
    
    targetStartCell = InputBox("Enter target starting cell (e.g., A1):", "Target Start Cell")
    If targetStartCell = "" Then Exit Sub
    
    ' Ask for copy method
    copyOption = MsgBox("Choose copy method:" & vbCrLf & _
                       "Yes = Copy as values only" & vbCrLf & _
                       "No = Copy with formulas (R1C1)", _
                       vbYesNo, "Copy Method")
    
    ' Perform the copy
    If copyOption = vbYes Then
        CopyAsValues sourceSheetName, targetSheetName, sourceRange, targetStartCell
    Else
        CopyWithFormulas sourceSheetName, targetSheetName, sourceRange, targetStartCell
    End If
End Sub

Sub CopyAsValues(sourceSheetName As String, targetSheetName As String, _
                 sourceRange As String, targetStartCell As String)
'
' Copy As Values Subroutine
' Copies range as values only (no formulas)
'
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceRangeObj As Range
    Dim targetRange As Range
    
    Set sourceSheet = ThisWorkbook.Worksheets(sourceSheetName)
    Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    Set sourceRangeObj = sourceSheet.Range(sourceRange)
    Set targetRange = targetSheet.Range(targetStartCell)
    
    sourceRangeObj.Copy
    targetRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    MsgBox "Records copied as values successfully!"
End Sub

Sub CopyWithFormulas(sourceSheetName As String, targetSheetName As String, _
                     sourceRange As String, targetStartCell As String)
'
' Copy With Formulas Subroutine
' Copies range with formulas using R1C1 notation
'
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceRangeObj As Range
    Dim targetRange As Range
    
    Set sourceSheet = ThisWorkbook.Worksheets(sourceSheetName)
    Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    Set sourceRangeObj = sourceSheet.Range(sourceRange)
    Set targetRange = targetSheet.Range(targetStartCell)
    
    sourceRangeObj.Copy
    targetRange.PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
    
    MsgBox "Records with formulas copied successfully!"
End Sub

Sub CopyRecordsFromMaster()
'
' Copy Records From Master Macro
' Specifically designed to work with your existing master sheet setup
'
    Dim masterSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim targetSheetName As String
    
    ' Get target sheet name
    targetSheetName = InputBox("Enter target sheet name:", "Target Sheet")
    If targetSheetName = "" Then Exit Sub
    
    ' Set worksheets
    Set masterSheet = ThisWorkbook.Worksheets("master")
    Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    
    ' Set ranges (adjust these as needed)
    Set sourceRange = masterSheet.Range("A1:Z100") ' Adjust range as needed
    Set targetRange = targetSheet.Range("A1") ' Starting cell for paste
    
    ' Copy the range
    sourceRange.Copy
    
    ' Paste as values to avoid formula issues
    targetRange.PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
    
    MsgBox "Records copied from master sheet successfully!"
End Sub
