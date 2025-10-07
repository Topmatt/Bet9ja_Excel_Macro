Sub adding_new_record()
'
' adding_new_record Macro - runs on every sheet except "master"
'

    ' Initialize progress tracking
    Dim startTime As Double
    Dim totalSheets As Integer
    Dim processedSheets As Integer
    Dim estimatedTime As String
    
    ' Variables for Odd_row_position code
    Dim col As Integer
    Dim startRow As Integer
    Dim offset As Integer
    Dim colNames As Variant
    
    startTime = Timer
    totalSheets = 0
    processedSheets = 0
    
    ' Count total sheets first
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase(ws.Name) <> "master" Then
            totalSheets = totalSheets + 1
        End If
    Next ws
    
    ' Create progress bar (no need to set as object)
    
    ' Process each worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase(ws.Name) <> "master" Then
            processedSheets = processedSheets + 1
            
            ' Update progress bar
            Dim progressPercent As Integer
            progressPercent = Int((processedSheets / totalSheets) * 100)
            
            ' Calculate estimated time remaining
            Dim elapsedTime As Double
            Dim avgTimePerSheet As Double
            Dim remainingSheets As Integer
            Dim estimatedSeconds As Integer
            
            elapsedTime = Timer - startTime
            avgTimePerSheet = elapsedTime / processedSheets
            remainingSheets = totalSheets - processedSheets
            estimatedSeconds = Int(avgTimePerSheet * remainingSheets)
            
            If estimatedSeconds < 60 Then
                estimatedTime = estimatedSeconds & " seconds"
            Else
                estimatedTime = Int(estimatedSeconds / 60) & " minutes " & (estimatedSeconds Mod 60) & " seconds"
            End If
            
            ' Update status bar
            Application.StatusBar = "Processing " & ws.Name & "... " & progressPercent & "% complete. Estimated time remaining: " & estimatedTime
            
            With ws
                ' Set SPLIT header
                .Range("C1").FormulaR1C1 = "SPLIT"
                
                ' Fill MAX formula down column C
                .Range("C2").FormulaR1C1 = "=MAX(RC[94]:RC[122])"
                .Range("C2:C17").FillDown
                
                ' Fill IF formula down column CS
                .Range("CS2").FormulaR1C1 = "=IF(RC[-92]="""",1,0)"
                .Range("CS2:CS17").FillDown
                
                ' Fill IF formula across columns CT to DU
                .Range("CT2").FormulaR1C1 = "=IF(RC[-92]="""",RC[-1]+1,0)"
                .Range("CT2").AutoFill Destination:=.Range("CT2:DU2"), Type:=xlFillDefault
                .Range("CT2:DU2").AutoFill Destination:=.Range("CT2:DU17"), Type:=xlFillDefault
                
                ' Clear C18
                .Range("C18").ClearContents
                
                ' --- Begin Odd_row_position code ---
                ' Set header row (1, 2, 3, 4... pattern)
                .Range("BN1").FormulaR1C1 = "1"
                .Range("BO1").FormulaR1C1 = "2"
                .Range("BN1:BO1").AutoFill Destination:=.Range("BN1:CQ1"), Type:=xlFillDefault
                
                ' Fill formulas for each column from BN to CQ
                colNames = Array("BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", _
                                "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ")
                
                For col = 0 To 29  ' BN to CQ = 30 columns
                    startRow = 22 + (col * 12)  ' Starting row pattern: 22, 34, 46, 58...
                    offset = 21 + (col * 12)    ' Offset pattern: 21, 33, 45, 57...
                    
                    ' Set the formula for this column
                    .Range(colNames(col) & "2").FormulaR1C1 = _
                        "=IFERROR(MATCH(RC[-" & (65 + col) & "],R" & startRow & "C1:R" & (startRow + 7) & "C1,0),MATCH(RC[-" & (65 + col) & "],R" & startRow & "C2:R" & (startRow + 7) & "C2,0))+" & offset
                    
                    ' Fill down to row 17
                    .Range(colNames(col) & "2:" & colNames(col) & "17").FillDown
                Next col
                ' --- End Odd_row_position code ---

            End With
            
            ' Force screen update
            Application.ScreenUpdating = True
            DoEvents
        End If
    Next ws
    
    ' Update status for cleanup phase
    Application.StatusBar = "Cleaning up sheets..."
    DoEvents
    
    ' Delete Sheet1 if it exists
    Dim sheetToDelete As Worksheet
    On Error Resume Next
    Set sheetToDelete = ThisWorkbook.Worksheets("Sheet1")
    On Error GoTo 0
    If Not sheetToDelete Is Nothing Then
        Application.DisplayAlerts = False
        sheetToDelete.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Update status for consolidation phase
    Application.StatusBar = "Consolidating data to master sheet..."
    DoEvents
    
    ' Find the maximum value for each column C2:C18 across all non-master sheets
    Dim maxValue As Double
    Dim masterSheet As Worksheet
    Dim cellValue As Variant
    
    Set masterSheet = ThisWorkbook.Worksheets("master")
    
    ' Loop through each column from C2 to C18
    For col = 2 To 18
        ' Update progress for consolidation
        Dim consolidationProgress As Integer
        consolidationProgress = Int(((col - 1) / 17) * 100)
        Application.StatusBar = "Consolidating data... " & consolidationProgress & "% complete"
        DoEvents
        
        ' Initialize maxValue for this column
        maxValue = -999999999
        
        ' Loop through all sheets except master to find the maximum value for this column
        For Each ws In ThisWorkbook.Worksheets
            If LCase(ws.Name) <> "master" Then
                cellValue = ws.Cells(col, 3).Value  ' Column 3 is C
                ' Check if the cell contains a number and is greater than current max
                If IsNumeric(cellValue) And cellValue > maxValue Then
                    maxValue = cellValue
                End If
            End If
        Next ws
        
        ' Write the maximum value to the master sheet's corresponding cell
        masterSheet.Cells(col, 3).Value = maxValue
    Next col
    
    ' Final completion message
    Dim totalTime As Double
    totalTime = Timer - startTime
    Application.StatusBar = "Complete! Total time: " & Format(totalTime, "0.0") & " seconds"
    
    ' Clear status bar after 3 seconds
    Application.OnTime Now + TimeValue("00:00:05"), "ClearStatusBar"
    
End Sub

' Helper sub to clear status bar
Sub ClearStatusBar()
    Application.StatusBar = False
End Sub