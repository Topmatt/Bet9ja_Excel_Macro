Option Explicit

Function IsColor(cellRange As Range, colNum As Long, rowNum As Long, R As Long, G As Long, B As Long) As Boolean
    ' Get the specific cell from the range based on column and row numbers
    Dim targetCell As Range
    Set targetCell = cellRange.Cells(rowNum, colNum)

    ' Check the color of the target cell
    IsColor = (targetCell.Interior.Color = RGB(R, G, B))
End Function

Sub FORMAT_ALL_AT_ONCE()
'
'
'
    Dim SheetTotal As Integer
    SheetTotal = Sheets.Count
    Dim SheetNum As Integer
    
    For SheetNum = 1 To SheetTotal
        Sheets(SheetNum).Select
        Rows("1:12").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A1").FormulaR1C1 = "HOME"
        Range("B1").FormulaR1C1 = "AWAY"
        Range("A2").FormulaR1C1 = "=R[12]C"
        Range("A2").AutoFill Destination:=Range("A2:A9"), Type:=xlFillDefault
        Range("A2:A9").AutoFill Destination:=Range("A2:B9"), Type:=xlFillDefault
        Cells.EntireColumn.AutoFit
        Columns("C:C").ColumnWidth = 3.43
        Range("D1").FormulaR1C1 = "1"
        Range("E1").FormulaR1C1 = "X"
        Range("F1").FormulaR1C1 = "2"
        Range("G1").FormulaR1C1 = "1"
        Range("H1").FormulaR1C1 = "X"
        Range("I1").FormulaR1C1 = "2"
        Range("J1").FormulaR1C1 = "O"
        Range("K1").FormulaR1C1 = "U"
        Columns("D:I").ColumnWidth = 6.43
    
        With Range("D1:K1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
    ' Adding Colour to our odds for formatting
    
        Dim i As Long
        
        For i = 22 To 85
            If InStr(1, Cells(i, 4).Value, "sr-selected sr-inactive") > 0 Then
                With Cells(i, 3).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = RGB(89, 175, 50)  ' Background color corresponding to #59AF32
                End With
                With Cells(i, 3).Font
                    .Color = RGB(255, 255, 255) ' White text
                End With
            Else
                With Cells(i, 3).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = RGB(217, 217, 217) ' Background color corresponding to #d9d9d9
                End With
                With Cells(i, 3).Font
                    .Color = RGB(66, 70, 76) ' Font color corresponding to rgba(66, 70, 76, 0.5)
                End With
            End If
        Next i
    
    ' INSERTING ALL ODDS TO THE TABLE
    '
    ' insert_all_odds Macro
    '
        Range("C22:C29").Copy
        Range("D2").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        Range("C30:C37").Copy
        Range("D3").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        Range("C38:C45").Copy
        Range("D4").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        Range("C46:C53").Copy
        Range("D5").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        Range("C54:C61").Copy
        Range("D6").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        Range("C62:C69").Copy
        Range("D7").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        Range("C70:C77").Copy
        Range("D8").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        Range("C78:C85").Copy
        Range("D9").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Columns("K:K").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("K:K,G:G").Clear
        Range("K2").FormulaR1C1 = "=R[12]C[-6]"
        Range("K2").AutoFill Destination:=Range("K2:K9")
    
        With Columns("K:K")
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        Columns("K:K").ColumnWidth = 4.71
        Columns("G:G").ColumnWidth = 1.86
        Range("A1:M1").Font.Bold = True
        With Range("A1:M1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark2
            .TintAndShade = -9.99786370433668E-02
            .PatternTintAndShade = 0
        End With
        
        ' convert_to_text TO NUMBER
'

'
        Columns("D:D").Select
        Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        Columns("E:E").Select
        Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("F:F").Select
        Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("H:H").Select
        Selection.TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("I:I").Select
        Selection.TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("J:J").Select
        Selection.TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("L:L").Select
        Selection.TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Columns("M:M").Select
        Selection.TextToColumns Destination:=Range("M1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
    
    ' CREATING A DASHBOARD TO CHECK THE STRATEGY MATCH
    '
        Range("O1").FormulaR1C1 = "COL"
        Range("O2").FormulaR1C1 = _
            "=MATCH(MIN(RC[-11]:RC[-9]),RC[-11]:RC[-9],0)+COLUMN(RC[-11])-1"
        Range("O2").AutoFill Destination:=Range("O2:O9"), Type:=xlFillDefault
        Range("P1").FormulaR1C1 = "ROW"
        Range("P2").Value = 2
        Range("P3").Value = 3
        Range("P4").Value = 4
        Range("P5").Value = 5
        Range("P6").Value = 6
        Range("P7").Value = 7
        Range("P8").Value = 8
        Range("P9").Value = 9
        Range("Q2").FormulaR1C1 = _
            "=IF(IsColor(R1C1:R9C6,RC[-2],RC[-1],89, 175, 50),""YES"","""")"
        Range("Q3").FormulaR1C1 = _
            "=IF(IsColor(R1C1:R9C6,RC[-2],RC[-1],89, 175, 50),""YES"","""")"
        Range("Q4").FormulaR1C1 = _
            "=IF(IsColor(R1C1:R9C6,RC[-2],RC[-1],89, 175, 50),""YES"","""")"
        Range("Q5").FormulaR1C1 = _
            "=IF(IsColor(R1C1:R9C6,RC[-2],RC[-1],89, 175, 50),""YES"","""")"
        Range("Q6").FormulaR1C1 = _
            "=IF(IsColor(R1C1:R9C6,RC[-2],RC[-1],89, 175, 50),""YES"","""")"
        Range("Q7").FormulaR1C1 = _
            "=IF(IsColor(R1C1:R9C6,RC[-2],RC[-1],89, 175, 50),""YES"","""")"
        Range("Q8").FormulaR1C1 = _
            "=IF(IsColor(R1C1:R9C6,RC[-2],RC[-1],89, 175, 50),""YES"","""")"
        Range("Q9").FormulaR1C1 = _
            "=IF(IsColor(R1C1:R9C6,RC[-2],RC[-1],89, 175, 50),""YES"","""")"
        Range("R1").FormulaR1C1 = "TEAMS"
        Range("A2:B9").Select
    Selection.Copy
    Range("R2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("S2:S9").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("R10").Select
    ActiveSheet.Paste
        Range("R2:R17").Select
        ActiveWorkbook.Worksheets(SheetNum).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(SheetNum).Sort.SortFields.Add Key:=Range("R2"), SortOn _
            :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets(SheetNum).Sort
            .SetRange Range("R2:R17")
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        Range("S1").FormulaR1C1 = "ROW NO."
        Range("S2").FormulaR1C1 = _
            "=IFERROR(MATCH(RC[-1],R1C1:R9C1,0),MATCH(RC[-1],R1C2:R9C2,0))"
        Range("S2").AutoFill Destination:=Range("S2:S17")
        Range("T1").FormulaR1C1 = "MATCH"
        Range("T2").FormulaR1C1 = "=INDEX(R1C17:R9C17,RC[-1])"
        Range("T2").AutoFill Destination:=Range("T2:T17")
        Columns("R:R").EntireColumn.AutoFit
    '   HIDECELL THE CALCULATION SHEET
        Columns("O:Q").EntireColumn.Hidden = True
        Columns("S:S").EntireColumn.Hidden = True
        
        ' ===== Double chance columns, headers, formulas, copy, formats, and alignment =====
        Dim ws As Worksheet
        Set ws = ActiveSheet

        ' Insert new columns H to K for Double Chance calculations
        ws.Columns("H:K").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Columns("H:J").ColumnWidth = 6.43

        ' Add Double Chance headers
        ws.Range("H1").Value = "1X"
        ws.Range("I1").Value = "12"
        ws.Range("J1").Value = "X2"

        ' Add formulas for Double Chance values in row 2
        ws.Range("H2").FormulaR1C1 = "=1/((1/RC[-4])+(1/RC[-3]))"   ' 1X
        ws.Range("I2").FormulaR1C1 = "=1/((1/RC[-5])+(1/RC[-3]))"   ' 12
        ws.Range("J2").FormulaR1C1 = "=1/((1/RC[-5])+(1/RC[-4]))"   ' X2

        ' Fill formulas down from row 2 to 9 and format as numeric with 2 decimals
        ws.Range("H2:J2").AutoFill Destination:=ws.Range("H2:J9"), Type:=xlFillDefault
        ws.Range("H2:J9").NumberFormat = "0.00"

        ' Insert and prepare O to R columns for results with spacing and clearing old formats
        ws.Columns("O:R").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Range("O2:R9").ClearFormats
        ws.Columns("O:O").ColumnWidth = 2

        ' Copy Double Chance headers to P1, which is in the newly inserted block
        ws.Range("H1:J1").Copy Destination:=ws.Range("P1")

        ' Format all headers H1:R1: center, no wrap, no merge, etc.
        With ws.Range("H1:R1")
            .HorizontalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With

        ' Copy all Double Chance formulas to P2:R9 as formulas only
        ws.Range("H2:J9").Copy
        ws.Range("P2").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        ws.Range("P2:R9").NumberFormat = "0.00"

        ' Insert a final single spacing column at S (right of results)
        ws.Columns("S:S").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Columns("S:S").ColumnWidth = 2

        ' Add U/O headers after operations
        ws.Range("U1").Value = "U"
        ws.Range("V1").Value = "O"

        ' ===== Double chance coloring logic for H2:J9 (D/E/F) and P2:R9 (L/M/N) =====
        Dim greenCode As Long, ashCode As Long
        greenCode = 3321689
        ashCode = 14277081
        For i = 2 To 9
            ' H2:H9 (1X): D or E green
            If ws.Cells(i, "D").Interior.Color = greenCode Or ws.Cells(i, "E").Interior.Color = greenCode Then
                ws.Cells(i, "H").Interior.Color = greenCode
                ws.Cells(i, "H").Font.Color = vbWhite
            Else
                ws.Cells(i, "H").Interior.Color = ashCode
                ws.Cells(i, "H").Font.Color = vbBlack
            End If
            ' I2:I9 (12): D or F green
            If ws.Cells(i, "D").Interior.Color = greenCode Or ws.Cells(i, "F").Interior.Color = greenCode Then
                ws.Cells(i, "I").Interior.Color = greenCode
                ws.Cells(i, "I").Font.Color = vbWhite
            Else
                ws.Cells(i, "I").Interior.Color = ashCode
                ws.Cells(i, "I").Font.Color = vbBlack
            End If
            ' J2:J9 (X2): E or F green
            If ws.Cells(i, "E").Interior.Color = greenCode Or ws.Cells(i, "F").Interior.Color = greenCode Then
                ws.Cells(i, "J").Interior.Color = greenCode
                ws.Cells(i, "J").Font.Color = vbWhite
            Else
                ws.Cells(i, "J").Interior.Color = ashCode
                ws.Cells(i, "J").Font.Color = vbBlack
            End If
            ' P2:P9 (1X): L or M green
            If ws.Cells(i, "L").Interior.Color = greenCode Or ws.Cells(i, "M").Interior.Color = greenCode Then
                ws.Cells(i, "P").Interior.Color = greenCode
                ws.Cells(i, "P").Font.Color = vbWhite
            Else
                ws.Cells(i, "P").Interior.Color = ashCode
                ws.Cells(i, "P").Font.Color = vbBlack
            End If
            ' Q2:Q9 (12): L or N green
            If ws.Cells(i, "L").Interior.Color = greenCode Or ws.Cells(i, "N").Interior.Color = greenCode Then
                ws.Cells(i, "Q").Interior.Color = greenCode
                ws.Cells(i, "Q").Font.Color = vbWhite
            Else
                ws.Cells(i, "Q").Interior.Color = ashCode
                ws.Cells(i, "Q").Font.Color = vbBlack
            End If
            ' R2:R9 (X2): M or N green
            If ws.Cells(i, "M").Interior.Color = greenCode Or ws.Cells(i, "N").Interior.Color = greenCode Then
                ws.Cells(i, "R").Interior.Color = greenCode
                ws.Cells(i, "R").Font.Color = vbWhite
            Else
                ws.Cells(i, "R").Interior.Color = ashCode
                ws.Cells(i, "R").Font.Color = vbBlack
            End If
        Next i
    Next SheetNum
' MASTER PAGE CREATOR
'

' Sheetspace = 12
' Cells(13,2).value = Round

    Dim SheetSpace As Integer
    Dim RoundSheet As Integer
    Dim RoundTop As Integer
    Sheets(1).Activate
    Worksheets.Add
    SheetTotal = Sheets.Count
    Sheets(1).Name = "Master"
    SheetTotal = Sheets.Count
    SheetSpace = 20
    RoundTop = 3
    
    
    'labelling page
    Sheets("1").Select
    Range("A2:B9").Select
    Selection.Copy
    Sheets("Master").Select
    Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B2:B9").Select
    Selection.Cut
    Range("A10").Select
    ActiveSheet.Paste
    Range("A2:A17").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Master").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Master").Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Master").Sort
        .SetRange Range("A2:A17")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 
    Range("A1").Value = "TEAMS"
    Range("B1").Value = "WINNER"
    
    

    
    For SheetNum = 1 To SheetTotal
        RoundSheet = SheetNum + 1
        If RoundSheet > SheetTotal Then Exit For
        Sheets("Master").Activate
        Cells(SheetSpace, 1).FormulaR1C1 = "Round"
        Cells(SheetSpace, 1).Offset(0, 1).Value = SheetNum
        Cells(SheetSpace + 1, 1).Select
        Sheets(RoundSheet).Select
        Range("A1:V9").Copy
        Range("A1:V9").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.Copy
        
        Sheets("Master").Activate
        ActiveSheet.Paste
        'let's copy round table
        Cells(1, RoundTop).FormulaR1C1 = SheetNum
        Cells(2, RoundTop).Select
        Sheets(RoundSheet).Select
        Range("AC2:AC17").Select
        Range("AC2:AC17").Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.Copy
        Sheets("Master").Activate
        ActiveSheet.Paste



        SheetSpace = SheetSpace + 12
        RoundTop = RoundTop + 1
        
    Next SheetNum
    
    Cells.EntireColumn.AutoFit
    Columns("D:M").ColumnWidth = 6.43
    Columns("C:C").ColumnWidth = 3.43
    Columns("K:K").ColumnWidth = 4.71
    Columns("G:G").ColumnWidth = 1.86
    Columns("P:R").ColumnWidth = 6.43
    Columns("U:V").ColumnWidth = 6.43
    
    'Adding the calculated streak rate
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=MAX(RC[31]:RC[60])"
    Range("B2").AutoFill Destination:=Range("B2:B17")
    Range("AG2").FormulaR1C1 = "=IF(RC[-30]=""YES"",1,0)"
    Range("AG2").AutoFill Destination:=Range("AG2:AG17"), Type:=xlFillDefault
    Range("AH2").FormulaR1C1 = "=IF(RC[-30]=""YES"",RC[-1]+1,0)"
    Range("AH2").AutoFill Destination:=Range("AH2:AH17"), Type:=xlFillDefault
    Range("AH2").AutoFill Destination:=Range("AH2:BJ2"), Type:=xlFillDefault
    Range("BJ2").AutoFill Destination:=Range("BJ2:BJ17")
    Range("BI2").AutoFill Destination:=Range("BI2:BI17")
    Range("BH2").AutoFill Destination:=Range("BH2:BH17")
    Range("BG2").AutoFill Destination:=Range("BG2:BG17")
    Range("BF2").AutoFill Destination:=Range("BF2:BF17")
    Range("BE2").AutoFill Destination:=Range("BE2:BE17")
    Range("BD2").AutoFill Destination:=Range("BD2:BD17")
    Range("BC2").AutoFill Destination:=Range("BC2:BC17")
    Range("BB2").AutoFill Destination:=Range("BB2:BB17")
    Range("BA2").AutoFill Destination:=Range("BA2:BA17")
    Range("AZ2").AutoFill Destination:=Range("AZ2:AZ17")
    Range("AY2").AutoFill Destination:=Range("AY2:AY17")
    Range("AX2").AutoFill Destination:=Range("AX2:AX17")
    Range("AW2").AutoFill Destination:=Range("AW2:AW17")
    Range("AV2").AutoFill Destination:=Range("AV2:AV17")
    Range("AU2").AutoFill Destination:=Range("AU2:AU17")
    Range("AT2").AutoFill Destination:=Range("AT2:AT17")
    Range("AS2").AutoFill Destination:=Range("AS2:AS17")
    Range("AR2").AutoFill Destination:=Range("AR2:AR17")
    Range("AQ2").AutoFill Destination:=Range("AQ2:AQ17")
    Range("AP2").AutoFill Destination:=Range("AP2:AP17")
    Range("AO2").AutoFill Destination:=Range("AO2:AO17")
    Range("AN2").AutoFill Destination:=Range("AN2:AN17")
    Range("AM2").AutoFill Destination:=Range("AM2:AM17")
    Range("AL2").AutoFill Destination:=Range("AL2:AL17")
    Range("AK2").AutoFill Destination:=Range("AK2:AK17")
    Range("AJ2").AutoFill Destination:=Range("AJ2:AJ17")
    Range("AI2").AutoFill Destination:=Range("AI2:AI17")
    Columns("AG:BJ").EntireColumn.Hidden = True
    Range("A18").FormulaR1C1 = "HIGHEST NUMBER"
    Range("B18").FormulaR1C1 = "=MAX(R[-16]C:R[-1]C)"

' Adding the count row
    Range("C1:C18").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:= _
        "Count 1, Count 2, Count 3, Count 4, Count 5, Count 6, Count 7, Count 8, Count 9, Count 10"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("C2").FormulaR1C1 = "=COUNTIF(RC[31]:RC[60],RIGHT(R1C3,1))"
    Range("C2").AutoFill Destination:=Range("C2:C17")
    Range("C18").FormulaR1C1 = "=SUM(R[-16]C:R[-1]C)"
    Columns("C:C").ColumnWidth = 7
    
End Sub
Sub delete_modifications()
'
' delete Macro
'

    Dim SheetTotal As Integer
    Sheets(1).Delete
    SheetTotal = Sheets.Count
    Dim SheetNum As Integer
    
    For SheetNum = 1 To SheetTotal
        Sheets(SheetNum).Activate
        Rows("1:12").Delete Shift:=xlUp
        Columns("N:T").Delete Shift:=xlToLeft

    Next SheetNum
End Sub

Sub CopyMasterSheetsAlphabetically()
    Dim folderPath As String, sourceFile As String
    Dim sourceWorkbook As Workbook, sourceWorksheet As Worksheet
    Dim targetWorkbook As Workbook
    Dim filenames() As String, i As Long, j As Long
    Dim tempName As String

    ' Get the folder path from the user
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Excel Files with MASTER Sheets"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "Folder selection cancelled.", vbInformation
            Exit Sub
        End If
    End With

    ' Set reference to the currently open workbook
    Set targetWorkbook = ThisWorkbook

    ' Collect filenames in the folder
    sourceFile = Dir(folderPath & "\*.xls*")
    ReDim filenames(0)
    Do While sourceFile <> ""
        ReDim Preserve filenames(UBound(filenames) + 1)
        filenames(UBound(filenames)) = sourceFile
        sourceFile = Dir()
    Loop

    ' Check if any files were found
    If UBound(filenames) = -1 Then
        MsgBox "No Excel files found in the selected folder.", vbExclamation
        Exit Sub
    End If

    ' Sort filenames alphabetically (case-insensitive)
    For i = 0 To UBound(filenames) - 1
        For j = i + 1 To UBound(filenames)
            If LCase(filenames(i)) > LCase(filenames(j)) Then
                tempName = filenames(i)
                filenames(i) = filenames(j)
                filenames(j) = tempName
            End If
        Next j
    Next i

    ' Copy MASTER sheets in alphabetical order and name after source file
    For i = 0 To UBound(filenames)
        sourceFile = filenames(i)

        ' Open the workbook (with error handling)
        On Error Resume Next
        Set sourceWorkbook = Workbooks.Open(Filename:=folderPath & "\" & sourceFile)
        On Error GoTo 0

        If Not sourceWorkbook Is Nothing Then ' Check if file opened successfully
            For Each sourceWorksheet In sourceWorkbook.Worksheets
                If LCase(sourceWorksheet.Name) = "master" Then
                    ' Copy the sheet
                    sourceWorksheet.Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count)

                    ' Rename the copied sheet to the source file name (without extension)
                    targetWorkbook.ActiveSheet.Name = Left(sourceFile, InStrRev(sourceFile, ".") - 1)

                    Exit For ' Found the sheet
                End If
            Next sourceWorksheet

            sourceWorkbook.Close SaveChanges:=False
        End If ' End of check if file opened successfully
    Next i

    ' Sort sheets in the target workbook by name
    Dim tempSheet As Worksheet
    With targetWorkbook
        For i = 1 To .Sheets.Count - 1
            For j = i + 1 To .Sheets.Count
                If LCase(.Sheets(i).Name) > LCase(.Sheets(j).Name) Then
                    ' Swap sheet positions
                    Set tempSheet = .Sheets(i)
                    .Sheets(i).Move After:=.Sheets(j)
                    .Sheets(j).Move After:=tempSheet
                End If
            Next j
        Next i
    End With


    'MsgBox "MASTER sheets copied and named alphabetically!", vbInformation
    Dim ws As Worksheet
    Dim maxVal As Double
    Dim targetCell As Range
    Dim currentVal As Double
    
    Worksheets(1).Select
    Sheets.Add
    ActiveSheet.Name = "MASTER"
    ActiveSheet.Next.Select ' Selects the next sheet
    Range("A1:B18").Select
    Selection.Copy
    Sheets("MASTER").Select
    ActiveSheet.Paste
    Columns("A:A").EntireColumn.AutoFit

    maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B2") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B2").FormulaR1C1 = maxVal
    
' =========Calculation for Range B3================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B3") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B3").FormulaR1C1 = maxVal
' =========Calculation for Range B4================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B4") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B4").FormulaR1C1 = maxVal
' =========Calculation for Range B5================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B5") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B5").FormulaR1C1 = maxVal
' =========Calculation for Range B6================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B6") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B6").FormulaR1C1 = maxVal
' =========Calculation for Range B7================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B7") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B7").FormulaR1C1 = maxVal
' =========Calculation for Range B8================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B8") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B8").FormulaR1C1 = maxVal
' =========Calculation for Range B9================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B9") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B9").FormulaR1C1 = maxVal
' =========Calculation for Range B10================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B10") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B10").FormulaR1C1 = maxVal
' =========Calculation for Range B11================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B11") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B11").FormulaR1C1 = maxVal
' =========Calculation for Range B12================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B12") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B12").FormulaR1C1 = maxVal
' =========Calculation for Range B13================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B13") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B13").FormulaR1C1 = maxVal
' =========Calculation for Range B14================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B14") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B14").FormulaR1C1 = maxVal
' =========Calculation for Range B15================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B15") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B15").FormulaR1C1 = maxVal
' =========Calculation for Range B16================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B16") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B16").FormulaR1C1 = maxVal
' =========Calculation for Range B17================
        maxVal = -1E+308 ' Initialize to a very small number

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        Set targetCell = ws.Range("B17") ' Cell to check

        ' Check if the cell contains a number
        If IsNumeric(targetCell.Value) Then
            currentVal = targetCell.Value

            ' Update maxVal if a larger number is found
            If currentVal > maxVal Then
                maxVal = currentVal
            End If
        End If
    Next ws

    ' Display the maximum value
    Range("B17").FormulaR1C1 = maxVal
    
'============ Conditional Formatting ==================
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=7"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
