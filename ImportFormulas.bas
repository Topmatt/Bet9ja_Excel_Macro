Attribute VB_Name = "ImportFormulas"
Option Explicit

' Main subroutine to import formulas using a better input method
Public Sub ImportFormulasFromText()
    Dim formulas As String
    Dim startCell As String
    Dim response As VbMsgBoxResult
    
    ' Check if a cell is selected
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select only one cell as the starting point.", vbExclamation, "Multiple Cells Selected"
        Exit Sub
    End If
    
    ' Use the currently selected cell as starting point
    startCell = Selection.Address
    
    ' Show options for input method
    response = MsgBox("How would you like to provide the formulas?" & vbCrLf & vbCrLf & _
                     "YES = Browse for text file" & vbCrLf & _
                     "NO = Use clipboard (paste first, then click NO)" & vbCrLf & _
                     "CANCEL = Exit", _
                     vbYesNoCancel + vbQuestion, "Formula Input Method")
    
    If response = vbYes Then
        ' Browse for file
        formulas = BrowseForTextFile()
    ElseIf response = vbNo Then
        ' Use clipboard method
        formulas = GetClipboardText()
    Else
        ' User cancelled
        Exit Sub
    End If
    
    ' If we have formulas, import them
    If formulas <> "" Then
        ImportFormulasToExcel formulas, startCell
    End If
End Sub

' Function to get text from clipboard
Private Function GetClipboardText() As String
    Dim dataObj As Object
    Dim response As VbMsgBoxResult
    
    ' Show instruction message
    response = MsgBox("Please copy your formulas to clipboard first, then click OK." & vbCrLf & vbCrLf & _
                     "Make sure each formula is on a separate line in your source.", _
                     vbOKCancel + vbInformation, "Clipboard Method")
    
    If response = vbCancel Then
        GetClipboardText = ""
        Exit Function
    End If
    
    ' Try to get clipboard data
    On Error GoTo ClipboardError
    Set dataObj = CreateObject("MSForms.DataObject")
    dataObj.GetFromClipboard
    GetClipboardText = dataObj.GetText
    Exit Function
    
ClipboardError:
    GetClipboardText = ""
    MsgBox "Could not read from clipboard. Please try the file browse method instead.", vbExclamation, "Clipboard Error"
End Function

' Function to browse for text file
Private Function BrowseForTextFile() As String
    Dim fileDialog As FileDialog
    Dim selectedFile As String
    Dim fileContent As String
    
    ' Create file dialog
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "Select Text File with Formulas"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Text Files", "*.txt"
        .Filters.Add "All Files", "*.*"
        
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
            fileContent = ReadTextFile(selectedFile)
            
            If fileContent <> "" Then
                BrowseForTextFile = fileContent
            Else
                BrowseForTextFile = ""
            End If
        Else
            BrowseForTextFile = ""
        End If
    End With
End Function

' Function to read text file contents
Private Function ReadTextFile(filePath As String) As String
    Dim fileNum As Integer
    Dim fileContent As String
    Dim line As String
    
    fileNum = FreeFile
    fileContent = ""
    
    On Error GoTo FileError
    Open filePath For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        If fileContent = "" Then
            fileContent = line
        Else
            fileContent = fileContent & vbCrLf & line
        End If
    Loop
    
    Close #fileNum
    ReadTextFile = fileContent
    Exit Function
    
FileError:
    If fileNum > 0 Then Close #fileNum
    ReadTextFile = ""
    MsgBox "Error reading file: " & Err.Description, vbCritical, "File Error"
End Function

' Function to extract formula from a line (everything after "=")
Private Function ExtractFormula(line As String) As String
    Dim equalPos As Integer
    Dim formula As String
    
    ' Find the position of "=" in the line
    equalPos = InStr(line, "=")
    
    If equalPos > 0 Then
        ' Extract everything from "=" onwards
        formula = Mid(line, equalPos)
    Else
        ' If no "=" found, use the whole line (might be a formula without =)
        formula = Trim(line)
        ' Add "=" if the line doesn't start with it and looks like a formula
        If Left(formula, 1) <> "=" And formula <> "" Then
            formula = "=" & formula
        End If
    End If
    
    ExtractFormula = Trim(formula)
End Function

' Function to import formulas into Excel
Private Sub ImportFormulasToExcel(formulas As String, startCell As String)
    Dim formulaArray() As String
    Dim i As Integer
    Dim targetRange As Range
    Dim startRange As Range
    Dim formulaCount As Integer
    Dim extractedFormula As String
    
    ' Split formulas by line breaks
    formulaArray = Split(formulas, vbCrLf)
    formulaCount = 0
    
    ' Set the starting range
    Set startRange = Range(startCell)
    
    ' Import each formula
    For i = 0 To UBound(formulaArray)
        If Trim(formulaArray(i)) <> "" Then
            ' Extract only the formula part (after "=")
            extractedFormula = ExtractFormula(formulaArray(i))
            
            ' Only import if we have a valid formula
            If extractedFormula <> "" Then
                ' Calculate the target cell (same column, going down)
                Set targetRange = startRange.Offset(formulaCount, 0)
                
                ' Insert the formula
                targetRange.Formula = extractedFormula
                formulaCount = formulaCount + 1
            End If
        End If
    Next i
    
    ' Show confirmation
    MsgBox "Successfully imported " & formulaCount & " formulas starting from " & startCell, vbInformation, "Import Complete"
End Sub
