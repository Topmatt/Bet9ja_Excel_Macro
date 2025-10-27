Option Explicit

' Checks the fill color of a cell and returns "WON" for green (3321689) or "LOST" for other colors
' Usage: =CheckCellColor(E22)
Public Function CheckCellColor(rng As Range) As String
    Select Case rng.Interior.Color
        Case 3321689        ' Green - WON
            CheckCellColor = "WON"
        Case 14277081       ' Red/Orange - LOST
            CheckCellColor = "LOST"
        Case Else           ' No color or other colors
            CheckCellColor = "LOST"  ' Default to LOST
    End Select
End Function

