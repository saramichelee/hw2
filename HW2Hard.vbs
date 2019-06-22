Attribute VB_Name = "Hard"
Sub HW2Hardd()
For Each ws In Worksheets
Dim ticker As String

Dim total As Double
total = 0

Dim change As Double
change = 0

Dim yearopen As Double
yearopen = 0

Dim ResultsRow As Integer
ResultsRow = 2

endrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim maxpercent As Double
Dim minpercent As Double
Dim maxvol As Double

For I = 2 To endrow
    
    If ws.Cells(I + 1, 1).Value = ws.Cells(I, 1).Value And ws.Cells(I - 1, 1).Value <> ws.Cells(I, 1).Value Then
        total = total + ws.Cells(I, 7).Value
        change = change + (ws.Cells(I, 6).Value - ws.Cells(I, 3))
        yearopen = ws.Cells(I, 3).Value
    ElseIf ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        ticker = ws.Cells(I, 1).Value
        total = total + ws.Cells(I, 7).Value
        change = change + (ws.Cells(I, 6).Value - ws.Cells(I, 3))
        ws.Range("I" & ResultsRow).Value = ticker
        ws.Range("J" & ResultsRow).Value = total
        ws.Range("K" & ResultsRow).Value = change
        Select Case change
            Case Is < 0
                ws.Range("K" & ResultsRow).Interior.ColorIndex = 3
            Case Is > 0
                ws.Range("K" & ResultsRow).Interior.ColorIndex = 10
        End Select
        ws.Range("L" & ResultsRow).Value = change / yearopen
        ws.Range("L" & ResultsRow).NumberFormat = "0.0000%"
        ResultsRow = ResultsRow + 1
        total = 0
        change = 0
        yearopen = 0
    Else
        total = total + ws.Cells(I, 7).Value
        change = change + (ws.Cells(I, 6).Value - ws.Cells(I, 3))
    End If

Next I

ws.Range("O1") = "Value"
ws.Range("P1") = "Ticker"

maxpercent = WorksheetFunction.Max(ws.Range("L:L"))
ws.Cells(2, 15).Value = maxpercent
ws.Cells(2, 15).NumberFormat = "0.0000%"
ws.Cells(2, 14).Value = "Greatest % Increase"

minpercent = WorksheetFunction.Min(ws.Range("L:L"))
ws.Cells(3, 15).Value = minpercent
ws.Cells(3, 15).NumberFormat = "0.0000%"
ws.Cells(3, 14).Value = "Greatest % Decrease"

maxvol = WorksheetFunction.Max(ws.Range("J:J"))
ws.Cells(4, 15).Value = maxvol
ws.Cells(4, 14).Value = "Greatest Total Volume"

endrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
For t = 2 To endrow2
    If ws.Cells(t, 12) = maxpercent Then
        ws.Cells(2, 16).Value = ws.Cells(t, 9).Value
    End If
    If ws.Cells(t, 12) = minpercent Then
        ws.Cells(3, 16).Value = ws.Cells(t, 9).Value
    End If
        If ws.Cells(t, 10) = maxvol Then
        ws.Cells(4, 16).Value = ws.Cells(t, 9).Value
    End If
Next t
Next ws

End Sub

