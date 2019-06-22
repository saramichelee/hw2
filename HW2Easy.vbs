Attribute VB_Name = "Easy"
Sub HW2Easy()

'Loop through all of the worksheets in the active workbook.
For Each ws In Worksheets


Dim ticker As String

Dim total As Double
total = 0

Dim ResultsRow As Integer
ResultsRow = 2

endrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To endrow

    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        ticker = ws.Cells(I, 1).Value
        total = total + ws.Cells(I, 7).Value
        ws.Range("I" & ResultsRow).Value = ticker
        ws.Range("J" & ResultsRow).Value = total
        ResultsRow = ResultsRow + 1
        total = 0
    Else
        total = total + ws.Cells(I, 7).Value
    End If
Next I
Next ws

End Sub
