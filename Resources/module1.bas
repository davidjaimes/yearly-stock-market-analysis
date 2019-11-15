Attribute VB_Name = "Module1"
Sub FILL_SINGLE_SHEET()
    Dim i, lastrow, counter As Long
    Dim summ As Double
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    counter = 2
    summ = 0
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Save unique ticker symbols in column I.
            Cells(counter, 9).Value = Cells(i, 1).Value
            ' Save Total Volume in column L.
            Cells(counter, 12).Value = summ
            counter = counter + 1
            summ = 0
        Else
            summ = summ + Cells(i, 7).Value
        End If
    Next i
End Sub
