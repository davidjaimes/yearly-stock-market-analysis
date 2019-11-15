Attribute VB_Name = "Module1"
Sub FILL_SINGLE_SHEET()
    Dim i, lastrow, counter As Long
    Dim summ As Double
    Dim priceFlag As Boolean
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    counter = 2
    summ = 0
    priceFlag = True
    
    For i = 2 To lastrow / 100
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Save unique ticker symbol in column I.
            Cells(counter, 9).Value = Cells(i, 1).Value
            ' Calculate Yearly Change and save in column J. Also, highlight cell red (negative) or green (positive).
            closePrice = Cells(i, 6).Value
            yearlyChange = closePrice - openPrice
            Cells(counter, 10).Value = yearlyChange
            If yearlyChange < 0 Then
                Cells(counter, 10).Interior.ColorIndex = 3
                Cells(counter, 11).Interior.ColorIndex = 3
            ElseIf yearlyChange > 0 Then
                Cells(counter, 10).Interior.ColorIndex = 4
                Cells(counter, 11).Interior.ColorIndex = 4
            End If
            ' Calculate percent change and save in column K. Careful when dividing by zero!
            If openPrice = 0 Then
                Cells(counter, 11).Value = Format(0#, "#.##%")
            Else
                Cells(counter, 11).Value = Format(yearlyChange / openPrice, "#.##%")
            End If
            ' Save Total Volume in column L.
            summ = summ + Cells(i, 7).Value
            Cells(counter, 12).Value = summ
            ' Reset variables and go to next ticker symbol.
            counter = counter + 1
            summ = 0
            priceFlag = True
        Else
            ' Use flag to save the open price value at the start of the year.
            If priceFlag Then
                openPrice = Cells(i, 3).Value
                priceFlag = False
            End If
            ' If adjacent ticker symbols are the same, then save volume value.
            summ = summ + Cells(i, 7).Value
        End If
    Next i
End Sub
