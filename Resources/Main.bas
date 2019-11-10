Attribute VB_Name = "Module11"
Sub main()

    Dim lastRow, ni, nj As Long
    Dim minCount, maxCount As Long
    Dim minHold, maxHold, totalStockVolume As Double
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("H1:K" & lastRow).ClearContents
    
    ' Find unique values in A column and place results in I column.
    Range("A1:A" & lastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True
    Range("I1").Value = "Ticker"
    
    ' Find min and max
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    For ni = 2 To 2
        minCount = 0
        maxCount = 0
        minHold = 1E+99
        maxHold = -1E+99
        totalStockVolume = 0
        For nj = 2 To lastRow
            If Range("A" & nj).Value = Range("I" & ni).Value And Range("B" & nj).Value < minHold Then
                minHold = Range("F" & nj).Value
                minCount = nj
            ElseIf Range("A" & nj).Value = Range("I" & ni).Value And Range("B" & nj).Value > maxHold Then
                maxHold = Range("F" & nj).Value
                maxCount = nj
            ElseIf Range("A" & nj).Value = Range("I" & ni).Value Then
                totalStockVolume = totalStockVolume + Range("G" & nj).Value
            End If
        Next nj
        Range("J" & ni).Value = Format(maxHold - minHold, "#.#########")
        Range("K" & ni).Value = Format((maxHold - minHold) / minHold, "#.##%")
        Range("L" & ni).Value = totalStockVolume
    Next ni
End Sub
