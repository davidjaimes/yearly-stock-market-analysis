Attribute VB_Name = "Module1"
Sub MAIN_PROGRAM()
    ' Declare all variables for MAIN_PROGRAM
    Dim lastRow, lastRowUnique, i, j As Long
    ' Find the total numbers of rows in column A and clear any past contents in columns I to Q
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("I1:Q" & lastRow).ClearContents
    Range("I1:Q" & lastRow).ClearFormats
    ' Find unique values in A column and place results in I column.
    Range("A1:A" & lastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True
    ' Fill in Header Names between Columns I and L
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    ' Find the total number of rows in the Ticker column (i.e., I column)
    lastRowUnique = Range("I" & Rows.Count).End(xlUp).Row
    ' Call FILL_FUNCTION to fill column J, K, and L.
    For j = 2 To 10
        Call FILL_FUNCTION(lastRow, j)
    Next j
    'Call CHALLANGES to find greatest percent increase/decrease and greatest volume
    Call CHALLANGES(lastRowUnique)
End Sub

Sub FILL_FUNCTION(lastRow, j)
    ' Declare all variables for FILL_FUNCTION
    Dim tickArray() As Long
    Dim size As Integer
    Dim summ As Double
    ' Get array of rows of for each Ticker symbol
    size = 0
    summ = 0
    For i = 1 To lastRow
        If Range("A" & i).Value = Range("I" & j).Value Then
            ReDim Preserve tickArray(size)
            tickArray(size) = i
            size = size + 1
            summ = summ + Range("G" & i).Value
        End If
    Next i
    ' Find Yearly Change and highlight cell either red or green
    startPrice = Range("C" & tickArray(0)).Value
    endPrice = Range("F" & tickArray(size - 1)).Value
    yearlyChange = endPrice - startPrice
    If yearlyChange < 0 Then
        Range("J" & j).Interior.ColorIndex = 3
    ElseIf yearlyChange > 0 Then
        Range("J" & j).Interior.ColorIndex = 4
    End If
    Range("J" & j).Value = endPrice - startPrice
    ' Find Percent Change
    Range("K" & j).Value = Format((endPrice - startPrice) / startPrice, "#.##%")
    ' Find the Total Stock Volume
    Range("L" & j).Value = summ
End Sub

Sub CHALLANGES(lastRowChal)
' Fill in Header Names columns O, P and Q
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
' Find the values for greatest increase, decrease, and volume
greatIncrease = Application.WorksheetFunction.Max(Range("K2:K" & lastRowChal))
Range("Q2").Value = Format(greatIncrease, "#.##%")
greatDecrease = Application.WorksheetFunction.Min(Range("K2:K" & lastRowChal))
Range("Q3").Value = Format(greatDecrease, "#.##%")
greatVolume = Application.WorksheetFunction.Max(Range("L2:L" & lastRowChal))
Range("Q4").Value = greatVolume
' Find Index for P column and return matching ticker symbol
indexIncrease = Application.WorksheetFunction.Match(greatIncrease, Range("K2:K" & lastRowChal), 0)
Range("P2").Value = Range("I" & indexIncrease + 1).Value
indexDecrease = Application.WorksheetFunction.Match(greatDecrease, Range("K2:K" & lastRowChal), 0)
Range("P3").Value = Range("I" & indexDecrease + 1).Value
indexVolume = Application.WorksheetFunction.Match(greatVolume, Range("L2:L" & lastRowChal), 0)
Range("P4").Value = Range("I" & indexVolume + 1).Value
End Sub
