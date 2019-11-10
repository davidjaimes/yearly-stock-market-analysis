Attribute VB_Name = "Module1"
Sub MAIN_PROGRAM()
    ' Declare all variables for MAIN_PROGRAM
    Dim lastRow, lastRowUnique, i, j As Long
    ' Find the total numbers of rows in column A and clear any past contents in cells
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("H1:K" & lastRow).ClearContents
    ' Find unique values in A column and place results in I column.
    Range("A1:A" & lastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True
    ' Fill in Header Names between Columns I and L
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    ' Find the total number of rows in the Ticker column (i.e., I column)
    lastRowUnique = Range("I" & Rows.Count).End(xlUp).Row
    ' Call FILL_FUNCTION
    For j = 2 To 20
        Call FILL_FUNCTION(lastRow, j)
    Next j
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
    ' Find Yearly Change
    startPrice = Range("C" & tickArray(0)).Value
    endPrice = Range("F" & tickArray(size - 1)).Value
    Range("J" & j).Value = endPrice - startPrice
    ' Find Percent Change
    Range("K" & j).Value = Format((endPrice - startPrice) / startPrice, "#.##%")
    ' Find the Total Stock Volume
    Range("L" & j).Value = summ
End Sub
