Attribute VB_Name = "Module1"
Sub main_array()
    ' Declare all variables
    Dim lastRow, i, tickArray() As Long
    Dim size As Integer
    
    ' Find the total numbers of rows in the data and clear any past contents
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("H1:K" & lastRow).ClearContents
    
    ' Find unique values in A column and place results in I column.
    Range("A1:A" & lastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True
    Range("I1").Value = "Ticker"
    
    ' Get array of rows of for each Ticker symbol
    size = 0
    For i = 1 To lastRow
        If Range("A" & i).Value = Range("I2").Value Then
            ReDim Preserve tickArray(size)
            tickArray(size) = i
            size = size + 1
        End If
    Next i
    
    ' Find Yearly Change
    Range("J1").Value = "Yearly Change"
    startPrice = Range("C" & tickArray(0)).Value
    endPrice = Range("F" & tickArray(size - 1)).Value
    Range("J2").Value = endPrice - startPrice
End Sub
