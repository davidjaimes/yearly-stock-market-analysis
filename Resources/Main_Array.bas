Attribute VB_Name = "Module1"
Sub main_array()
    ' Declare all variables
    Dim lastRow, i, tickArray() As Long
    
    ' Find the total numbers of rows in the data and clear any past contents
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("H1:K" & lastRow).ClearContents
    
    ' Find unique values in A column and place results in I column.
    Range("A1:A" & lastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True
    Range("I1").Value = "Ticker"
    
    ' Get array of rows of for each Ticker symbol
    For i = 1 To lastRow
        If Range("A" & i).Value = Range("I2").Value Then
            tickArray() = i
        End If
    Next i
End Sub
