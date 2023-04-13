Attribute VB_Name = "Module1"
Sub Stocks()

    Dim Ticker As String
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Range("I" & 1).Value = "Ticker"
    Range("J" & 1).Value = "Yearly Change"
    Range("K" & 1).Value = "Percent Change"
    Range("L" & 1).Value = "Total Stock Volume"
    Dim Vol As Double
    Vol = 0
    
    
    For i = 2 To LastRow
        Ticker = Cells(i, 1).Value
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            opening = Cells(i, 3).Value
        End If
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            closing = Cells(i, 6).Value
            Change = (closing - opening)
            Percent_change = ((closing - opening) / opening)
            Vol = Vol + Cells(i, 7).Value

            If Change > 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            Range("I" & Summary_Table_Row).Value = Ticker
            Range("J" & Summary_Table_Row).Value = Change
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            Range("K" & Summary_Table_Row).Value = Percent_change
            Range("L" & Summary_Table_Row).Value = Vol
            Summary_Table_Row = Summary_Table_Row + 1
            Vol = 0
        Else
            Vol = Vol + Cells(i, 7).Value
        End If
        
    Next i
    
                    
       'change = (close - open) respect to first line(open) and last line(close)
       'Percent change  = (close - open)/ open
       
End Sub


    
