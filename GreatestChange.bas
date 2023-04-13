Attribute VB_Name = "Module2"
Sub GreatestChange()
    Range("P" & 1).Value = "Ticker"
    Range("Q" & 1).Value = "Value"

    Range("O" & 2).Value = "Greatest % Increase"
    Range("O" & 3).Value = "Greatest % Decrease"
    Range("O" & 4).Value = "Greatest Total Volume"

    LastRow = Cells(Rows.Count, 9).End(xlUp).Row
    Dim incTicker As String
    Dim decTicker As String
    Dim volTicker As String
    Dim decrease As Double
    decrease = 0
    Dim increase As Double
    increase = 0
    Dim volume As Double
    volume = 0

    For i = 2 To LastRow
    
        If Cells(i, 11) < 0 Then
            If Cells(i, 11) < decrease Then
                decrease = Cells(i, 11).Value
                decTicker = Cells(i, 9).Value
            End If
        End If
        
        If Cells(i, 11) > 0 Then
            If Cells(i, 11) > increase Then
                increase = Cells(i, 11).Value
                incTicker = Cells(i, 9).Value
            End If
        End If
        
        If Cells(i, 12) > volume Then
            volume = Cells(i, 12).Value
            volTicker = Cells(i, 9).Value
        End If
    
        Range("P" & 2).Value = incTicker
        Range("P" & 3).Value = decTicker
        Range("P" & 4).Value = volTicker
        
        Range("Q" & 2).NumberFormat = "0.00%"
        Range("Q" & 2).Value = increase
        Range("Q" & 3).NumberFormat = "0.00%"
        Range("Q" & 3).Value = decrease
        Range("Q" & 4).Value = volume
    Next i
    
End Sub
