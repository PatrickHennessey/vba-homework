Sub TickerEasyMode()

      ' Set an initial variable for holding the ticker symbol, ticker volume, summary table, and last row.
      Dim Ticker_Symbol As String
      Dim Ticker_Volume As Double
      Dim Summary_Table_Row As Integer
      Dim LastRow As Long
      Dim WS As Worksheet
      
      For Each WS In Worksheets
      
        ' Set initial vaules for the variables used.
        Ticker_Total = 0
        Summary_Table_Row = 2
        LastRow = WS.Cells(Rows.Count, "A").End(xlUp).Row + 1
    
        ' Name New Columns
        WS.Cells(1, 9).Value = "Ticker Symbol"
        WS.Cells(1, 10).Value = "Ticker Stock Value"
        
        ' Loop through all the unique ticker symbols unitl the last row.
        For i = 2 To LastRow
    
            ' Check if we are still within the same ticker symbol, if it is not...
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    
            ' Set the Ticker_Symbol name
            Ticker_Symbol = WS.Cells(i, 1).Value
    
            ' Add to the Ticker Total
            Ticker_Total = Ticker_Total + WS.Cells(i, 7).Value
    
            ' Print the Ticker Symbol in the Summary Table
            WS.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
    
            ' Print the Ticker_Symbol Total to the Summary Table
            WS.Range("J" & Summary_Table_Row).Value = Ticker_Total
    
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the Brand Total
            Ticker_Total = 0
    
            ' If the cell immediately following a row is the same ticker
            Else
    
            ' Add to the Ticker_Symbol Total
            Ticker_Total = Ticker_Total + WS.Cells(i, 7).Value
    
            End If
            
        Next i
            
    Next WS
    
End Sub

