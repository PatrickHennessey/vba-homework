Sub TickerMediumMode()

      ' Set an initial variable for holding the ticker symbol, ticker volume, summary table, and last row.
      Dim Ticker_Symbol As String
      Dim Yearly_Change As Double
      Dim Opening_Price As Double
      Dim Closing_Price As Double
      Dim Percent_Change As Double
      Dim Ticker_Volume As Double
      Dim Summary_Table_Row As Integer
      Dim LastRow As Long
      Dim ws As Worksheet
      
      For Each ws In Worksheets
      
        ' Set initial vaules for the variables used.
        Ticker_Total = 0
        Summary_Table_Row = 2
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
    
        ' Name New Columns
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        
        ' Grab Opening_Price before starting to loop through the rows
        Opening_Price = ws.Cells(2, "C").Value
        
        ' Loop through all the unique ticker symbols until the last row.
        For i = 2 To LastRow
            
            ' Check if we are still within the same ticker symbol, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ' Set the Ticker_Symbol name
            Ticker_Symbol = ws.Cells(i, "A").Value
    
            ' Add to the Ticker Total
            Ticker_Total = Ticker_Total + ws.Cells(i, "G").Value
    
            ' Print the Ticker Symbol in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol

            ' Grab the Closing_Price for Calculations
            Closing_Price = ws.Cells(i, "F").Value
            
            ' Add Yearly_Change
            Yearly_Change = Closing_Price - Opening_Price
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        
            ' Add Percent_Change, dividing by 0 resultes in an error, so we have to make sure it's higher than 0.
            If (Opening_Price = 0 And Closing_Price = 0) Then
                Percent_Change = 0
            ElseIf (Opening_Price = 0 And Closing_Price <> 0) Then
                Percent_Change = 1
            Else
                Percent_Change = Yearly_Change / Opening_Price
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            End If
            
            ' Print the Ticker_Symbol Total to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
    
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the Opening_Price & Ticker_Total
            Opening_Price = ws.Cells(i + 1, "C")
            Ticker_Total = 0
    
            ' If the cell immediately following a row is the same ticker
            Else
    
            ' Reset the Ticker_Total
            
            Ticker_Total = Ticker_Total + ws.Cells(i, "G").Value
    
            End If
    
        Next i
            
        ' Set the Cell Colors
        For j = 2 To LastRow
            If (ws.Cells(j, "J").Value > 0 Or Cells(j, "J").Value = 0) Then
                ws.Cells(j, "J").Interior.ColorIndex = 10
            ElseIf ws.Cells(j, "J").Value < 0 Then
                ws.Cells(j, "J").Interior.ColorIndex = 3
            End If
        Next j

        ' Autofit to display data
        Columns("A:L").AutoFit
    
    Next ws
    
End Sub
