
Sub Testing()

'Looping through all the sheets
For Each ws In Worksheets

'Creating new column headers
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"
  
 'Creating variables to store total volume, opening and closing price of stock
  Dim Ticker_name As String
  
  Dim total_stock_volume As Double
  total_stock_volume = 0
  
  Dim Summary_table_row As Integer
  Summary_table_row = 2
  
  Dim closing_price As Double
  
  Dim opening_price As Double
  opening_price = Cells(2, 3).Value
  
  last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  'Looping through all the stock values to find total stock volume, Yearly change in opening and closing of stock price
  For i = 2 To last_row
  
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      Ticker_name = ws.Cells(i, 1).Value
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
      closing_price = ws.Cells(i, 6).Value
      
    
      ws.Range("I" & Summary_table_row).Value = Ticker_name
      ws.Range("L" & Summary_table_row).Value = total_stock_volume
      ws.Range("J" & Summary_table_row).Value = closing_price - opening_price
      
      'Finding the percentage change of opening price at the beginning of year to the closing price at end of the year
      
      If opening_price = 0 Then
        percentage_change = 0
      Else
        percentage_change = (closing_price - opening_price) / opening_price
      ws.Range("K" & Summary_table_row).Value = percentage_change
      ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"
      End If
      
      
      Summary_table_row = Summary_table_row + 1
      
      'Resetting the total stock volume
      total_stock_volume = 0
      opening_price = ws.Cells(i + 1, 3).Value
      
    Else
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
      
    End If
    
  Next i
  'Creating headers to find the greatest % increase, % decrease, total volume and its corresponding ticker symbol
  ws.Range("O4").Value = "Greatest % Increase"
  ws.Range("O5").Value = "Greatest % Decrease"
  ws.Range("O6").Value = "Greatest Total Volume"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  
  last_row_summarytable = ws.Cells(Rows.Count, 9).End(xlUp).Row
  greatest_increase = 0
  greatest_decrease = 0
  greatest_totvolume = 0

  For i = 2 To last_row_summarytable
  
  'Conditional formatting to highlight positive change in green and negative change in red
    If ws.Cells(i, 10) >= 0 Then
      ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
      ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
    
    If ws.Cells(i, 11).Value > greatest_increase Then
      greatest_increase = ws.Cells(i, 11).Value
      ws.Range("Q4").Value = greatest_increase
      ws.Range("Q4").NumberFormat = "0.00%"
      ws.Range("P4").Value = ws.Cells(i, 9).Value
    End If
    
    If ws.Cells(i, 11).Value < greatest_decrease Then
      greatest_decrease = ws.Cells(i, 11).Value
      ws.Range("Q5").Value = greatest_decrease
      ws.Range("Q5").NumberFormat = "0.00%"
      ws.Range("P5").Value = ws.Cells(i, 9).Value
    End If
    
    If ws.Cells(i, 12).Value > greatest_totvolume Then
      greatest_totvolume = ws.Cells(i, 12).Value
      ws.Range("Q6").Value = greatest_totvolume
      ws.Range("P6").Value = ws.Cells(i, 9).Value
    End If
    
  Next i
ws.Columns("A:Q").AutoFit
  
Next ws
    
  
      
End Sub


