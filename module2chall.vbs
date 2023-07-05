Sub MULTIPLEYEARSTOCK()
  


   Dim Ticker As String
   Dim Total_stock_volume As Double
    Dim ticker_row As Integer
    Dim openprice As Double
    Dim closeprice As Double
    Dim ws As Worksheet
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_total_volume As Double
    
 Total_stock_volume = 0
    For Each ws In ThisWorkbook.Worksheets

 
   ws.Activate
 
   ticker_row = 2
   openprice = ws.Cells(2, 3).Value
    ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly change"
 ws.Cells(1, 11).Value = "percent change"
 ws.Cells(1, 12).Value = "Total stock volume"
 ws.Cells(1, 18).Value = "Ticker"
 ws.Cells(1, 19).Value = "value"
 ws.Cells(3, 17).Value = "Greatest % Increase"
 ws.Cells(4, 17).Value = "Greatest % Decrease"
 ws.Cells(5, 17).Value = "Greatest Total Volume"
 
   Greatest_Increase = 0
 
 Greatest_total_volume = 0
 
   
  
  
   For i = 2 To 759001
   
   Ticker = ws.Cells(i, 1).Value
   Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value


  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   ws.Cells(ticker_row, 12).Value = Total_stock_volume
   ws.Cells(ticker_row, 9).Value = Ticker
   closeprice = ws.Cells(i, 6).Value
   
   Yearly_change = closeprice - openprice
   
   
  ws.Cells(ticker_row, 10).Value = Yearly_change
  
  
  If Yearly_change < 0 Then
ws.Cells(ticker_row, 10).Interior.ColorIndex = 3

Else


ws.Cells(ticker_row, 10).Interior.ColorIndex = 4
End If
  
  
  
  
  
Percent_change = Yearly_change / openprice
If Percent_change > Greatest_Increase Then
Greatest_Increase = Percent_change
ws.Cells(3, 19).Value = Greatest_Increase
ws.Cells(3, 19).NumberFormat = "0.00%"
ws.Cells(3, 18).Value = Ticker
End If



If Percent_change < Greatest_Decrease Then
Greatest_Decrease = Percent_change
ws.Cells(4, 19).Value = Greatest_Decrease
ws.Cells(4, 19).NumberFormat = "0.00%"

ws.Cells(4, 18).Value = Ticker
End If

If Total_stock_volume > Greatest_stock_volume Then
Greatest_stock_volume = Total_stock_volume
ws.Cells(5, 19).Value = Greatest_stock_volume
ws.Cells(5, 18).Value = Ticker


End If




ws.Cells(ticker_row, 11).Value = Percent_change
ws.Cells(ticker_row, 11).NumberFormat = "0.00%"





openprice = ws.Cells(i + 1, 3).Value
ticker_row = ticker_row + 1
  Total_stock_volume = 0
  
End If
Next i
