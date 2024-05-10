# Module2_Challenge
Sub Stocks():

    
    For Each ws In Worksheets
    
        
        Dim ticker As String
        
        Dim quarterly_change As Double
        
        Dim percentage_change As Double
        
        Dim total_stock_volume As Variant
        
        Dim tickerRowNumber As Integer
        
        Dim opening_price As Double
        
        Dim closing_price As Double
        
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        tickerRowNumber = 2
        
        total_stock_volume = 0
        
        
        ws.Cells(1, 8).Value = "Ticker"
        
        ws.Cells(1, 9).Value = "Quarterly Change ($)"
        
        ws.Cells(1, 10).Value = "Percent Change"
        
        ws.Cells(1, 11).Value = "Total Stock Volume"
 
        
        For i = 2 To LastRow
        
        
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                opening_price = ws.Cells(i, 3)
                
            End If
            
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ticker = ws.Cells(i, 1).Value
                
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                
                closing_price = ws.Cells(i, 6).Value
                
                quarterly_change = closing_price - opening_price
                
                percent_change = quarterly_change / opening_price
                
                ws.Cells(tickerRowNumber, 8).Value = ticker
                
                ws.Cells(tickerRowNumber, 11).Value = total_stock_volume
                
                ws.Cells(tickerRowNumber, 9) = quarterly_change
                
                ws.Cells(tickerRowNumber, 10) = percent_change
                
                tickerRowNumber = tickerRowNumber + 1
            
                total_stock_volume = 0
                
            Else
            
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                
            End If
            
    
            ws.Cells(i, 10).NumberFormat = "0.00%"
            
            ws.Cells(i, 9).NumberFormat = "0.00"
            
            ws.Cells(i, 11).NumberFormat = "0"
            
          
        Next i
        
    
        
        For i = 2 To 1501
        
        
            quarterly_change = ws.Cells(i, 9).Value
            
            total_stock_volume = ws.Cells(i, 11).Value
            
    
            If quarterly_change > 0 Then
            
                ws.Cells(i, 9).Interior.ColorIndex = 4
    
                
            ElseIf quarterly_change < 0 Then
            
                ws.Cells(i, 9).Interior.ColorIndex = 3
                
                
            End If
            

        
        Dim max_percent_increase As Double
           
        Dim max_percent_decrease As Double
        
        Dim max_volume As Variant
        
        
        max_percent_increase = 0
        
        max_percent_decrease = 0
        
        max_volume = 0
        
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ws.Cells(1, 15).Value = "Ticker"
        
        ws.Cells(1, 16).Value = "Value"
    
 
        
        If ws.Cells(i, 10).Value > ws.Cells(i + 1, 10).Value Then
        
            max_percent_increase = ws.Cells(i, 10)
        
            ws.Cells(2, 16).Value = max_percent_increase
            
            ws.Cells(2, 15).Value = ws.Cells(i, 8)
            
        
            
        End If
        
            
        If ws.Cells(i, 10).Value < ws.Cells(i + 1, 10).Value Then
        
            max_percent_decrease = ws.Cells(i, 10)
        
            ws.Cells(3, 16).Value = max_percent_decrease
            
            ws.Cells(3, 15).Value = ws.Cells(i, 8)
            
        
               
        End If
        
        
        If ws.Cells(i, 11).Value > ws.Cells(i + 1, 11) Then
        
            max_volume = ws.Cells(i, 11)
            
            ws.Cells(4, 16).Value = max_volume
            
            ws.Cells(4, 15).Value = ws.Cells(i, 8)
    
              
        End If
        
         
        Next i
        
       
        
        For i = 2 To 3
        
            ws.Cells(i, 16).NumberFormat = "0.00%"
            
        Next i
        
       
    Next ws
             
End Sub



Module 2 Challenge
