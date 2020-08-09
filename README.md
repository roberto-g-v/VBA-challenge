Sub stock_data()


    Dim ticker_name As String
    Dim ticker_total As Variant
    Dim yearly_change As Variant
    Dim year_close As Variant
    Dim year_open As Variant
    Dim summary_Row As String
    Dim percentage_change As Variant
        
    For Each ws In Worksheets

            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percentage Change"
            ws.Range("L1").Value = "Total Stock Volume"
          
            ws.Cells(3, 15).Value = "Greatest % Increase"
            ws.Cells(4, 15).Value = "Greatest % Deacrease"
            ws.Cells(5, 15).Value = "Greatest Total Value"
            ws.Cells(2, 16).Value = "Ticker"
            ws.Cells(2, 17).Value = "Value"
            
        ticker_total = 0
        yearly_change = 0
        total_stock = 0
        year_close = 0
        year_open = 0
        summary_Row = 2
        percentage_change = 0
        ticker_total = 0
        summary_Row = 2
    
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To last_row

       If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            ws.Range("I" & summary_Row) = ws.Cells(i, 1)
            ticker_total = ticker_total + ws.Cells(i, 7)
            ws.Range("L" & summary_Row) = ticker_total
            ticker_total = 0
            summary_Row = summary_Row + 1
        Else
            ticker_total = ticker_total + ws.Cells(i, 7)
            
        End If
        
          If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            year_open = Cells(i, 3).Value
            
        End If
        
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            year_close = ws.Cells(i, 6).Value
            ws.Range("I" & summary_Row) = ws.Cells(i, 1)
            yearly_change = year_change + (year_close - year_open)
            ws.Range("J" & summary_Row) = yearly_change
            
        End If
        
        If yearly_change > 0 Then
            ws.Range("J" & summary_Row).Interior.ColorIndex = 4
        Else
            ws.Range("J" & summary_Row).Interior.ColorIndex = 3
        
        End If
        
        If yearly_change <> 0 And year_open <> 0 Then
            percetage_change = percentage_change + (yearly_change / year_open)
            ws.Range("K" & summary_Row) = percentage_change
            ws.Range("K" & summary_Row).NumberFormat = "0.00%"
        Else
            ws.Range("K" & summary_Row) = percentage_change
            ws.Range("K" & summary_Row).NumberFormat = "0.00%"
        
        End If
        
    Next i



'Challenge analysis


    Dim great_ticker As String
    Dim bad_ticker As String
    Dim volume_ticker As String
    Dim great_value As Double
    Dim bad_value As Variant
    Dim volume_value As Variant
    
        great_value = ws.Cells(3, 16)
        bad_value = ws.Cells(4, 16)
        volume_value = ws.Cells(5, 16)
        last_row2 = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To last_row2
    
        If ws.Cells(i, 11).Value > great_value Then
            great_value = ws.Cells(a, 11).Value
            great_tick = ws.Cells(a, 9).Value
        End If
        
        If ws.Cells(i, 11).Value < volume_value Then
            bad_value = ws.Cells(i, 9).Value
            bad_ticker = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 12).Value > volume_value Then
            volume_value = ws.Cells(i, 12).Value
            ticker_volume = ws.Cells(i, 9).Value
        End If
        
    ws.Range("o2") = great_ticker
    ws.Range("o3") = bad_ticker
    ws.Range("o4") = volume_ticker
    ws.Range("p2") = great_value
    ws.Range("p3") = bad_value
    ws.Range("p4") = volume_value
    ws.Range("p2") = NumberFormat = "0.00%"
    ws.Range("p3") = NumberFormat = "0.00%"
    
  
    Next i

Next ws

End Sub
