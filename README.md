Attribute VB_Name = "Module1"
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
            ws.Range("J1").Value = "Yearly Change
            ws.Range("K1").Value = "Percentage Ch
            ws.Range("L1").Value = "Total Stock V
          
            ws.Cells(3, 15).Value = "Greatest % I
            ws.Cells(4, 15).Value = "Greatest % D
            ws.Cells(5, 15).Value = "Greatest Tot
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
    
        last_row = ws.Cells(Rows.Count, 1).End(xl
        
    For i = 2 To last_row

       If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Th
            ws.Range("I" & summary_Row) = ws.Cell
            ticker_total = ticker_total + ws.Cell
            ws.Range("L" & summary_Row) = ticker_
            ticker_total = 0
            summary_Row = summary_Row + 1
        Else
            ticker_total = ticker_total + ws.Cell
            
        End If
        
          If Cells(i, 1).Value <> Cells(i - 1, 1)
            year_open = Cells(i, 3).Value
            
        End If
        
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) T
            year_close = ws.Cells(i, 6).Value
            ws.Range("I" & summary_Row) = ws.Cell
            yearly_change = year_change + (year_c
            ws.Range("J" & summary_Row) = yearly_
            
        End If
            
        If yearly_change <> 0 And year_open <> 0 
            percetage_change = percentage_change 
            ws.Range("K" & summary_Row) = percent
            ws.Range("K" & summary_Row).NumberFor
        Else
            ws.Range("K" & summary_Row) = percent
            ws.Range("K" & summary_Row).NumberFor
        
        End If
        
        If yearly_change > 0 Then
            ws.Range("J" & summary_Row).Interior.
        Else
            ws.Range("J" & summary_Row).Interior.
        
        End If
        
       
        
    Next i



'Challenge analysis


    Dim great_ticker As String
    Dim bad_ticker As String
    Dim volume_ticker As String
    Dim great_value As Double
    Dim bad_value As Double
    Dim volume_value As Variant
    
        great_value = ws.Cells(3, 16)
        bad_value = ws.Cells(4, 16)
        volume_value = ws.Cells(5, 16)
        last_row2 = ws.Cells(Rows.Count, 1).End(x
        
    For i = 2 To last_row2
    
        If ws.Cells(i, 11).Value > great_value Th
            great_value = ws.Cells(a, 11).Value
            great_tick = ws.Cells(a, 9).Value
        End If
        
        If ws.Cells(i, 11).Value < volume_value T
            bad_value = es.Cells(i, 9).Value
            bad_ticker = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 12).Value > volume_value T
            volume_value = ws.Cells(i, 12).Value
            ticker_volume = ws.calls(i, 9).Value
        End If
        
    ws.Range("o2") = great_ticker
  
Next i

Next ws

End Sub



