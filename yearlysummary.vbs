Sub YearSummary():

    For Each ws In Worksheets

        Dim lastrow As Long
        Dim i As Long
        Dim ticker_symbol As String
        Dim yearly_change As Double
        Dim ticker_first_row As Boolean
        Dim year_opening_price As Double
        Dim year_closing_price As Double
        Dim percent_change As Double
        Dim total_stock_volume As Double
        Dim summary_table_row As Long
        Dim greatest_percent_increase As Double
        Dim greatest_percent_decrease As Double
        Dim greatest_total_volume As Double
        Dim lastcolumn As Integer
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ticker_symbol = ""
        yearly_change = 0
        ticker_first_row = True
        year_opening_price = 0
        year_closing_price = 0
        percent_change = 0
        total_stock_volume = 0
        summary_table_row = 2
        greatest_percent_increase = 0
        greatest_percent_decrease = 0
        greatest_total_volume = 0
        lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ws.Cells(1, lastcolumn + 2).Value = "Ticker"
        ws.Cells(1, lastcolumn + 3).Value = "Yearly Change"
        ws.Cells(1, lastcolumn + 4).Value = "Percent Change"
        ws.Cells(1, lastcolumn + 5).Value = "Total Stock Volume"
    
        ws.Cells(1, lastcolumn + 8).Value = "Ticker"
        ws.Cells(1, lastcolumn + 9).Value = "Value"
        
        ws.Cells(2, lastcolumn + 7).Value = "Greatest % Increase"
        ws.Cells(2, lastcolumn + 7).Value = "Greatest % Decrease"
        ws.Cells(2, lastcolumn + 7).Value = "Greatest Total Volume"
    
        For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ticker_symbol = ws.Cells(i, 1).Value
                year_closing_price = ws.Cells(i, 6).Value
                
                yearly_change = year_closing_price - year_opening_price
                
                If Not year_opening_price = 0 Then
                    percent_change = ((year_closing_price - year_opening_price) / year_opening_price)
                Else
                    percent_change = year_opening_price - year_closing_price
                End If
                
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                
                ws.Cells(summary_table_row, lastcolumn + 2).Value = ticker_symbol
                ws.Cells(summary_table_row, lastcolumn + 3).Value = yearly_change
                
                If yearly_change > 0 Then
                    ws.Cells(summary_table_row, lastcolumn + 3).Interior.ColorIndex = 4
                Else
                    ws.Cells(summary_table_row, lastcolumn + 3).Interior.ColorIndex = 3
                End If
                
                ws.Cells(summary_table_row, lastcolumn + 4).Value = percent_change
                ws.Cells(summary_table_row, lastcolumn + 4).NumberFormat = "0.00%"
                ws.Cells(summary_table_row, lastcolumn + 5).Value = total_stock_volume
                
                summary_table_row = summary_table_row + 1
                total_stock_volume = 0
                ticker_first_row = True
                
            Else
                
                If ticker_first_row = True Then
                
                    year_opening_price = ws.Cells(i, 3).Value
                    ticker_first_row = False
                
                End If
                
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                    
            End If
        
        Next i
    
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
        
            If ws.Cells(i, lastcolumn + 4).Value > 0 Then
                If greatest_percent_increase < ws.Cells(i, lastcolumn + 4).Value Then
                    greatest_percent_increase = ws.Cells(i, lastcolumn + 4).Value
                    ticker_symbol = ws.Cells(i, lastcolumn + 2).Value
                    ws.Cells(4, lastcolumn + 9).Value = ticker_symbol
                    ws.Cells(2, lastcolumn + 10).Value = greatest_percent_increase
                End If
            End If
        Next i
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
            If ws.Cells(i, lastcolumn + 4).Value < 0 Then
                If greatest_percent_decrease > ws.Cells(i, lastcolumn + 4).Value Then
                    greatest_percent_decrease = ws.Cells(i, lastcolumn + 4).Value
                    ticker_symbol = ws.Cells(i, lsatcolumn + 2).Value
                    ws.Cells(4, lastcolumn + 9).Value = ticker_symbol
                    ws.Cells(2, lastcolumn + 10).Value = greatest_percent_decrease
                End If
            End If
        Next i
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
        
            If greatest_total_volume < ws.Cells(i, lastcolumn + 5).Value Then
                greatest_total_volume = ws.Cells(i, lastcolumn + 5).Value
                ticker_symbol = ws.Cells(i, lsatcolumn + 2).Value
                ws.Cells(4, lastcolumn + 9).Value = ticker_symbol
                ws.Cells(4, lastcolumn + 10).Value = greatest_total_volume
            End If
                
        Next i
            
            
    Next ws

End Sub
