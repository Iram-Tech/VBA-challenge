Sub check_stocks():
    Dim row_count As Long
    Dim log_ticker As String
    Dim log_opening_price As Double
    
    ' LOOP THROUGH EACH WORKSHEET
    For Each ws In Worksheets
        Dim first_value As Long
        Dim opening_price As Double
        Dim closing_price As Double
        Dim total_stock_count As Long
        total_stock_count = 0
        first_value = 2
        opening_price = Cells(first_value, 3).Value
        
        ' Find total number of rows in each worksheet
        ' Add field headers
        ' Ticker Sample   Yearly Change   Percent Change  Total Stock Volume
        row_count = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("i1").Value = "Ticker"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"
        
        For i = 2 To row_count
            log_ticker = Cells(i, 1).Value
            total_stock_count = total_stock_count + Cells(i, 7)
            ' IF THE TICKER CHANGES THEN SUBTRACT THE
            ' CLOSING PRICE AT THE END OF THE YEAR
            ' WITH THE OPENING PRICE AT THE BEGINNING OF THE YEAR
            If log_ticker <> Cells(i + 1, 1).Value Then
                Cells(first_value, 9).Value = log_ticker
                Cells(first_value, 10).Value = Cells(i, 6).Value - opening_price
                If Cells(i, 6).Value - opening_price < 0 Then
                    Cells(first_value, 10).Interior.ColorIndex = 3
                Else
                    Cells(first_value, 10).Interior.ColorIndex = 4
                End If
                
                first_value = first_value + 1
                opening_price = Cells(i + 1, 3).Value
                log_ticker = Cells(i + 1, 1).Value
                Cells(first_value, 11).Value = Str(Round(((Cells(i, 6).Value - opening_price) / opening_price) * 100, 2)) + "%"
                Cells(first_value, 12).Value = total_stock_count
                
            End If
            total_stock_count = 0
        Next i
    Next ws
    
End Sub
Sub check_stocks():
    Dim row_count As Long
    Dim log_ticker As String
    Dim log_opening_price As Double
    
    ' LOOP THROUGH EACH WORKSHEET
    For Each ws In Worksheets
        Dim first_value As Long
        Dim opening_price As Double
        Dim closing_price As Double
        Dim total_stock_count As Long
        total_stock_count = 0
        first_value = 2
        opening_price = Cells(first_value, 3).Value
        
        ' Find total number of rows in each worksheet
        ' Add field headers
        ' Ticker Sample   Yearly Change   Percent Change  Total Stock Volume
        row_count = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("i1").Value = "Ticker"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"
        
        For i = 2 To row_count
            log_ticker = Cells(i, 1).Value
            total_stock_count = total_stock_count + Cells(i, 7)
            ' IF THE TICKER CHANGES THEN SUBTRACT THE
            ' CLOSING PRICE AT THE END OF THE YEAR
            ' WITH THE OPENING PRICE AT THE BEGINNING OF THE YEAR
            If log_ticker <> Cells(i + 1, 1).Value Then
                Cells(first_value, 9).Value = log_ticker
                Cells(first_value, 10).Value = Cells(i, 6).Value - opening_price
                If Cells(i, 6).Value - opening_price < 0 Then
                    Cells(first_value, 10).Interior.ColorIndex = 3
                Else
                    Cells(first_value, 10).Interior.ColorIndex = 4
                End If
                
                first_value = first_value + 1
                opening_price = Cells(i + 1, 3).Value
                log_ticker = Cells(i + 1, 1).Value
                Cells(first_value, 11).Value = Str(Round(((Cells(i, 6).Value - opening_price) / opening_price) * 100, 2)) + "%"
                Cells(first_value, 12).Value = total_stock_count
                
            End If
            total_stock_count = 0
        Next i
    Next ws
    
End Sub
