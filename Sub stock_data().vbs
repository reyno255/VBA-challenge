Sub stock_data()

Dim Sheet_Name As String
Dim end_row As Long
Dim Summary_Count As Long
Dim ws As Worksheet

Dim ticker As String
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Double

Dim inc_ticker As String
Dim dec_ticker As String
Dim vol_ticker As String

'find last row
For Each ws In Worksheets
    With ws
        end_row = .Cells(.Rows.Count, "C").End(xlUp).Row
    End With

'set initial variables
    Summary_Count = 2
    total_volume = 0
    open_price = 0
    final_price = 0
    
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    'ticker = ""
    
    'set up summary headers
    ws.Range("I1") = ("Ticker")
    ws.Range("J1") = ("Yearly Change")
    ws.Range("K1") = ("Percent Change")
    ws.Range("L1") = ("Total Stock Volume")
    
    'parameters to identify start row and end row for each ticker + running summation of volume
    For stock_Count = 2 To end_row
        total_volume = total_volume + ws.Cells(stock_Count, 7)
        
        If ws.Cells(stock_Count, 1) <> ws.Cells(stock_Count - 1, 1) Then
            open_price = ws.Cells(stock_Count, 3)
        
        ElseIf Cells(stock_Count, 1) <> Cells(stock_Count + 1, 1) Then
            final_price = ws.Cells(stock_Count, 6)
            ws.Cells(Summary_Count, 9) = ws.Cells(stock_Count, 1)
        
            If final_price > 0 And open_price > 0 Then
                'MsgBox (VarType(final_price))
                'MsgBox (VarType(open_price))
                ws.Cells(Summary_Count, 10) = Format(final_price - open_price, "#,##0.00")
                ws.Cells(Summary_Count, 11) = Format((final_price / open_price) - 1, "Percent")
                ws.Cells(Summary_Count, 12) = Format(total_volume, "#,##0")
                If ((final_price / open_price) - 1) > greatest_increase Then
                    greatest_increase = ws.Cells(Summary_Count, 11)
                    inc_ticker = ws.Cells(Summary_Count, 9)
                ElseIf ws.Cells(Summary_Count, 11) < greatest_decrease Then
                    greatest_decrease = ws.Cells(Summary_Count, 11)
                    dec_ticker = ws.Cells(Summary_Count, 9)
                End If
                
            'parameters to handle scenarios where the prevailing stock has a zero price at the opening or final observation date
            ElseIf final_price = 0 And open_price > 0 Then
                ws.Cells(Summary_Count, 10) = Format(final_price - open_price, "#,##0.00")
                ws.Cells(Summary_Count, 11) = Format(-1, "Percent")
                ws.Cells(Summary_Count, 12) = Format(total_volume, "#,##0")
            
            ElseIf final_price > 0 And open_price = 0 Then
                ws.Cells(Summary_Count, 10) = Format(final_price - open_price, "#,##0.00")
                ws.Cells(Summary_Count, 11) = Format(1, "Percent")
                ws.Cells(Summary_Count, 12) = Format(total_volume, "#,##0")
            
            ElseIf final_price = 0 And open_price = 0 Then
                ws.Cells(Summary_Count, 10) = Format(final_price - open_price, "#,##0.00")
                ws.Cells(Summary_Count, 11) = Format(0, "Percent")
                ws.Cells(Summary_Count, 12) = Format(total_volume, "#,##0")
            
            End If
            
            'conditional formatting
            If (final_price - open_price) >= 0 Then
                ws.Cells(Summary_Count, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(Summary_Count, 10).Interior.ColorIndex = 3
            End If
        If total_volume > greatest_volume Then
            greatest_volume = total_volume
            vol_ticker = ws.Cells(Summary_Count, 9)
        End If
        
        total_volume = 0
        Summary_Count = Summary_Count + 1
        
        End If
        
    Next stock_Count
       
    'Bonus:  Find Max, Min, Max annualized returns by year
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Volume"
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    
    ws.Range("O2") = inc_ticker
    ws.Range("P2") = Format(greatest_increase, "0.0000")
    ws.Range("O3") = dec_ticker
    ws.Range("P3") = Format(greatest_decrease, "0.0000")
    ws.Range("O4") = vol_ticker
    ws.Range("P4") = Format(greatest_volume, "#,##0")
    
    'general formatting
    ws.Range("A1:P1").Font.Bold = True
    ws.Range("N2:N4").Font.Bold = True
    ws.Range("A:P").HorizontalAlignment = xlCenter
    ws.Range("A:P").ColumnWidth = 18
    
Next ws

End Sub

Sub clear_cells()

Application.ScreenUpdating = False

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Range("I:P").Clear
    
    ws.Range("I1") = ("Ticker")
    ws.Range("J1") = ("Yearly Change")
    ws.Range("K1") = ("Percent Change")
    ws.Range("L1") = ("Total Stock Volume")
    
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Volume"
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    
    ws.Range("A1:P1").Font.Bold = True
    ws.Range("N2:N4").Font.Bold = True
    ws.Range("A:P").HorizontalAlignment = xlCenter
    ws.Range("A:P").ColumnWidth = 18

Next ws

Application.ScreenUpdating = True

End Sub


