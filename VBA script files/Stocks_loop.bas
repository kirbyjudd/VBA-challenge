Attribute VB_Name = "Stocks_loop"
Sub Stocks_loop():

'Define worksheet variable
Dim ws As Worksheet
For Each ws In Worksheets

'Variables-------------------------------------------
Dim ticker_symbol As String
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock As Double
Dim start_row As Long
start_row = 2
Dim lastRow As Long
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Double
Dim increase_stock As String
Dim decrease_stock As String
Dim volume_stock As String
greatest_increase = 0
greatest_decrease = 0
volume_stock = 0
'----------------------------------------------------

    'Create headers in given cell ranges
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker Symbol"
    ws.Range("Q1").Value = "Value"
    
    'Find the last row of each worksheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    'For Loop
    For i = 2 To lastRow
    
        'If the ticker symbol does not equal the ticker symbol before it: retrieve and print the ticker symbol value and the opening price
        If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
            ticker_symbol = ws.Cells(i, 1).Value
            ws.Cells(start_row, 9).Value = ticker_symbol
            open_price = ws.Cells(i, 3).Value
            'Starts counting total stock volume
            total_stock = 0
            total_stock = ws.Cells(i, 7).Value
    
        'If the ticker symbol does not equal the ticker symbol after it: retrieve, calculate, and print the closing price and yearly change values
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            close_price = ws.Cells(i, 6).Value
            yearly_change = (close_price - open_price)
            ws.Cells(start_row, 10).Value = yearly_change
            
                'Conditional formatting for yearly change: If yearly change is greater or equal to 0 then green cell, else red cell
                If yearly_change >= 0 Then
                    ws.Cells(start_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(start_row, 10).Interior.ColorIndex = 3
                End If
                
            'Calculate and print the percent change value
            percent_change = ((close_price - open_price) / open_price)
            ws.Cells(start_row, 11).Value = percent_change
            
                'Conditional formatting for yearly change: If yearly change is greater or equal to 0 then green cell, else red cell
                If percent_change >= 0 Then
                    ws.Cells(start_row, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(start_row, 11).Interior.ColorIndex = 3
                End If
            
            'Add stock volume values to get total stock volume and print
            total_stock = total_stock + ws.Cells(i, 7).Value
            ws.Cells(start_row, 12).Value = total_stock
            
                'Greatest % increase If statement to find the greatest increase percentange change
                If greatest_increase < percent_change Then
                    greatest_increase = percent_change
                    increase_stock = ticker_symbol
                Else
                End If
                                
                'Greatest % decrease If statment to find the greatest decrease percentage change
                If greatest_decrease > percent_change Then
                    greatest_decrease = percent_change
                    decrease_stock = ticker_symbol
                Else
                End If
             
                'Greatest total volume If statement to find the greatest total volume value for a stock
                If greatest_volume < total_stock Then
                    greatest_volume = total_stock
                    volume_stock = ticker_symbol
                Else
                End If
                
                'Add 1 to start row
                start_row = start_row + 1
            
            'If ticker symbol equals the previous symbol then add it to the total stock volume value
            ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
                total_stock = total_stock + ws.Cells(i, 7).Value
                
            End If
        
        Next i
        
        'Percentage formatting and input values for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        ws.Columns("K:K").NumberFormat = "0.00%"
        
        ws.Cells(2, 16).Value = increase_stock
        ws.Cells(2, 17).Value = greatest_increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
    
    
        ws.Cells(3, 16).Value = decrease_stock
        ws.Cells(3, 17).Value = greatest_decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
    
    
        ws.Cells(4, 16).Value = volume_stock
        ws.Cells(4, 17).Value = greatest_volume
        
        'Reset Values to 0 after input in cell
        greatest_decrease = 0
        greatest_increase = 0
        greatest_volume = 0
        
        'Autofit column widths
        ws.Columns("I:Q").AutoFit
        
    Next ws

End Sub



