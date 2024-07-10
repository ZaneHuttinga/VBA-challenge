Sub stock_stats()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets

        ws.Select

        'Quarterly Changes

        'Insert headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
    
        'Define number of rows for the For loop
        Dim last_row As Long
        last_row = Cells(Rows.Count, 1).End(xlUp).Row

        'Define counter for the rows giving quarterly changes
        Dim i As Integer
        i = 2
    
        'Define start of each set of rows for a given ticker
        Dim stock_first_row As Long
        stock_first_row = 2
    
        'Define end of each set of rows for a given ticker
        Dim stock_last_row As Long
   
        'Loop through rows and pull out desired information
        For j = 2 To last_row
    
            'Define steps when ticker changes
            If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
            
                'Ticker
                Cells(i, 9).Value = Cells(j, 1).Value
            
                'Quarterly change
                Cells(i, 10).Value = Cells(stock_last_row, 6).Value - Cells(stock_first_row, 3).Value
            
                'Percent Change
                Cells(i, 11).Value = Cells(i, 10).Value / Cells(stock_first_row, 3).Value
            
                'Total Volume
                Cells(i, 12).Value = Application.Sum(Range(Cells(stock_first_row, 7), Cells(stock_last_row, 7)))
            
                'Move start of data to the next ticker
                stock_first_row = j + 1
            
                'Move to the next row in the Quarterly Change columns
                i = i + 1
            
            Else
        
                'Move on if ticker hasn't changed
                stock_last_row = j + 1
            
            End If
        
        Next j
    
    
    
        'Conditional Formatting
    
        'Define number of distinct tickers
        Dim last_row_max As Long
        last_row_max = Cells(Rows.Count, 9).End(xlUp).Row
    
        'Loop through distinct tickers
        For l = 2 To last_row_max
    
            If Cells(l, 10).Value < 0 Then
        
                Cells(l, 10).Interior.ColorIndex = 3
            
            ElseIf Cells(l, 10).Value > 0 Then
        
                Cells(l, 10).Interior.ColorIndex = 4
            
            End If
            
            If Cells(l, 11).Value < 0 Then
        
                Cells(l, 11).Interior.ColorIndex = 3
            
            ElseIf Cells(l, 11).Value > 0 Then
        
                Cells(l, 11).Interior.ColorIndex = 4
            
            End If
        
        Next l
    
    
    
        'Greatest Values
    
        'Insert headers
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
    
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
    
        'Insert greatest values
        Cells(2, 17).Value = WorksheetFunction.Max(Range(Cells(2, 11), Cells(last_row_max, 11)))
        Cells(3, 17).Value = WorksheetFunction.Min(Range(Cells(2, 11), Cells(last_row_max, 11)))
        Cells(4, 17).Value = WorksheetFunction.Max(Range(Cells(2, 12), Cells(last_row_max, 12)))
    
        'Insert corresponding tickers
        For m = 2 To last_row_max
    
            'Ticker of greatest % increase
            If Cells(m, 11).Value = Cells(2, 17).Value Then
        
                Cells(2, 16).Value = Cells(m, 9).Value
            
            'Ticker of greatest % decrease
            ElseIf Cells(m, 11).Value = Cells(3, 17).Value Then
        
                Cells(3, 16).Value = Cells(m, 9).Value
            
            'Ticker of greatest total volume
            ElseIf Cells(m, 12).Value = Cells(4, 17).Value Then
        
                Cells(4, 16).Value = Cells(m, 9).Value
            
            End If
        
        Next m
            
    Next ws
    
End Sub