Attribute VB_Name = "Module1"
Sub stockticker()
    'set initial variable type for ticker symbol
    Dim ticker As String
    
    'loop through all worksheets
    For Each ws In Worksheets
    
    'set variable for holding the, open price, close price and YoY change
    Dim open_price As Double
    Dim close_price As Double
    Dim change_year As Double
    Dim stock_vol As String

        
    'Define initial values
    open_price = 0
    close_price = 0
    change_year = 0
    stock_vol = 0
    perc_delta = 0
    
    'give headeers for the table
    ws.Range("J1").Value = "Stock Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "% Delta"
    ws.Range("M1").Value = "Total Volume"
    ws.Range("N1").Value = "open"
    ws.Range("O1").Value = "close"
    
    
    'Location where to start looking stock information
    Dim stock_table_row As Integer
    stock_table_row = 2
    
    'loop thru the table for stock tickers
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'check if stock ticker is still same else move to the next
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
            'set the ticker symbol
           ticker = ws.Cells(i, 1).Value
            
            
            'retrieve the closing price
            ws.Cells(2, 14).Value = ws.Cells(2, 3).Value
            close_price = ws.Cells(i, 6).Value
            open_price = ws.Cells(i + 1, 3).Value
           
            
            'get the total stock volume
            stock_vol = stock_vol + ws.Cells(i, 7).Value
  
            
            'compute the price change
            change_year = close_price - ws.Cells(i - 1, 3).Value
            
            'Compute % change
            perc_delta = change_year / ws.Cells(i - 1, 3).Value
            
            'Print the ticker symbol in the summary table
            ws.Range("J" & stock_table_row).Value = ticker
            
            'Print the YoY change in price
            ws.Range("K" & stock_table_row).Value = change_year
            
            'Print the stock volume to the summary table
            ws.Range("L" & stock_table_row).Value = FormatPercent(perc_delta, 0)
            
            'Print the Closing Price
            ws.Range("M" & stock_table_row).Value = stock_vol
            
            'Print the Closing Price
            ws.Range("N" & 1 + stock_table_row).Value = open_price
            
            'Print the Closing Price
            ws.Range("O" & stock_table_row).Value = close_price
                        
            'Add one to the summary table
            stock_table_row = stock_table_row + 1
            
            ' reset the change year
            change_year = 0
            
            'reset the percentage change
            perc_delta = 0
            
            'reset the stock volume
            stock_vol = 0
            
            'if cell immediately follwing row is the same brand
            Else
                
                'add to the stock volume
                stock_vol = stock_vol + ws.Cells(i, 7).Value
            End If
            
        Next i
        
        ' Set the Font color to Red
        For i = 2 To 91
    
        If ws.Cells(i, 11).Value < 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 3
        
        ElseIf ws.Cells(i, 11).Value > 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
            
        Else
            ws.Cells(i, 11).Interior.ColorIndex = 6
        
        
        End If
        
        Next i
        
        Next ws
        
End Sub



