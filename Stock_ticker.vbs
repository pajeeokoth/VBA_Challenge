Sub stockticker()
    'set initial variable type for ticker symbol
    Dim ticker As String
    
    
    'set variable for holding the, open price, close price and YoY change
    Dim open_price As Double
    Dim close_price As Double
    Dim change_year As Double
    Dim stock_vol As String
    Dim perc_delta As Double

        
    'Define initial values
    open_price = 0
    close_price = 0
    change_year = 0
    stock_vol = 0
    perc_delta = 0
    
    'give headeers for the table
    Range("J1").Value = "Stock Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "% Delta"
    Range("M1").Value = "Total Volume"
    
    'Location where to start looking stock information
    Dim stock_table_row As Integer
    stock_table_row = 2
    
    'loop thru the table for stock tickers
    For i = 2 To Sheets("2018").Cells(Rows.Count, 1).End(xlUp).Row
        
        'check if stock ticker is still same else move to the next
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    
            'set the ticker symbol
           ticker = Cells(i, 1).Value
            
            If Cells(i, 2).Value = 20180102 Then
                open_price = Cells(i, 3).Value
                
            Else
                open_price = Cells(i + 1, 3).Value
                
            End If
          
            'retrieve the closing price
            If Cells(i, 2).Value = 20181231 Then
                close_price = Cells(i, 6).Value
                
            Else
                close_price = Cells(i + 1, 6).Value
                
            End If
                             
            'get the total stock volume
            stock_vol = stock_vol + Cells(i, 7).Value
            
            'compute the price change
            change_year = close_price - open_price
            
            'Compute % change
            perc_delta = change_year / open_price
            
            'Print the ticker symbol in the summary table
            Range("J" & stock_table_row).Value = ticker
            
            'Print the YoY change in price
            Range("K" & stock_table_row).Value = change_year
            
            'Print the stock volume to the summary table
            Range("L" & stock_table_row).Value = FormatPercent(perc_delta, 0)
            
            'Print the Closing Price
            Range("M" & stock_table_row).Value = stock_vol

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
                stock_vol = stock_vol + Cells(i, 7).Value
            End If
            
        Next i
        
        ' Set the Font color to Red
        For i = 2 To 3001
    
        If Cells(i, 11).Value < 0 Then
            Cells(i, 11).Interior.ColorIndex = 3
        
        ElseIf Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
            
        Else
            Cells(i, 11).Interior.ColorIndex = 6
            
        Exit For
        
        End If
        
        Next i
        
End Sub