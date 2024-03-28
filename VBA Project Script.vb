'arranges tables on worksheets
Sub testing()
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
        'assigning headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        'initializing variables to store info
        Dim ticker As String
        Dim sumrow, ticker_count As Integer
        Dim volume, open_val, close_val, yearly_change, percent_change As Double
        
        sumrow = 2
        ticker_count = 0
        volume = 0
        open_val = 0
        close_val = 0
        
        'calculating last row
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'for loop to loop through every row
        For i = 2 To lastrow
        
            'checks if the next cell has a different ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'if so assign the ticker to variable "ticker" and calculate final volume
                'and then place ticker and volume in new location
                ticker = Cells(i, 1).Value
                volume = volume + Cells(i, 7).Value
                Range("I" & sumrow).Value = ticker
                Range("L" & sumrow).Value = volume
            
                'assigns closing value then calculates and places yearly change and percent change
                close_val = Cells(i, 6).Value
                yearly_change = close_val - open_val
                percent_change = (yearly_change / open_val)
                
                'creation and designs yearly and percent change columns
                Range("K" & sumrow).Value = percent_change
                Range("K" & sumrow).NumberFormat = "0.00%"
                Range("J" & sumrow).Value = yearly_change
                If Range("J" & sumrow).Value > 0 Then
                    Range("J" & sumrow).Interior.ColorIndex = 4
                Else
                    Range("J" & sumrow).Interior.ColorIndex = 3
                End If
                
                
                'increase sumrow to store next values in next row and return volume to 0 to start counting for new ticker
                'ticker count is set to 0 so we can find the opening value of next ticker
                sumrow = sumrow + 1
                volume = 0
                ticker_count = 0
                
            'if tickers are the same
            Else
                'add the volumes together to keep a running count
                volume = volume + Cells(i, 7).Value
                
                'this finds opening value for new ticker
                ticker_count = ticker_count + 1
                If ticker_count = 1 Then
                    open_val = Cells(i, 3).Value
                End If
            End If
        
        Next i
        
        'assigns headers and titles
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
            
        
        'creates table with greatest/least percent change and greatest volume
        
        
        'initializing variables
        Dim highticker, lowticker, volticker As String
        highticker = Range("I2").Value
        lowticker = Range("I2").Value
        volticker = Range("I2").Value
        Dim rnglastrow As Integer
        rnglastrow = Cells(Rows.Count, "I").End(xlUp).Row
        Dim high_percentage, low_percentage, high_volume As Double
        high_percentage = 0
        low_percentage = 0
        high_volume = 0
        
        For i = 2 To rnglastrow
                'continuously stores new highest percent increase
                If Cells(i, 11).Value > high_percentage Then
                    high_percentage = Cells(i, 11).Value
                    highticker = Cells(i, 9).Value
                End If
                'continuously stores new least percent decrease
                If Cells(i, 11) < low_percentage Then
                    low_percentage = Cells(i, 11).Value
                    lowticker = Cells(i, 9).Value
                End If
                'continuously stores new highest volume
                If Cells(i, 12).Value > high_volume Then
                    high_volume = Cells(i, 12).Value
                    volticker = Cells(i, 9).Value
                End If
        Next i
        
        'places values
        Range("P2").Value = highticker
        Range("P3").Value = lowticker
        Range("P4").Value = volticker
        Range("Q2").Value = high_percentage
        Range("Q3").Value = low_percentage
        Range("Q2:Q3").NumberFormat = "0.00%"
        Range("Q4").Value = high_volume
                
    Next ws
    
End Sub