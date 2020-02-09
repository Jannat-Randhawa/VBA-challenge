# VBA-challenge
Sub Analysis()

    ' Set a variable for ws
    
    Dim ws As Worksheet

    'create a loop
    
    For Each ws In Worksheets

        'Create labels for columns
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Set variable to hold the ticker symbol
        Dim ticker_symbol As String

        'Set variable
        
        Dim total_vol As Double
        total_vol = 0
        Dim rowcount As Long
        rowcount = 2
        Dim year_open As Double
        year_open = 0
        Dim year_close As Double
        year_close = 0
        Dim year_change As Double
        year_change = 0
        Dim percent_change As Double
        percent_change = 0
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Create a loop for the tickers
        
        For i = 2 To lastrow
            
            'Create a conditional to get the year open price
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                year_open = ws.Cells(i, 3).Value

            End If

            'Total stock volume for the year
            
            total_vol = total_vol + ws.Cells(i, 7)
            
            
            'ticker value
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                    ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value

                'Stock volume for the summary table
                ws.Cells(rowcount, 12).Value = total_vol

                'year end price
                year_close = ws.Cells(i, 6).Value

                'Price change for the year and summary table value
                year_change = year_close - year_open
                ws.Cells(rowcount, 10).Value = year_change

                'Colours for positive or negative change.
                If year_change >= 0 Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If

                'Percent change for the year and summary table value
                If year_open = 0 And year_close = 0 Then
                
                    percent_change = 0
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                
                ElseIf year_open = 0 Then
                
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    ws.Cells(rowcount, 11).Value = percent_change
                
                Else
                    
                    percent_change = year_change / year_open
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                
                End If

               
                rowcount = rowcount + 1
                
                total_vol = 0
                year_open = 0
                year_close = 0
                year_change = 0
                percent_change = 0
                
            End If
        Next i

        'Titles for best/worst performance table
        '
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        'Summary table last  row
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'Set variables to hold best performer, worst performer, and stock with the most volume
        Dim best_stock As String
        Dim best_value As Double

        'Set best and worst performer equal to the first stock
        
        best_value = ws.Cells(2, 11).Value

        Dim worst_stock As String
        Dim worst_value As Double

     
        worst_value = ws.Cells(2, 11).Value

        Dim most_vol_stock As String
        Dim most_vol_value As Double

        most_vol_value = ws.Cells(2, 12).Value

        'Loop for summary table
        For j = 2 To lastrow

            If ws.Cells(j, 11).Value > best_value Then
                best_value = ws.Cells(j, 11).Value
                best_stock = ws.Cells(j, 9).Value
            End If

            If ws.Cells(j, 11).Value < worst_value Then
                worst_value = ws.Cells(j, 11).Value
                worst_stock = ws.Cells(j, 9).Value
            End If

            If ws.Cells(j, 12).Value > most_vol_value Then
                most_vol_value = ws.Cells(j, 12).Value
                most_vol_stock = ws.Cells(j, 9).Value
            End If

        Next j

        'Move best performer, worst performer, and stock with the most volume items to the performance table
        ws.Cells(2, 16).Value = best_stock
        ws.Cells(2, 17).Value = best_value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = worst_stock
        ws.Cells(3, 17).Value = worst_value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = most_vol_stock
        ws.Cells(4, 17).Value = most_vol_value

        
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("O:Q").EntireColumn.AutoFit

    Next ws
    
End Sub
