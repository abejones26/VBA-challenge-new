Sub VBA_Of_Wall_Street()
    Dim ws As Worksheet, i As Long, last_row As Long, result_table_row As Integer
    Dim open_price As Double, close_price As Double, yearly_change As Double
    Dim yearly_change_percentage As Double, Stock_Total_Volume As Long
    Dim greatest_increase_ticker As String, greatest_increase_percentage As Double, greatest_decrease_ticker As String, greatest_decrease_percentage As Double, greatest_total_ticker As String, greatest_total_volume As Long
    
    ' Create Loop for all sheets
    For Each ws In Worksheets

        'add the variable to columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly_Changes"
        ws.Cells(1, 11).Value = "Percent_Changes"
        ws.Cells(1, 12).Value = "Stock_Total_Volume"
        ws.Cells(1, 15).Value = "Tickers"
        ws.Cells(1, 16).Value = "Values"
        
        ' Find last row and count up
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        result_table_row = 2
        Stock_Total_Volume = 0
        greatest_increase_ticker = ""
        greatest_increase_percentage = 0
        greatest_decrease_ticker = ""
        greatest_decrease_percentage = 0
        greatest_total_ticker = ""
        greatest_total_volume = 0
        
        ' Set First Varable of the first ticker
        ws.Cells(result_table_row, 9).Value = ws.Cells(2, 1).Value
        
        ' Set Opening Price of the first ticker
        open_price = ws.Cells(2, 3).Value
        
        ' Cells(result_table_row, 13).Value = open_price
        
        ' Loop through all row until and stop at lastrow
        For i = 2 To last_row
        
        Stock_Total_Volume = Stock_Total_Volume + ws.Cells(i, 7).Value
            
            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                'Stock_Total_Volume = Stock_Total_Volume + ws.Cells(i, 7).Value

                ' Set close price for the previous ticker and yearly change before open_price gets overridden.
                close_price = ws.Cells(i, 6).Value

                ' Find the difference from the beginning of the year's OPEN and end of the year's Close
                yearly_change = close_price - open_price

                If open_price <> 0 Then
                    yearly_change_percentage = yearly_change / open_price
                Else
                    yearly_change_percentage = 0
                End If
                
                ws.Cells(result_table_row, 12).Value = Stock_Total_Volume

                ' Print yearly change in table
                ' Cells(result_table_row, 14).Value = close_price
                ws.Cells(result_table_row, 10).Value = yearly_change

                    If ws.Cells(result_table_row, 10).Value >= 0 Then

                        ' Color the Passing grade green
                        ws.Cells(result_table_row, 10).Interior.ColorIndex = 4

                    Else
                        
                        ' Color the Failing grade red
                        ws.Cells(result_table_row, 10).Interior.ColorIndex = 3

                    End If

                ws.Cells(result_table_row, 11).Value = Format(yearly_change_percentage, "0.00%")
                
                If yearly_change_percentage > greatest_increase_percentage Then
                    greatest_increase_percentage = yearly_change_percentage
                    greatest_increase_ticker = ws.Cells(i, 1).Value
                End If
                
                If yearly_change_percentage < greatest_decrease_percentage Then
                    greatest_decrease_percentage = yearly_change_percentage
                    greatest_decrease_ticker = ws.Cells(i, 1).Value
                End If
                
                If Stock_Total_Volume > greatest_total_volume Then
                    greatest_total_volume = Stock_Total_Volume
                    greatest_total_ticker = ws.Cells(i, 1).Value
                End If

                ' Add on the summary table row
                result_table_row = result_table_row + 1

                Stock_Total_Volume = 0

                ' Print the next ticker value (A, AA, etc.)
                ws.Cells(result_table_row, 9).Value = Cells(i + 1, 1).Value
                
                ' Set open price for the next ticker
                open_price = ws.Cells(i + 1, 3).Value
                
                'Cells(result_table_row, 13).Value = open_price
            End If
        Next i

        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = greatest_increase_ticker
        ws.Cells(2, 16).Value = Format(greatest_increase_percentage, "0.00%")
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = greatest_decrease_ticker
        ws.Cells(3, 16).Value = Format(greatest_decrease_percentage, "0.00%")
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = greatest_total_ticker
        ws.Cells(4, 16).Value = greatest_total_volume

    Next ws
End Sub


