Sub SeperateStock()
    Dim StockIndex, LastRow As Long
    StockIndex = 0
    
    
    'Get the last row for use in for loop ranges
    LastRow = Range("A:A").SpecialCells(xlCellTypeLastCell).Row
    'Set the titles in the spreadsheet
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Annual Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume Traded"
    
    
    'initialize values
    Dim Stocks As String 'ticker name
    Dim StockVol As Double 'amount traded additive
    Dim StockOpen As Double 'value of stock on open on first day of the year
    Dim StockChangeAnnual, StockChangeAnnualPercent, StockChangeMaximum As Double  'change as value, percent and recording max for the year
    
    'initialize before loop begins
    StockVol = 0
    StockOpen = Cells(2, 3).Value

    For i = 2 To LastRow
    'go through list
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then  'transition from one stock name to next
            Stocks = Cells(i, 1).Value
            StockVol = StockVol + Cells(i, 7).Value
            StockChangeAnnual = Cells(i, 6) - StockOpen
            StockChangeAnnualPercent = StockChangeAnnual '/ StockOpen
            StockIndex = StockIndex + 1
                       
            Cells(StockIndex + 1, 9).Value = Stocks
            Cells(StockIndex + 1, 10).Value = StockChangeAnnual
            Cells(StockIndex + 1, 11).Value = StockChangeAnnualPercent
            Cells(StockIndex + 1, 12).Value = StockVol
            
            If StockChangeAnnual <= 0 Then 'colors the cells depending on gain or loss
                Cells(StockIndex + 1, 10).Interior.ColorIndex = 3
            Else
                Cells(StockIndex + 1, 10).Interior.ColorIndex = 4
            End If
            
            'format cells to best represent the data
            Cells(StockIndex + 1, 11).NumberFormat = "#,###.##%;-#,###.##%;"
            Cells(StockIndex + 1, 12).NumberFormat = "#,###;"
            'reset for next stock
            StockVol = 0
            StockOpen = Cells(i + 1, 3).Value
            
            Else ' if its not new add the trade volume
            StockVol = StockVol + Cells(i, 7).Value
        End If
    Next i
End Sub