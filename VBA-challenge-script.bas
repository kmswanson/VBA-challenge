Attribute VB_Name = "Module1"
Sub alphaTesting():

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Columns.AutoFit
    'hardcoding column titles and formatting to autofit
        
    Dim lastRow As Long
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'declared variable to reference last row in the sheet
    
    Dim tickerName As String
    'ticker name
    Dim totalStockValue As Double
    'volume total
    totalStockValue = 0
    'initialize stock value at 0
    Dim summaryTableRow As Double
    summaryTableRow = 2
    'start at row 2
    
    Dim openPrice As Double
    Dim closePrice As Double
    Dim percentChange As Double
    Dim tickerIncrease, tickerDecrease, tickerVolume As String
    'assigning variable names
    
    openToClose = 0
    greatValue = 0
    greatIncrease = 0
    greatDecrease = 0
    'initializing values at 0
    
    For Row = 2 To lastRow
    'start loop at row two until the end (lastRow)
    
        If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then  'if ticker changes
            tickerName = ws.Cells(Row, 1).Value
            'set ticker name
            
            totalStockValue = totalStockValue + ws.Cells(Row, 7).Value
            'add onto value
            
            ws.Cells(summaryTableRow, 9).Value = tickerName
            'place ticker name in column I
            
             openPrice = ws.Range("C" & (Row - openToClose))
             closePrice = ws.Range("F" & Row)
             yearlyChange = (closePrice - openPrice)
             ws.Cells(summaryTableRow, 10).Value = yearlyChange
             percentChange = ((closePrice - openPrice) / openPrice) * 100
             ws.Cells(summaryTableRow, 11).Value = percentChange & "%"
             
            If yearlyChange <= 0 Then
               ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
               Else
               ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
            End If
            'conditional formatting to assign color to values either less than or greater than 0

            
            If percentChange > greatIncrease Then
                tickerIncrease = tickerName
                greatIncrease = percentChange
                ElseIf percentChange < greatDecrease Then
                    tickerDecrease = tickerName
                    greatDecrease = percentChange
            End If
                
            If totalStockValue > greatValue Then
                tickerVolume = tickerName
                greatValue = totalStockValue
            End If
                    
                
            ws.Cells(summaryTableRow, 12).Value = totalStockValue
            'put totalStockValue in column L
            
            summaryTableRow = summaryTableRow + 1
            'add one to summaryTableRow to move down the list
            
            totalStockValue = 0
            'resets totalStockValue to 0
            
             openToClose = -1
            
        Else
            totalStockValue = totalStockValue + ws.Cells(Row, 7).Value
            'if ticker does not change, add values from column G to totalStockValue
            
        End If
        
             openToClose = openToClose + 1
                
    Next Row

    lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 3
    lr2 = ws.Cells(Rows.Count, lc + 1).End(xlUp).Row + 1
            ws.Cells(lr2, lc).Value = "Greatest % Increase"
            ws.Cells(lr2, lc + 1).Value = tickerIncrease
            ws.Cells(lr2, lc + 2).Value = greatIncrease & "%"
            ws.Cells(lr2 + 1, lc).Value = "Greatest % Decrease"
            ws.Cells(lr2 + 1, lc + 1).Value = tickerDecrease
            ws.Cells(lr2 + 1, lc + 2).Value = greatDecrease & "%"
            ws.Cells(lr2 + 2, lc).Value = "Greatest Total Volume"
            ws.Cells(lr2 + 2, lc + 1).Value = tickerVolume
            ws.Cells(lr2 + 2, lc + 2).Value = greatValue
            ws.Cells(lr2 - 1, lc + 1).Value = "Ticker"
            ws.Cells(lr2 - 1, lc + 2).Value = "Value"
            ws.Columns.AutoFit
    'hardcoding to place smaller summary table (O:Q) including titles and data
    
    Next ws
End Sub
