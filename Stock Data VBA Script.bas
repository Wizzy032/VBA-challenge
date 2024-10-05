Attribute VB_Name = "Module1"
Sub StockData()
    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Long
    Dim startPrice As Double
    Dim endPrice As Double
    Dim totalVolume As Double
    Dim change As Double
    Dim percentChange As Double
    Dim currentTicker As String
    Dim previousTicker As String
    Dim resultRow As Long
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestVolume As Double

'Loop through every worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
'Loop through the rows
        startRow = 2
        previousTicker = ws.Cells(startRow, 1).Value
        startPrice = ws.Cells(startRow, 3).Value
        totalVolume = 0
        
'Output the result in the same worksheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        resultRow = 2
        
'Initialize variables to very low/high values
        greatestPercentIncrease = -100000
        greatestPercentDecrease = 100000
        greatestVolume = 0
        
        For i = 2 To lastRow + 1
            If i <= lastRow Then
                currentTicker = ws.Cells(i, 1).Value
                If currentTicker <> previousTicker Or i = lastRow Then
                    If i = lastRow And currentTicker = previousTicker Then
                        totalVolume = totalVolume + ws.Cells(i, 7).Value
                    End If
                    endPrice = ws.Cells(i - 1, 6).Value
                    change = endPrice - startPrice
                     If startPrice <> 0 Then
                        percentChange = (change / startPrice)
                    Else
                        percentChange = 0
                    End If
                    
'Check for greatest percent increase, decrease and total volume
                     If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        tickerGreatestIncrease = previousTicker
                    End If
                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        tickerGreatestDecrease = previousTicker
                    End If
                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        tickerGreatestVolume = previousTicker
                    End If
                    ws.Cells(resultRow, 9).Value = previousTicker
                    ws.Cells(resultRow, 10).Value = change
                    
'Apply color formatting
                    If ws.Cells(resultRow, 10).Value > 0 Then
                        ws.Cells(resultRow, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(resultRow, 10).Value < 0 Then
                        ws.Cells(resultRow, 10).Interior.ColorIndex = 3
                    End If
                    ws.Cells(resultRow, 11).Value = percentChange
                    ws.Cells(resultRow, 11).NumberFormat = "0.00%"
                    ws.Cells(resultRow, 12).Value = totalVolume
                    resultRow = resultRow + 1
                    If i < lastRow Then
                        startPrice = ws.Cells(i, 3).Value
                        totalVolume = 0
                    End If
                End If
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                previousTicker = currentTicker
            Else
                ws.Cells(resultRow, 9).Value = previousTicker
                ws.Cells(resultRow, 10).Value = change
                ws.Cells(resultRow, 11).Value = percentChange
                ws.Cells(resultRow, 12).Value = totalVolume
            End If
        Next i
        
'Output the greatest values for each sheet
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = tickerGreatestIncrease
        ws.Cells(3, 16).Value = tickerGreatestDecrease
        ws.Cells(4, 16).Value = tickerGreatestVolume
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 17).Value = greatestVolume
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        
    Next ws
End Sub
