Sub CalculateStockMetrics()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim ticker As String
    Dim startRow As Long
    Dim endRow As Long
    Dim totalVolume As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    
    Dim outputRow As Long
    outputRow = 1
    
    ws.Cells(outputRow, 9).Value = "Ticker"
    ws.Cells(outputRow, 10).Value = "Quarterly Change"
    ws.Cells(outputRow, 11).Value = "Percent Change"
    ws.Cells(outputRow, 12).Value = "Total Volume"
    outputRow = outputRow + 1
    
    ticker = ws.Cells(2, 1).Value
    startRow = 2
    totalVolume = 0
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> ticker Or i = lastRow Then
            If i = lastRow Then
                endRow = i
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            Else
                endRow = i - 1
            End If
            
            openingPrice = ws.Cells(startRow, 3).Value
            closingPrice = ws.Cells(endRow, 6).Value
            quarterlyChange = closingPrice - openingPrice
            percentChange = (quarterlyChange / openingPrice) * 100

            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).Value = Format(percentChange, "0.00") & "%"
            ws.Cells(outputRow, 12).Value = totalVolume
            outputRow = outputRow + 1
            
            ticker = ws.Cells(i, 1).Value
            startRow = i
            totalVolume = 0
        End If
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    Next i
    
    MsgBox "Calculations complete!"
End Sub

Sub CalculateStockMetricsAndHighlights()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim ticker As String
    Dim startRow As Long
    Dim endRow As Long
    Dim totalVolume As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncTicker As String
    Dim maxDecTicker As String
    Dim maxVolTicker As String
    maxIncrease = -1000000
    maxDecrease = 1000000
    maxVolume = 0
    
    Dim outputRow As Long
    outputRow = 1
    
    ws.Cells(outputRow, 9).Value = "Ticker"
    ws.Cells(outputRow, 10).Value = "Quarterly Change"
    ws.Cells(outputRow, 11).Value = "Percent Change"
    ws.Cells(outputRow, 12).Value = "Total Volume"
    outputRow = outputRow + 1
    
    ticker = ws.Cells(2, 1).Value
    startRow = 2
    totalVolume = 0
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> ticker Or i = lastRow Then
            If i = lastRow Then
                endRow = i
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            Else
                endRow = i - 1
            End If
            
            openingPrice = ws.Cells(startRow, 3).Value
            closingPrice = ws.Cells(endRow, 6).Value
            quarterlyChange = closingPrice - openingPrice
            percentChange = (quarterlyChange / openingPrice) * 100
            
            If percentChange > maxIncrease Then
                maxIncrease = percentChange
                maxIncTicker = ticker
            End If
            If percentChange < maxDecrease Then
                maxDecrease = percentChange
                maxDecTicker = ticker
            End If
            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                maxVolTicker = ticker
            End If
            
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).Value = Format(percentChange, "0.00") & "%"
            ws.Cells(outputRow, 12).Value = totalVolume
            outputRow = outputRow + 1
            
            ticker = ws.Cells(i, 1).Value
            startRow = i
            totalVolume = 0
        End If
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    Next i
    
    ws.Cells(2, 14).Value = "Metric"
    ws.Cells(2, 15).Value = "Ticker"
    ws.Cells(2, 16).Value = "Value"
    
    ws.Cells(3, 14).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = maxIncTicker
    ws.Cells(3, 16).Value = Format(maxIncrease, "0.00") & "%"
    
    ws.Cells(4, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = maxDecTicker
    ws.Cells(4, 16).Value = Format(maxDecrease, "0.00") & "%"
    
    ws.Cells(5, 14).Value = "Greatest Total Volume"
    ws.Cells(5, 15).Value = maxVolTicker
    ws.Cells(5, 16).Value = maxVolume
    
    MsgBox "Calculations complete!"
End Sub

Sub CalculateStockMetricsAcrossQuarters()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, startRow As Long, endRow As Long
    Dim ticker As String, maxIncTicker As String, maxDecTicker As String, maxVolTicker As String
    Dim totalVolume As Double, openingPrice As Double, closingPrice As Double
    Dim quarterlyChange As Double, percentChange As Double
    Dim maxIncrease As Double, maxDecrease As Double, maxVolume As Double
    
    For Each ws In wb.Sheets
        If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then
            ws.Activate
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            maxIncrease = -1000000
            maxDecrease = 1000000
            maxVolume = 0
            
            ticker = ws.Cells(2, 1).Value
            startRow = 2
            totalVolume = 0
            
            Dim outputRow As Long
            outputRow = 1
            
            With ws
                .Cells(outputRow, 9).Value = "Ticker"
                .Cells(outputRow, 10).Value = "Quarterly Change"
                .Cells(outputRow, 11).Value = "Percent Change"
                .Cells(outputRow, 12).Value = "Total Volume"
                outputRow = outputRow + 1
                
                For i = 2 To lastRow
                    If .Cells(i, 1).Value <> ticker Or i = lastRow Then
                        If i = lastRow Then
                            endRow = i
                            totalVolume = totalVolume + .Cells(i, 7).Value
                        Else
                            endRow = i - 1
                        End If
                        
                        openingPrice = .Cells(startRow, 3).Value
                        closingPrice = .Cells(endRow, 6).Value
                        quarterlyChange = closingPrice - openingPrice
                        percentChange = (quarterlyChange / openingPrice) * 100
                        
                        If percentChange > maxIncrease Then
                            maxIncrease = percentChange
                            maxIncTicker = ticker
                        End If
                        If percentChange < maxDecrease Then
                            maxDecrease = percentChange
                            maxDecTicker = ticker
                        End If
                        If totalVolume > maxVolume Then
                            maxVolume = totalVolume
                            maxVolTicker = ticker
                        End If
                        
                        .Cells(outputRow, 9).Value = ticker
                        .Cells(outputRow, 10).Value = quarterlyChange
                        .Cells(outputRow, 11).Value = Format(percentChange, "0.00") & "%"
                        .Cells(outputRow, 12).Value = totalVolume
                        outputRow = outputRow + 1
                        
                        ticker = .Cells(i, 1).Value
                        startRow = i
                        totalVolume = 0
                    End If
                    totalVolume = totalVolume + .Cells(i, 7).Value
                Next i
            End With
            
            With ws
                .Cells(2, 14).Value = "Metric"
                .Cells(2, 15).Value = "Ticker"
                .Cells(2, 16).Value = "Value"
                
                .Cells(3, 14).Value = "Greatest % Increase"
                .Cells(3, 15).Value = maxIncTicker
                .Cells(3, 16).Value = Format(maxIncrease, "0.00") & "%"
                
                .Cells(4, 14).Value = "Greatest % Decrease"
                .Cells(4, 15).Value = maxDecTicker
                .Cells(4, 16).Value = Format(maxDecrease, "0.00") & "%"
                
                .Cells(5, 14).Value = "Greatest Total Volume"
                .Cells(5, 15).Value = maxVolTicker
                .Cells(5, 16).Value = maxVolume
            End With
        End If
    Next ws
    
    MsgBox "Calculations complete for all quarters!"
End Sub

Sub ApplyPercentChangeFormatting()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    For Each ws In wb.Worksheets

        lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        
        With ws.Range("J2:J" & lastRow)

            .FormatConditions.Delete
            
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)
            
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 255)
            
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(3).Interior.Color = RGB(0, 255, 0)
        End With
    Next ws
    
    MsgBox "Conditional formatting applied to all sheets!"
End Sub
