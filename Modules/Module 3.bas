Attribute VB_Name = "Module3"
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

