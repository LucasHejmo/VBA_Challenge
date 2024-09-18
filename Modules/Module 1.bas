Attribute VB_Name = "Module1"
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
