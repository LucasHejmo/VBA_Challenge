Attribute VB_Name = "Module4"
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
