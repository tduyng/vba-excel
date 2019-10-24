Sub ColorChartColumnsbyCellColor()

    Dim xChart As Chart
    Dim I As Long, xRows As Long
    Dim xRg As Range, xCell As Range
    On Error Resume Next
    Set xChart = ActiveSheet.ChartObjects("Chart 1").Chart
    If xChart Is Nothing Then Exit Sub
    With xChart.SeriesCollection(1)
        Set xRg = ActiveSheet.Range(Split(Split(.Formula, ",")(1), "!")(1))
        xRows = xRg.Rows.Count
        Set xRg = xRg(1)
        For I = 1 To xRows
            .Points(I).Format.Fill.ForeColor.RGB = ThisWorkbook.Colors(xRg.Offset(I - 1, 0).Interior.ColorIndex)
        Next
    End With
End Sub