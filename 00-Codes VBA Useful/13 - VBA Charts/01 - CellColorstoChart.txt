Sub CellColorsToChart()

    Dim xChart As Chart
    Dim I As Long, J As Long
    Dim xRowsOrCols As Long, xSCount As Long
    Dim xRg As Range, xCell As Range
    On Error Resume Next
    Set xChart = ActiveSheet.ChartObjects("Graphique 6").Chart
    If xChart Is Nothing Then Exit Sub
    xSCount = xChart.SeriesCollection.count
    For I = 1 To xSCount
        J = 1
        With xChart.SeriesCollection(I)
            Set xRg = ActiveSheet.Range(Split(Split(.Formula, ",")(2), "!")(1))
            If xSCount > 4 Then
                xRowsOrCols = xRg.Columns.count
            Else
                xRowsOrCols = xRg.Rows.count
            End If
            For Each xCell In xRg
                .Points(J).Format.Fill.ForeColor.RGB = ThisWorkbook.Colors(xCell.Interior.ColorIndex)
                .Points(J).Format.Line.ForeColor.RGB = ThisWorkbook.Colors(xCell.Interior.ColorIndex)
                J = J + 1
            Next
        End With
    Next

End Sub