Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sh As Worksheet, lastRow&
    Dim shp As Shape
    Set sh = ActiveSheet
    If sh.Range("A1") = "" Then
    lastRow = 0
    Else
        lastRow = sh.Range("A" & Rows.Count).End(xlUp).Row
    End If

        sh.Cells(lastRow + 1, 1) = "I    have    a    great    day!"
    If lastRow > 20 Then ActiveWindow.SmallScroll (1)
    Set shp = sh.Shapes("Rectangle 1")
    With shp.Fill.ForeColor
    If .RGB = RGB(0, 0, 255) Then
        .RGB = RGB(197, 240, 255)
    Else
        .RGB = RGB(0, 0, 255)
    End If
     End With
End Sub
