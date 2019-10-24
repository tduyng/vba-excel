Sub QUERY_NAME_SHEET()

Dim sh As Worksheet
Dim arr As Variant, i As Integer

ReDim arr(1 To ThisWorkbook.Worksheets.Count)
i = 1
For Each sh In Worksheets
    arr(i) = sh.Name
    i = i + 1
Next sh

If i = 1 Then Exit Sub
Sheet2.Range("H2").Resize(UBound(arr), 1).Value = Application.Transpose(arr)
End Sub
