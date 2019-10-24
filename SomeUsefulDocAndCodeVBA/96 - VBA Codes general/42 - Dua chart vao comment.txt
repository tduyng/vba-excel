Sub PlaceGraph()
    Dim x As String, z As Range
    Application.ScreenUpdating = False
    ' Ðu?ng d?n t?m th?i d? luu gi? hình ?nh
    ' Các b?n có th? thay d?i tùy theo nhu c?u c?a mình
    x = "C:\XWMJGraph.gif"
    ' Ô ch?a comment
    Set z = Worksheets("ChartInComment").Range("A3")
    ' Xóa comment t?i ô này
    On Error Resume Next
    z.Comment.Delete
    On Error GoTo 0
    ' Ch?n và xu?t chart
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Export x
    ' Thêm comment m?i vào ô, thi?t l?p kích thu?c và thêm chart (d?ng hình ?nh) vào comment
    With z.AddComment
        With .Shape
            .Height = 322
            .Width = 465
            .Fill.UserPicture x
        End With
    End With
    ' Xóa t?p tin hình ?nh t?m
    Kill x
    Range("A1").Activate
    Application.ScreenUpdating = True
    Set z = Nothing
End Sub