Function CotDoiSangChu(n As Integer) As String
    Dim s As String
    s = Cells(1, n).Address
    CotDoiSangChu = Mid(s, 2, InStr(2, s, "$") - 2)
End Function