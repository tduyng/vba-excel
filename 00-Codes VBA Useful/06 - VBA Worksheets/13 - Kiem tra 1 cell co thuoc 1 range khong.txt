Sub CellinRange()
    Dim rngArea As Range

    ' Vùng d? li?u b?n mu?n ki?m tra
    ' B?n có th? thay d?i theo ý b?n
    Set rngArea = Range("A1:C5")   
   
' Dùng Intersect d? ki?m tra
    If Application.Intersect(rngArea, ActiveCell) Is Nothing Then
       MsgBox ("Ô hi?n t?i không có trong vùng này.")
    Else
       MsgBox ("Ô hi?n t?i dang ? trong vùng này.")
    End If
End Sub