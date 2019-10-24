Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'If we've disabled the [x] close button, [B]prevent the Alt+F4 keyboard shortcut too[/B] Không cho phép ngay c? khi ngu?i dùng nh?n t? h?p Alt + F4
    If CloseMode = vbFormControlMenu And Not cbCloseBtn.Value Then
        Cancel = True
    End If
End Sub