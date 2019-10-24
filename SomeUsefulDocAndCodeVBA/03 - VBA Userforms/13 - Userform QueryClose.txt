'Query: 1 dnag truy van

Private Sub cbClose_Click()
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 1 Then
    'Do nothing
    Else
        Cancel = True
        MsgBox "Can't close this way!", vbOKOnly + vbCritical
    End If
End Sub
