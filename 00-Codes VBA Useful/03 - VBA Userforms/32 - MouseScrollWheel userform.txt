Private Sub cmbMyList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListBoxScroll Me, Me.cmbMyList
End Sub

Private Sub lbxMyList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListBoxScroll Me, Me.lbxMyList
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
UnhookListBoxScroll
End Sub
