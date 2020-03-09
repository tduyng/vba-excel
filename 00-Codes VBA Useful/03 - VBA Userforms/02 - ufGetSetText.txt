Option Explicit

Private Sub cbGetText_Click()
   MsgBox txtText.Value
   Unload Me
End Sub

Private Sub cbSetText_Click()
    lbText.Caption = txtText.Value
    'Unload Me
End Sub

Private Sub txtText_Change()
    lbText.Caption = txtText.Value
End Sub

Private Sub txtText_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
    Me.txtText.MaxLength = 6


End Sub