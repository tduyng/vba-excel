Option Explicit



Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    With Image1
        .SpecialEffect = fmSpecialEffectRaised
        .Left = 103
        .Top = 19
    End With
End Sub

Private Sub Image2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    With Image2
        .SpecialEffect = fmSpecialEffectRaised
        .Left = 103
        .Top = 85
    End With
End Sub
Private Sub Image3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    With Image3
        .SpecialEffect = fmSpecialEffectRaised
        .Left = 103
        .Top = 151
    End With
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    With Image1
        .SpecialEffect = fmSpecialEffectFlat
        .Left = 102
        .Top = 18
    End With
    
        With Image2
        .SpecialEffect = fmSpecialEffectFlat
        .Left = 102
        .Top = 84
    End With
    
        With Image3
        .SpecialEffect = fmSpecialEffectFlat
        .Left = 102
        .Top = 150
    End With
    
End Sub


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
