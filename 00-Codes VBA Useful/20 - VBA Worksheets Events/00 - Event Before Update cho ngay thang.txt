Private Sub tbxStart_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not Me.tbxStart Like "??:??" Then
        MsgBox "Please use format 'hh:mm'"
        Cancel = True
        Exit Sub
    End If
    
    myVar = Application.WorksheetFunction.Text(Me.tbxStart, "hh:mm am")
    Me.tbxStart = myVar
End Sub


Private Sub tbxEndDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    tbxEndDate.Value = CDate(tbxEndDate.Value)

End Sub

Private Sub tbxStartDate_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    tbxStartDate.Value = CDate(tbxStartDate.Value)

End Sub