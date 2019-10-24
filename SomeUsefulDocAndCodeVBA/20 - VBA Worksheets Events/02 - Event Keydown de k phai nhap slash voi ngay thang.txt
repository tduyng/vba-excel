Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyBack Then
        If Len(Me.TextBox1) = 4 Then
            Me.TextBox1 = Left(Me.TextBox1, 2)
            KeyCode = False
        ElseIf Len(Me.TextBox1) = 7 Then
            Me.TextBox1 = Left(Me.TextBox1, 5)
            KeyCode = False
        End If

    Else
        If Len(Me.TextBox1) = 2 Or Len(Me.TextBox1) = 5 Then
            'add slash
            Me.TextBox1 = Me.TextBox1 & "/"
    
        End If
    End If
End Sub