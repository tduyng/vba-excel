Private Sub LoginButton_Click()


'Source ref: https://stackoverflow.com/questions/45654907/multiple-user-login-form-for-excel-vba


    Application.ScreenUpdating = False
    Dim Username As String
    Dim Password As String
    'Use a variable to flag whether the userid is valid or not
    Dim useridValid As Boolean
    Dim i As Integer
    'Dim j As Integer
    Dim u As String
    Dim p As String
    If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
        MsgBox "Enter username and password.", vbOKOnly
    ElseIf Trim(TextBox1.Text) = "" Then
        MsgBox "Enter the username ", vbOKOnly
    ElseIf Trim(TextBox2.Text) = "" Then
        MsgBox "Enter the Password ", vbOKOnly
    Else
        Username = Trim(TextBox1.Text)
        Password = Trim(TextBox2.Text)
        useridValid = False
        i = 1
        'Don't perform a loop which is dependent on a fixed cell that
        'isn't updated within the loop
        'Use a variable row counter instead
        'Do While Cells(1, 1).Value <> ""
        Do While Cells(i, 1).Value <> ""
            'There is no point in having a variable simply to specify a
            'column that doesn't change
            'j = 1
            u = Cells(i, "A").Value
            'j = j + 1
            p = Cells(i, "B").Value
            'Only perform tests once a valid username has been found
            If Username = u Then
                'Flag that we have found the userid
                useridValid = True
                If Cells(i, "C").Value = "fail" Then
                    'Too many login attempts
                    MsgBox "Your Account temporarily locked", vbCritical
                ElseIf Password = p Then
                    'Clear invalid attempts count
                    Cells(i, 4).Value = 0
                    Cells(i, 3).Value = ""

                    Call clr
                    Unload Me
                    MsgBox ("Welcome " + u + ", :)")
                Else
                    'Invalid password
                    'Increment failed attempts counter
                    Cells(i, 4).Value = Cells(i, 4) + 1
                    'Lock account on 3rd failed password
                    If Cells(i, 4).Value > 2 Then
                        'lock the account
                        Cells(i, 3).Value = "fail"
                        'Cells(i, 2).Value = ""
                        Cells(i, 2).Interior.ColorIndex = 3
                        'Tell the user that password was invalid and now locked
                        MsgBox "Invalid password - account locked", vbCritical
                    Else
                        'Tell the user that password was invalid
                        MsgBox "Invalid password", vbCritical
                    End If
                End If
                'Don't check any further usernames
                Exit Do
            End If
            i = i + 1
        Loop
        'If the flag saying that we found the userid isn't set, display
        'a message
        If Not useridValid Then
            MsgBox "Username not matched", vbCritical + vbOKCancel
        End If
    End If
    Application.ScreenUpdating = True
End Sub