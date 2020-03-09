Private Sub CommandButton1_Click()
stDate = CDate("1/1/2013")

For x = 1 To 3
    If Me.Controls("cbP" & x) = True Then
        Me.Controls("tbP" & x).Text = stDate
    Else
        Me.Controls("tbP" & x) = ""
    End If
    stDate = DateAdd("m", 1, stDate)
    
Next x

Me.Controls("CommandButton1").Caption = "Good Work!"

End Sub