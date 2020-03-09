Function VBAProper(ByVal MyText As String) As String
    Dim i As Long
    On Error Resume Next
    If Trim(MyText) <> "" And Not (IsNumeric(MyText)) Then
        If Len(MyText) = 1 Then
            VBAProper = UCase(MyText)
        Else
            MyText = UCase(Left(MyText, 1)) & LCase(Mid(MyText, 2, Len(MyText)))
            For i = 2 To Len(MyText)
                If UCase(Mid(MyText, i, 1)) <> LCase(Mid(MyText, i, 1)) Then
                    If UCase(Mid(MyText, i - 1, 1)) = LCase(Mid(MyText, i - 1, 1)) Then
                        MyText = Left(MyText, i - 1) & Replace(MyText, Mid(MyText, i, 1), UCase(Mid(MyText, i, 1)), i, 1)
                    End If
                End If
            Next
            VBAProper = MyText
        End If
    End If
End Function