Private Sub Worksheet_Change(ByVal target As Range)
    If Application.EnableEvents = False Then Application.EnableEvents = True
       If Not Intersect(target, Range("A1:A65356")) Is Nothing And target.Rows.Count <> ActiveSheet.Rows.Count Then
        Dim rngi As Range
        For Each rngi In target
            Select Case LCase(rngi.Value)
                    Case "w"
                        rngi.Resize(, 6).Interior.ColorIndex = 3
                    Case "a"
                        rngi.Resize(, 6).Interior.ColorIndex = 4
                    Case Else
                        rngi.Resize(, 6).Interior.ColorIndex = xlColorIndexNone
            End Select
        Next rngi
    
    
    ElseIf Not Intersect(target, Range("A1:A65356")) Is Nothing And target.Rows.Count = ActiveSheet.Rows.Count Then
        On Error GoTo ExitPoint
        Application.EnableEvents = False
        If Application.WorksheetFunction.CountA(target) = 0 Then
            Application.Undo
            MsgBox " You can't delete All cell contents from this range " _
            , vbCritical, "Kutools for Excel"
        End If
    End If
ExitPoint:
        Application.EnableEvents = True
End Sub


'Khong cho phep xoa mot cells dac biet
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Range("A1:E7")) Is Nothing Then Exit Sub
    On Error GoTo ExitPoint
    Application.EnableEvents = False
    If Not IsDate(Target(1)) Then
        Application.Undo
        MsgBox " You can't delete cell contents from this range " _
        , vbCritical, "Kutools for Excel"
    End If
ExitPoint:
    Application.EnableEvents = True
End Sub