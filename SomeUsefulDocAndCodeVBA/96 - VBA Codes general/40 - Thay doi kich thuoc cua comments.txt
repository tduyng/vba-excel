Sub CommentFitter1()
    Application.ScreenUpdating = False
    Dim x As Range, y As Long
    For Each x In Cells.SpecialCells(xlCellTypeComments)
        Select Case True
        Case Len(x.NoteText) <> 0
            With x.Comment
                .Shape.TextFrame.AutoSize = True
                If .Shape.Width > 250 Then
                    y = .Shape.Width * .Shape.Height
                    .Shape.Width = 150
                    .Shape.Height = (y / 200) * 1.3
                End If
            End With
        End Select
    Next x
    Application.ScreenUpdating = True
End Sub