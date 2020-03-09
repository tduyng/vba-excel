Sub CommentFitter2()
    Application.ScreenUpdating = False
    Dim x As Range, y As Long
    For Each x In Cells.SpecialCells(xlCellTypeComments)
        Select Case True
        Case Len(x.NoteText) <> 0
            With x.Comment
                .Shape.TextFrame.AutoSize = True
                If .Shape.Width > 250 Then
                    y = .Shape.Width * .Shape.Height
                    .Shape.ScaleHeight 0.9, msoFalse, msoScaleFromTopLeft
                    .Shape.ScaleWidth 1#, msoFalse, msoScaleFromTopLeft
                End If
            End With
        End Select
    Next x
    Application.ScreenUpdating = True
End Sub