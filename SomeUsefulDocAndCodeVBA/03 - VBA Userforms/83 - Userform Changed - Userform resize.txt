Private Sub UserForm_Resize()

    Dim dFrameCols As Double, dFrameRows As Double, dFrameHeight As Double
    Dim i As Integer, j As Integer

    'Standard control gap of 6pts
    Const dGAP As Integer = 6

    'Exit the sub if we've been minimized
    If Me.InsideWidth = 0 Then Exit Sub

    'Set controls that don't move/size
    With lblMessage                              'The position of the "Message" label
        .Top = dGAP
        .Left = dGAP
    End With

    With tbMessage                               'The position of the message box (the size changes, not the position)
        .Top = dGAP + lblMessage.Height + dGAP
        .Left = dGAP
    End With

    fraStyle.Left = dGAP

    'Don't let the form get less than a certain height - must have at least the message and button
    If Me.InsideHeight < lblMessage.Height + btnOK.Height + fraStyle.Height + dGAP * 5 Then

        'Reset the height, allowing for the form's border (Height - InsideHeight)
        Me.Height = lblMessage.Height + btnOK.Height + fraStyle.Height + dGAP * 5 + Me.Height - Me.InsideHeight
    End If

    'Don't let the form get less than a certain width - must be as wide as the biggest check box, plus the standard gap
    If Me.InsideWidth < cbMaximize.Width + fraStyle.Width - fraStyle.InsideWidth + dGAP * 4 Then

        'Reset the width, allowing for the form's border (Width - InsideWidth)
        Me.Width = cbMaximize.Width + fraStyle.Width - fraStyle.InsideWidth + dGAP * 4
    End If

    'Work out the new dimensions of the frame (as the check boxes move within the frame)
    With fraStyle
        dFrameCols = Application.Max(1, (Me.InsideWidth - dGAP * 3 - (.Width - .InsideWidth)) \ (cbMaximize.Width + dGAP))
        dFrameRows = .Controls.Count / dFrameCols

        If dFrameRows <> Int(dFrameRows) Then dFrameRows = Int(dFrameRows) + 1
        dFrameHeight = dFrameRows * cbMaximize.Height + dGAP + .Height - .InsideHeight
    End With

    'Don't allow the form width to decrease so that there's no room for the checkboxes
    'i.e. decreasing the width causes the check boxes to require an extra row, which doesn't fit.
    If Me.InsideHeight <= btnOK.Height + lblMessage.Height + dFrameHeight + dGAP * 5 Then

        'Reset the width, allowing for the form's border (Width - InsideWidth)
        Me.Width = fraStyle.Width + dGAP * 2 + Me.Width - Me.InsideWidth

        'Recalculate the frame's dimensions with the changed form's width
        With fraStyle
            dFrameCols = Application.Max(1, (Me.InsideWidth - dGAP * 3 - (.Width - .InsideWidth)) \ (cbMaximize.Width + dGAP))
            dFrameRows = .Controls.Count / dFrameCols

            If dFrameRows <> Int(dFrameRows) Then dFrameRows = Int(dFrameRows) + 1
            dFrameHeight = dFrameRows * cbMaximize.Height + dGAP + .Height - .InsideHeight
        End With

    End If

    'Set the OK button to be in the middle at the bottom
    With btnOK
        .Left = (Me.InsideWidth - btnOK.Width) / 2
        .Top = Me.InsideHeight - btnOK.Height - dGAP
    End With

    'Sometimes the OK button leaves white lines from its edges, so use a label to clear them
    With lblBlank
        .Width = Me.InsideWidth
        .Top = btnOK.Top - 0.75
    End With

    'Set the frame to be as wide as the box and move the check boxes in it to fit
    With fraStyle
        .Width = Me.InsideWidth - dGAP * 2
        .Height = dFrameHeight

        'Reposition the controls in the frame, according to their tab order
        For i = 0 To .Controls.Count - 1
            For j = 0 To .Controls.Count - 1
                With .Controls(j)
                    If .TabIndex = i Then
                        .Left = (i Mod dFrameCols) * (cbMaximize.Width + dGAP) + dGAP
                        .Top = Int(i / dFrameCols) * cbMaximize.Height + dGAP
                    End If
                End With
            Next
        Next

        .Top = btnOK.Top - dGAP - .Height
    End With

    'Userform is big enough, so set the message box's height and width to fill it
    With tbMessage
        .Width = Me.InsideWidth - dGAP * 2

        'Don't allow the height to go negative
        .Height = Application.Max(0, fraStyle.Top - .Top - dGAP)
    End With
End Sub