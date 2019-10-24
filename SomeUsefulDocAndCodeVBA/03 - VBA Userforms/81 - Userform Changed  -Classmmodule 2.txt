'***************************************************************
'* Private procedures to perform the updates
'***************************************************************

'Procedure to set the form's window style
Private Sub SetFormStyle()

    Dim lStyle As Long, hMenu As Long

    'Have we got a form to set?
    If mhWndForm = 0 Then Exit Sub

    'Get the basic window style
    lStyle = GetWindowLong(mhWndForm, GWL_STYLE)

    'Build up the basic window style flags for the form
    SetBit lStyle, WS_CAPTION, mbCaption
    SetBit lStyle, WS_SYSMENU, mbSysMenu
    SetBit lStyle, WS_THICKFRAME, mbSizeable
    SetBit lStyle, WS_MINIMIZEBOX, mbMinimize
    SetBit lStyle, WS_MAXIMIZEBOX, mbMaximize
    
    'Set the basic window styles
    SetWindowLong mhWndForm, GWL_STYLE, lStyle

    'Get the extended window style
    lStyle = GetWindowLong(mhWndForm, GWL_EXSTYLE)

    'Build up and set the extended window style
    SetBit lStyle, WS_EX_APPWINDOW, mbAppWindow
    SetBit lStyle, WS_EX_TOOLWINDOW, mbToolWindow
    
    SetWindowLong mhWndForm, GWL_EXSTYLE, lStyle

    'Handle the close button differently
    If mbCloseBtn Then
        'We want it, so reset the control menu
        hMenu = GetSystemMenu(mhWndForm, 1)
    Else
        'We don't want it, so delete it from the control menu
        hMenu = GetSystemMenu(mhWndForm, 0)
        DeleteMenu hMenu, SC_CLOSE, 0&
    End If

    'Update the window with the changes
    DrawMenuBar mhWndForm
    SetFocus mhWndForm

End Sub

'Procedure to set or clear a bit from a style flag
Private Sub SetBit(ByRef lStyle As Long, ByVal lBit As Long, ByVal bOn As Boolean)
    If bOn Then
        lStyle = lStyle Or lBit
    Else
        lStyle = lStyle And Not lBit
    End If
End Sub

'Procedure to set the form's icon
Private Sub ChangeIcon()

    Dim hIcon As Long

    On Error Resume Next

    If mhWndForm <> 0 Then

        Err.Clear
        If msIconPath = "" Then
            hIcon = 0
        ElseIf Dir(msIconPath) = "" Then
            hIcon = 0
        ElseIf Err.Number <> 0 Then
            hIcon = 0
        ElseIf Not mbIcon Then
            'Get the icon from the source
            hIcon = ExtractIcon(0, msIconPath, 0)
        Else
            hIcon = 0
        End If

        'Set the big (32x32) and small (16x16) icons
        SendMessage mhWndForm, WM_SETICON, True, hIcon
        SendMessage mhWndForm, WM_SETICON, False, hIcon
    End If

End Sub