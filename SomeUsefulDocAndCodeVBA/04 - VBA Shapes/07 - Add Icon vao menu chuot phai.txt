Sub AddToCellMenu()
    Dim contextMenu As CommandBar
    Dim MySubMenu As CommandBarControl
    Set contextMenu = Application.CommandBars("Cell")
    'them nut save vao menu chuot phai
    contextMenu.Controls.Add Type:=msoControlButton, ID:=3, before:=1
    With contextMenu.Controls.Add(Type:=msoControlButton, before:=2)
        .OnAction = "SayIt" 'thay doi cho phu hop voi ten marcro
        .FaceId = 59
        .Caption = "SayIt"
        .Tag = "Thuc hanh Excel"
    End With
    Set MySubMenu = contextMenu.Controls.Add(Type:=msoControlPopup, before:=3)
    With MySubMenu
        .Caption = "Case Menu"
        .Tag = "Thuc hanh Excel"
        With .Controls.Add(Type:=msoControlsButton)
            .OnAction = "UpperMacro"
            .FaceId = 100
            .Caption = "Upper Case"
        End With
        With .Controls.Add(Type:=msoControlsButton)
            .OnAction = "LowerMacro"
            .FaceId = 100
            .Caption = "Lower Case"
        End With
        With .Controls.Add(Type:=msoControlsButton)
            .OnAction = "ProperMacro"
            .FaceId = 100
            .Caption = "Proper Case"
        End With
    End With
    contextMenu.Controls(4).BeginGroup
            
            
End Sub
Sub DeleteFromCellMenu()
    Dim contextMenu As CommandBar
    Dim ctrl As CommandBar
    Set contextMenu = Application.CommandBars("Cell")
    For Each ctrl In contextMenu.Controls
        If ctrl.Tag = "Thuc hanh Excel" Then
            ctrl.Delete
        End If
    Next
    On Error Resume Next
        contextMenu.FindControl(ID:=3).Delete
    On Error GoTo 0
End Sub
Sub SayIt()
    Application.Speech.Speak Selection.Value
End Sub
