Private Sub cbModal_Change()
    mclsFormChanger.Modal = cbModal.Value
    CheckEnabled
End Sub

Private Sub cbSizeable_Change()
    mclsFormChanger.Sizeable = cbSizeable.Value

    CheckBorderStyle
End Sub

Private Sub cbCaption_Change()
    mclsFormChanger.ShowCaption = cbCaption.Value

    CheckBorderStyle
    CheckEnabled
End Sub

Private Sub cbSmallCaption_Change()
    mclsFormChanger.SmallCaption = cbSmallCaption.Value
    CheckEnabled
End Sub

Private Sub cbTaskBar_Change()
    mclsFormChanger.ShowTaskBarIcon = cbTaskBar.Value
    CheckEnabled
End Sub

Private Sub cbSysmenu_Change()
    mclsFormChanger.ShowSysMenu = cbSysmenu.Value
    CheckEnabled
End Sub

Private Sub cbIcon_Change()
    mclsFormChanger.ShowIcon = cbIcon.Value
    If cbIcon.Value And mclsFormChanger.IconPath = "" Then btnChangeIcon_Click
    CheckEnabled
End Sub

Private Sub btnChangeIcon_Click()

    Dim vFile As Variant

    vFile = Application.GetOpenFilename("Icon files (*.ico;*.exe;*.dll),*.ico;*.exe;*.dll", 0, "Open Icon File", "Open", False)

    'Showing dialog sets the form modeless, so check it
    mclsFormChanger.Modal = cbModal

    If vFile = False Then Exit Sub

    mclsFormChanger.IconPath = vFile

End Sub

Private Sub cbCloseBtn_Change()
    mclsFormChanger.ShowCloseBtn = cbCloseBtn.Value
    CheckEnabled
End Sub

Private Sub cbMinimize_Change()
    mclsFormChanger.ShowMinimizeBtn = cbMinimize.Value
    CheckEnabled
End Sub

Private Sub cbMaximize_Change()
    mclsFormChanger.ShowMaximizeBtn = cbMaximize.Value
    CheckEnabled
End Sub

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub CheckBorderStyle()

    'If the userform is not sizeable and doesn't have a caption,
    'Windows draws it without a border, and we need to apply our
    'own 3D effect.
    If Not (cbSizeable Or cbCaption) Then
        Me.SpecialEffect = fmSpecialEffectRaised
    Else
        Me.SpecialEffect = fmSpecialEffectFlat
    End If

End Sub

Private Sub CheckEnabled()

    'Without a system menu, we can't have the close, max or min buttons
    cbSysmenu.Enabled = cbCaption
    cbCloseBtn.Enabled = cbSysmenu And cbCaption
    cbIcon.Enabled = cbSysmenu And cbCaption And Not cbSmallCaption
    cbMaximize.Enabled = cbSysmenu And cbCaption And Not cbSmallCaption
    cbMinimize.Enabled = cbSysmenu And cbCaption And Not cbSmallCaption

    btnChangeIcon.Enabled = cbIcon.Value And cbIcon.Enabled

End Sub