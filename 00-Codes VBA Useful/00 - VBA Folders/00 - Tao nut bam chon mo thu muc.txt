Sub ChonMoFileVBA()
    Dim TenFile As Long
    'Mo thuoc tinh File Open
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Show
        ' Hien thi duong dan cua file duoc chon
        For TenFile = 1 To .SelectedItems.Count
            MsgBox .SelectedItems(TenFile)
        Next TenFile
    End With
End Sub