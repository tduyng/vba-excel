Public Sub selectInputFile()
    Dim fd As Office.FileDialog
 
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
 
      .AllowMultiSelect = False
      .InitialFileName = Application.ActiveWorkbook.path & "\"
      .Filters.Clear
      .Filters.Add "Excel 2007", "*.xlsx"
      .Filters.Add "All Files", "*.*"
 
      If .Show = True Then
        SelectedFile = .SelectedItems(1)
      End If
    End With
End Sub
