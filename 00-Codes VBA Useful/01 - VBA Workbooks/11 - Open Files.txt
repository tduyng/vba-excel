Sub OpenFile_Excel()
    '    Dim fd As Office.FileDialog
    Dim SelectedFile As String
    Dim ckx_IsOpenWb As Boolean
'    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With Application.FileDialog(msoFileDialogFilePicker)
             .AllowMultiSelect = False
             .InitialFileName = Application.ActiveWorkbook.Path & "\"
             .Filters.Clear
             .Filters.Add "Files Excel", "*.xl*;*.xlsx;*.xlsm;*.xlsb;*.xlam;*.xltx,*.xltm,*.xla,*.xlt,*.xlm,*.xlw"
		'.Filters.Add "Files Excel", "*.xl*;*.xlsx;*.xlsm;*.xlsb;*.xls"
             .Filters.Add "All Files", "*.*"
        
             If .Show = True Then
               SelectedFile = .SelectedItems(1)
             End If
    End With
    If SelectedFile = vbNullString Then
        Exit Sub
    End If
    On Error Resume Next
    ckx_IsOpenWb = IsWorkBookOpen(SelectedFile)
    If ckx_IsOpenWb Then
        'already open
        Set wbSQ = Workbooks(getNameOfPath(SelectedFile)) 'workbook suivi
    Else
        Set wbSQ = Workbooks.Open(SelectedFile) 'workbook suivi
    End If
    On Error GoTo 0
End Sub


Function IsWorkBookOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function