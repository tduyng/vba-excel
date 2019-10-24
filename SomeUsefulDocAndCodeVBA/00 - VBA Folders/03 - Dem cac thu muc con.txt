Sub CountFiles()
Dim xFolder As String
Dim xPath As String
Dim xCount As Long
Dim xFiDialog As FileDialog
Dim xFile As String
Set xFiDialog = Application.FileDialog(msoFileDialogFolderPicker)
If xFiDialog.Show = -1 Then
xFolder = xFiDialog.SelectedItems(1)
End If
If xFolder = "" Then Exit Sub
xPath = xFolder & "\*.xlsx"
xFile = Dir(xPath)
Do While xFile <> ""
xCount = xCount + 1
xFile = Dir()
Loop
MsgBox xCount & " files found"
End Sub



Sub COUNT_FILES()
    Dim xFolder As String, xFileName As String
    Dim xFileDialog As FileDialog
    Dim xCount As Integer, xListFileName As String
    
    'On Error Resume Next
    Set xFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With xFileDialog
        .AllowMultiSelect = False
        .Title = "Choose a folder"
        .Show
    End With
    xFolder = xFileDialog.SelectedItems(1)
    If xFolder = "" Then Exit Sub
    xFileName = Dir(xFolder & "\*.xl*")
    Do While xFileName <> ""
        xCount = xCount + 1
        xListFileName = xListFileName + vbNewLine + xFileName
        xFileName = Dir()
    Loop
    MsgBox "You have " & xCount & " files Excel found" + vbNewLine + _
            "with the names:" + vbNewLine + xListFileName
End Sub
