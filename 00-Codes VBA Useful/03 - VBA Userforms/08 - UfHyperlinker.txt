Option Explicit

Private Sub cbBrowse_Click()
    Dim diafolder As FileDialog
    Set diafolder = Application.FileDialog(msoFileDialogFolderPicker)
    diafolder.AllowMultiSelect = False
    diafolder.Show
    txtPath = diafolder.SelectedItems(1)
    
    Set diafolder = Nothing
End Sub


Private Sub lbFiles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim fileName, path, link As String
    fileName = lbFiles.List(lbFiles.ListIndex)
    path = txtPath & "\"
    link = path & fileName
    ActiveSheet.Hyperlinks.Add anchor:=Selection, Address:=link, TextToDisplay:=Selection.value
End Sub

Private Sub txtPath_Change()
    lbFiles.List = LIST_MY_FILES(txtPath.Text)
     
End Sub
Private Function LIST_MY_FILES(dirPath As String) As String()
    Dim myFiles As String, counter As Long
    Dim DirArray() As String
    ReDim DirArray(1000)
    myFiles = Dir$(dirPath & "\*.*")
    Do While myFiles <> ""
        DirArray(counter) = myFiles
        myFiles = Dir$
        counter = counter + 1
    Loop
    ReDim Preserve DirArray(counter - 1)
    LIST_MY_FILES = DirArray
End Function

Private Sub UserForm_Click()

End Sub
