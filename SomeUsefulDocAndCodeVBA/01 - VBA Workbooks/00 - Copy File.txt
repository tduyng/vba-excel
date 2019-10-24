Option Explicit

Sub SELECT_TEMPLATE()
    Sheet4.Range("A1") = Application.GetOpenFilename()
End Sub

Sub SELECT_DESTIANATION_FOLDER()
    Dim dialog As FileDialog
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    dialog.AllowMultiSelect = False
    dialog.Show
    Sheet4.Range("A2") = dialog.SelectedItems(1)
End Sub

Sub COPY_FILE_EXCEL()
    Dim rngi As Range, ext As String, template As String, dest_folder As String
    Dim count As Integer
    template = Sheet4.Range("A1").Value
    dest_folder = Sheet4.Range("A2").Value
    ext = Split(template, ".")(UBound(Split(template, ".")))
    count = 0
    For Each rngi In Sheet4.Range("B2:B16")
        FileCopy template, dest_folder & "\" & rngi.Value & "." & ext
        count = count + 1
    Next rngi
    MsgBox "Done!" + vbNewLine + "You have successed to copy " & count & " new files"
End Sub
