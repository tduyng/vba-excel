Sub GopFileExcel()
Dim FilesToOpen
Dim x As Integer
On Error GoTo ErrHandler
Application.ScreenUpdating = False
FilesToOpen = Application.GetOpenFilename _
(FileFilter:=”Microsoft Excel Files (*.xlsx), *.xlsx”, MultiSelect:=True, Title:=”Files to Merge”)
If TypeName(FilesToOpen) = “Boolean” Then
MsgBox “No Files were selected”
GoTo ExitHandler
End If
x = 1
While x <= UBound(FilesToOpen)
Workbooks.Open Filename:=FilesToOpen(x)
Sheets().Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
x = x + 1
Wend
ExitHandler:
Application.ScreenUpdating = True
Exit Sub
ErrHandler:
MsgBox Err.Description
Resume ExitHandler
End Sub