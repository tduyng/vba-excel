Dim strFile As String
Dim WB As Workbook
strFile = Trim(TextBox1.Value)
Dim DirFile As String
If Len(strFile) = 0 Then Exit Sub

DirFile = "C:\Documents and Settings\Administrator\Desktop\" & strFile
If Len(Dir(DirFile)) = 0 Then
  MsgBox "File does not exist"
Else
 On Error Resume Next
 Set WB = Workbooks.Open(DirFile)
 On Error GoTo 0
 If WB Is Nothing Then MsgBox DirFile & " is invalid", vbCritical
End If



Function IsFile(ByVal fName As String) As Boolean
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function


Private Function FileExists(fname) As Boolean
'   Returns TRUE if the file exists
    Dim x As String
    x = Dir(fname)
    If x <> "" Then FileExists = True _
        Else FileExists = False
End Function
