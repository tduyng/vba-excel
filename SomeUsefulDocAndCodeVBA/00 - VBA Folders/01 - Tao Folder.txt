Sub CREATE_FOLDER()
Set sh = Worksheets("xTest")
Dim filePath, fileName1, fileName2, fileName3 As String
filePath = Application.ActiveWorkbook.Path


fileName1 = filePath & "\" & "File1"
fileName2 = filePath & "\" & "File2"
fileName3 = filePath & "\" & "File3"

Debug.Print Len(Dir(fileName1, vbDirectory))
Debug.Print Len(Dir(fileName2, vbDirectory))
Debug.Print Len(Dir(fileName3, vbDirectory))

If Len(Dir(fileName1, vbDirectory)) = 0 & Len(Dir(fileName2, vbDirectory)) = 0 & Len(Dir(fileName3, vbDirectory)) = 0 Then

'Folder co ten File1 can tao chua ton tai
MkDir fileName1
MkDir fileName2
MkDir fileName3

On Error Resume Next
Else
'Tap tin can tao da ton tai
MsgBox "Trong tap tin cua ban da ton tai 3 folder co ten la " & vbNewLine & "File1" & vbNewLine & "File2" & vbNewLine & "File3", vbCritical

        
End If

End Sub




Sub CREATE_FOLDER_WITH_VALUE_EXCEL()
Dim fileName As Variant, c As Variant

Set sh = Worksheets("xTest")
Set rng = sh.Range("E7:E24")
'Debug.Print sh.UsedRange.Rows.Count + sh.UsedRange.Rows.Row - 1
'Debug.Print rng.Columns.Column
'Debug.Print sh.Cells(sh.Rows.Count, rng.Columns.Column).End(xlUp).Row
For Each c In rng
    If Len(Dir(c.Value, vbDirectory)) = 0 Then
    MkDir c.Value
    On Error Resume Next
    End If
Next c

End Sub



Sub createFolderOutput(path As String, folderName As String)
    Dim outputPath As String
    Dim objFso As FileSystemObject
     
    Set objFso = New FileSystemObject
     
    outputPath = objFso.BuildPath(path, folderName)
     
    If objFso.FolderExists(outputPath) = False Then
        objFso.CreateFolder outputPath
    End If
     
    Set objFso = Nothing
End Sub
