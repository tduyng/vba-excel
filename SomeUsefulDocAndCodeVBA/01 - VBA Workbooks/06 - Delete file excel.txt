Sub deleteExcelFileExample1()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim currentPath As String
    Dim fileName As String
     
    currentPath = Application.ActiveWorkbook.Path
    fileName = currentPath & "\" & "test1.xlsx"
     
     
    If (fso.FileExists(fileName)) Then
        fso.DeleteFile fileName, True
    End If
End Sub


Sub deleteExcelFileExample2()
    Dim currentPath As String
    Dim fileName As String
    Dim test As String
     
    currentPath = Application.ActiveWorkbook.Path
    fileName = currentPath & "\" & "test1.xlsx"
     
    ' delete fileName
    test = Dir(fileName)
    If Not test = "" Then
        Kill (fileName)
    End If
End Sub

Sub Remove_File() 'Xóa file trong folder b?ng VBA
'T?o thông báo xác nh?n xóa
Dim Msg as String
  Msg=MsgBox("Delete Folder", vbYesNo)

'Th?c hi?n vi?c xóa
If Msg=vbYes Then Kill "C:\Users\HYMC\Downloads\Test\*.*"
    'C:\Users\HYMC\Downloads\Test\ là du?ng d?n t?i thu m?c c?n xóa
    '*.* là t?t c? các file c?n xóa (có bao g?m d?u ch?m duôi file)
    
End Sub



Public lCount As Long
Private Sub DeleteFiles(FolderName As String, Search As String, InSub As Boolean)
  Dim FileItem As Object, SubFolder As Object, fle As String
  On Error GoTo ExitSub
  With CreateObject("Scripting.FileSystemObject")
    With .GetFolder(FolderName)
      For Each FileItem In .Files
        If FileItem.Path Like Search Then
          FileItem.Delete
          lCount = lCount + 1
        End If
      Next FileItem
      If InSub Then
        For Each SubFolder In .subFolders
          DeleteFiles SubFolder.Path, Search, True
        Next SubFolder
      End If
    End With
  End With
ExitSub:
End Sub


Sub Main()
  Dim FolderName As String
  FolderName = "D:\Excel\New Folder (2)"
  lCount = 0
  DeleteFiles[COLOR=#ff0000] [B]FolderName[/B][/COLOR][B], [COLOR=#008000]"*.xls"[/COLOR], [COLOR=#0000cd]True[/COLOR][/B]
  MsgBox "Da xoa " & lCount & " files trong thu muc " & FolderName
End Sub
