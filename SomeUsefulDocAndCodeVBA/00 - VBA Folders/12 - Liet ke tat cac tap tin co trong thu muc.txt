Sub ExcelFileSearch()
    Dim srchExt As Variant, srchDir As Variant, i As Long, j As Long
    Dim strName As String, varArr(1 To 1048576, 1 To 3) As Variant
    Dim strFileFullName As String
    Dim ws As Worksheet
    Dim fso As Object
    Let srchExt = Application.InputBox("Please Enter File Extension", "Info Request")
    If srchExt = False And Not TypeName(srchExt) = "String" Then
        Exit Sub
    End If
    Let srchDir = BrowseForFolderShell
    If srchDir = False And Not TypeName(srchDir) = "String" Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets.Add(Sheets(1))
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("FileSearch Results").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    ws.Name = "FileSearch Results"
    Let strName = Dir$(srchDir & "\*" & srchExt)
    Do While strName <> vbNullString
        Let i = i + 1
        Let strFileFullName = srchDir & strName
        Let varArr(i, 1) = strFileFullName
        Let varArr(i, 2) = FileLen(strFileFullName) \ 1024
        Let varArr(i, 3) = FileDateTime(strFileFullName)
        Let strName = Dir$()
    Loop
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call recurseSubFolders(fso.GetFolder(srchDir), varArr(), i, CStr(srchExt))
    Set fso = Nothing
    ThisWorkbook.Windows(1).DisplayHeadings = False
    With ws
        If i > 0 Then
            .Range("A2").Resize(i, UBound(varArr, 2)).Value = varArr
            For j = 1 To i
                .Hyperlinks.Add anchor:=.Cells(j + 1, 1), Address:=varArr(j, 1)
            Next
        End If
        .Range(.Cells(1, 4), .Cells(1, .Columns.Count)).EntireColumn.Hidden = True
        .Range(.Cells(.Rows.Count, 1).End(xlUp)(2), _
               .Cells(.Rows.Count, 1)).EntireRow.Hidden = True
        With .Range("A1:C1")
            .Value = Array("Full Name", "Kilobytes", "Last Modified")
            .Font.Underline = xlUnderlineStyleSingle
            .EntireColumn.AutoFit
            .HorizontalAlignment = xlCenter
        End With
    End With
    Application.ScreenUpdating = True
End Sub
Private Sub recurseSubFolders(ByRef Folder As Object, _
                              ByRef varArr() As Variant, _
                              ByRef i As Long, _
                              ByRef srchExt As String)
    Dim SubFolder As Object
    Dim strName As String, strFileFullName As String
    For Each SubFolder In Folder.SubFolders
        Let strName = Dir$(SubFolder.Path & "\*" & srchExt)
        Do While strName <> vbNullString
            Let i = i + 1
            Let strFileFullName = SubFolder.Path & "\" & strName
            Let varArr(i, 1) = strFileFullName
            Let varArr(i, 2) = FileLen(strFileFullName) \ 1024
            Let varArr(i, 3) = FileDateTime(strFileFullName)
            Let strName = Dir$()
        Loop
        If i > 1048576 Then Exit Sub
        Call recurseSubFolders(SubFolder, varArr(), i, srchExt)
    Next
End Sub
Private Function BrowseForFolderShell() As Variant
    Dim objShell As Object, objFolder As Object
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "Please select a folder", 0, "C:\")
    If Not objFolder Is Nothing Then
        On Error Resume Next
        If IsError(objFolder.Items.Item.Path) Then
            BrowseForFolderShell = CStr(objFolder)
        Else
            On Error GoTo 0
            If Len(objFolder.Items.Item.Path) > 3 Then
                BrowseForFolderShell = objFolder.Items.Item.Path & _
                                       Application.PathSeparator
            Else
                BrowseForFolderShell = objFolder.Items.Item.Path
            End If
        End If
    Else
        BrowseForFolderShell = False
    End If
    Set objFolder = Nothing: Set objShell = Nothing
End Function