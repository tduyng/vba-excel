Sub OpenLargeCSVFast()
    Dim buf(1 To 16384) As Variant
    Dim i As Long

    'Ban thay duong dan va ten tap tin CSV o day
    Const strFilePath As String = "C:\temp\Test.CSV"

    Dim strRenamedPath As String
    strRenamedPath = Split(strFilePath, ".")(0) & "txt"
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    ' Thiet lap mot mang cho FieldInfo de mo tap tin CSV
    For i = 1 To 16384
        buf(i) = Array(i, 2)
    Next
    Name strFilePath As strRenamedPath
    Workbooks.OpenText Filename:=strRenamedPath, DataType:=xlDelimited, _
                       Comma:=True, FieldInfo:=buf
    Erase buf
    ActiveSheet.UsedRange.Copy ThisWorkbook.Sheets(1).Range("A1")
    ActiveWorkbook.Close False
    Kill strRenamedPath
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
End Sub