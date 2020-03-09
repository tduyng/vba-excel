Private Sub Workbook_Open()
    On Error Resume Next
    X = Workbooks(“Code.xlsm”).Name
    If Not Err = 0 Then
        On Error GoTo 0
        Workbooks.Open Filename:= _
                       ThisWorkbook.Path & Application.PathSeparator & “Code.xlsm”
    End If
    On Error GoTo 0
    Application.Run “Code.xlsm!CustFileOpen”
End Sub