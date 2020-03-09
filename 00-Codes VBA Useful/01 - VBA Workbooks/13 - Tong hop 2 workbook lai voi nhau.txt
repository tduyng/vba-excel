Sub CombineWorkbooks()
    Dim CurFile As String, DirLoc As String
    Dim DestWB As Workbook
    Dim ws As Object    'Nh?m cho phép nhi?u lo?i sheet
    ' Ðu?ng d?n d?n các t?p tin
    DirLoc = ThisWorkbook.Path & "\tst\"   
    CurFile = Dir(DirLoc & "*.xls*")
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set DestWB = Workbooks.Add(xlWorksheet)
    Do While CurFile <> vbNullString
        Dim OrigWB As Workbook
        Set OrigWB = Workbooks.Open(Filename:=DirLoc & CurFile, ReadOnly:=True)
        ' Limit to valid sheet names and removes .xls*
        CurFile = Left(Left(CurFile, Len(CurFile) - 5), 29)
        For Each ws In OrigWB.Sheets
            ws.Copy After:=DestWB.Sheets(DestWB.Sheets.Count)
            If OrigWB.Sheets.Count > 1 Then
                DestWB.Sheets(DestWB.Sheets.Count).Name = CurFile & ws.Index
            Else
                DestWB.Sheets(DestWB.Sheets.Count).Name = CurFile
            End If
        Next
        OrigWB.Close SaveChanges:=False
        CurFile = Dir
    Loop
    Application.DisplayAlerts = False
    DestWB.Sheets(1).Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Set DestWB = Nothing
End Sub