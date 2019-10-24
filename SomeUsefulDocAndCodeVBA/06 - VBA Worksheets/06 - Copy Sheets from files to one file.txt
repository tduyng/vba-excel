Sub GET_SHEETS()
    Dim filePath As String, wb As Workbook, sht As Worksheet
    Dim fileName As String
    Dim fso As New Scripting.FileSystemObject
    filePath = "E:\00 - BIM\07 - EXCEL\00 - Code VBA useful\01 - Thao tac luu, tach, tao moi files excel\"
    fileName = Dir(filePath & "*.xl*")
    Application.ScreenUpdating = False
    Do While fileName <> ""
        Set wb = Workbooks.Open(fileName:=filePath & fileName, ReadOnly:=True)
        For Each sht In Worksheets
            sht.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = fso.GetbaseName(wb.Name) & "-" & sht.Name
        Next sht
        wb.Close
        fileName = Dir()
    Loop
    Application.ScreenUpdating = True
    'MsgBox "Done!"
End Sub

Sub DELETE_SHEETS()
    Dim sht As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each sht In Worksheets
        If sht.Name <> "Feuil1" Then sht.Delete
    Next
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Function GET_BASE_NAME(fileName As String) As String
    If InStr(fileName, ".") > 0 Then
        GET_BASE_NAME = Left(fileName, (InStr(fileName, ".") - 1))
    End If
End Function


Sub GET_SHEETS2()
    Dim filePath As String, wb As Workbook, sht As Worksheet
    Dim fileName As String
    Dim fso As New Scripting.FileSystemObject
    filePath = "E:\00 - BIM\07 - EXCEL\00 - Code VBA useful\01 - Thao tac luu, tach, tao moi files excel\"
    fileName = Dir(filePath & "*.xl*")
    Application.ScreenUpdating = False
    Do While fileName <> ""
        Set wb = Workbooks.Open(fileName:=filePath & fileName, ReadOnly:=True)
        For Each sht In Worksheets
            sht.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = GET_BASE_NAME(wb.Name) & "-" & sht.Name
        Next sht
        wb.Close
        fileName = Dir()
    Loop
    Application.ScreenUpdating = True
    'MsgBox "Done!"
End Sub
