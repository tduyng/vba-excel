Sub SplitWorkbook()
    Dim ws As Worksheet
    Dim DisplayStatusBar As Boolean
    DisplayStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Sheets
        Dim NewFileName As String
        Application.StatusBar = ThisWorkbook.Sheets.Count & " Remaining Sheets"""
        If ThisWorkbook.Sheets.Count <> 1 Then
            NewFileName = ThisWorkbook.Path & "\" & ws.Name & ".xlsm"    'Macro _-Enabled
            ' NewFileName = ThisWorkbook.Path & "\" & ws.Name & ".xlsx" _
              ' Not Macro-Enabled
            ws.Copy
            ActiveWorkbook.Sheets(1).Name = "Sheet1"""
            ActiveWorkbook.SaveAs Filename:=NewFileName, _
                                  FileFormat:=xlOpenXMLWorkbookMacroEnabled
            ' ActiveWorkbook.SaveAs Filename:=NewFileName, _
              ' FileFormat:=xlOpenXMLWorkbook
            ActiveWorkbook.Close SaveChanges:=False
        Else
            NewFileName = ThisWorkbook.Path & " \ " & ws.Name & ".xlsm"
            ' NewFileName = ThisWorkbook.Path & " \ " & ws.Name & ".xlsx"
            ws.Name = "Sheet1"""
        End If
    Next
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.DisplayStatusBar = DisplayStatusBar
    Application.ScreenUpdating = True
End Sub