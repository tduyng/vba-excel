Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
  Dim sht As Worksheet

  If wb Is Nothing Then Set wb = ThisWorkbook
  On Error Resume Next
  Set sht = wb.Sheets(shtName)
  On Error GoTo 0
  SheetExists = Not sht Is Nothing
End Function



Function CheckIfSheetExists(SheetName As String) As Boolean
  CheckIfSheetExists = False
  For Each WS In Worksheets
    If SheetName = WS.name Then
      CheckIfSheetExists = True
      Exit Function
    End If
  Next WS
End Function


Function isSheetExist(shName As String) As Boolean
  isSheetExist = Not Evaluate("IsError(" & shName & "!1:1)")
End Function


Private Function SheetExists(sname) As Boolean
'   Returns TRUE if sheet exists in the active workbook
    Dim x As Object
    On Error Resume Next
    Set x = ActiveWorkbook.Sheets(sname)
    If Err = 0 Then SheetExists = True _
        Else SheetExists = False
End Function