Option Explicit
Sub Exp2Pdf()
Application.ScreenUpdating = False
Application.EnableEvents = False
Dim cRng As Range, DATA As Variant, I As Long
Dim sPath As String, sName As String, sDate As Variant
DATA = Sheet1.Range("A2:I" & Sheet1.Cells(Sheet1.Rows.Count, "A").End(xlUp).Row).Value
On Error GoTo Errhander
For Each cRng In Sheet2.Range("G1:G" & Sheet2.Cells(Sheet2.Rows.Count, "G").End(xlUp).Row)
    Sheet2.Range("E4").Value = cRng.Value
    sPath = vbNullString:    sName = vbNullString
    sPath = ThisWorkbook.Path
    If sPath = vbNullString Then sPath = Application.DefaultFilePath
    sPath = sPath & "\"
    sDate = Application.WorksheetFunction.VLookup(cRng.Value, DATA, 2, 0)
    On Error Resume Next
    If CDate(sDate) Then
       If Err.Number <> 0 Then GoTo Errhander
       sName = CStr(Format(CDate(sDate), "dd-mm-yyyy")) & cRng.Value & ".Pdf"
    End If
    sPath = sPath & sName
    Sheet2.ExportAsFixedFormat _
                 xlTypePDF, _
                 sPath, _
                 xlQualityStandard, _
                 True, _
                 False, , , False
Next
Application.ScreenUpdating = True
Application.EnableEvents = True
Exit Sub
Errhander:
Application.ScreenUpdating = True
Application.EnableEvents = True
MsgBox "Error founded ! Nothing Pdf file created"
Exit Sub
End Sub