Sub addExcelFileExample1()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim currentPath As String
     
    currentPath = Application.ActiveWorkbook.Path
     
    ' add excel file
    Set wb = Workbooks.Add
 
    ' set ws to sheet1 of wb
    Set ws = wb.Worksheets(1)
     
    ' fill data to column "A1"
    ws.Cells(1, 1) = "Hello VBA!"
     
    ' save file
    With wb
        .SaveAs Filename:=currentPath & "\" & "test1.xlsx"
        .Close
    End With
End Sub



Sub addExcelFileExample2()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim currentPath As String
     
    currentPath = Application.ActiveWorkbook.Path
     
    ' add excel file
    Set wb = Workbooks.Add
 
    ' set ws to sheet1 of wb
    Set ws = wb.Worksheets(1)
     
    ' fill data to column "A1"
    ws.Cells(1, 1) = "Hello VBA!"
     
    ' save file
    With wb
        .SaveAs Filename:=currentPath & "\" & "test2.xlsx", _
                FileFormat:=xlOpenXMLWorkbook
        .Close SaveChanges:=True
    End With
End Sub
