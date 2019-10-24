Public Function openExcelFile(ByVal openFileName As String, _
        Optional ByVal readOnly As Boolean) As Workbook
    readOnly = (readOnly And True)
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=openFileName, readOnly:=readOnly
    Set openExcelFile = ActiveWorkbook
End Function


Sub openExcelFileExample1()
    Dim wb As Workbook
     
    ' open excel file
    Set wb = Workbooks.Open("D:\test\Sample.xlsx")
     
    ' read file
     
    ' close excel file without save changes
    wb.Close SaveChanges:=False
End Sub


Sub openExcelFileExample2()
    Dim wb As Workbook
    Dim ws As Worksheet
     
    ' open excel file
    Set wb = Workbooks.Open(Filename:="D:\test\Sample.xlsx")
     
    ' set ws to sheet1 of wb
    Set ws = wb.Worksheets(1)
     
    ' fill data to column "A1"
    ws.Cells(1, 1) = "Hello VBA!"
     
    ' close excel file with save changes
    wb.Close SaveChanges:=True
End Sub


Sub openExcelFileExample3()
    Dim wb As Workbook
    Dim ws As Worksheet
     
    ' open excel file
    Set wb = Workbooks.Open(Filename:="D:\test\Sample.xlsx", ReadOnly:=True)
     
    ' read file
    ' write file
         
    ' set ws to sheet1 of wb
    Set ws = wb.Worksheets(1)
     
    ' fill data to column "A1"
    ws.Cells(1, 1) = "Hello VBA!"
     
    ' close excel file without save changes
    wb.Close SaveChanges:=False
End Sub
