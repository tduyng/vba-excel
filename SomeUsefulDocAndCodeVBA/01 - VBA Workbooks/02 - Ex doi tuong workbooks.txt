
' 1. kh?i t?o b?ng Workbook hi?n t?i (ThisWorkbook)
Set wb = Application.ThisWorkbook
' 2. Kh?i t?o b?ng Workbook dang ho?t d?ng (ActiveWorkbook)
Set wb = Application.ActiveWorkbook
' 3. Kh?i t?o b?ng cách m? m?t Workbook khác (Workbooks.Open)
Set wb = Application.Workbooks.Open("D:\test\myFile.xlsx", ReadOnly:=True)
' 4. kh?i t?o b?ng cách s? d?ng tên c?a m?t Workbook
Set wb = Application.Workbooks("myFile.xlsx")
' 5. kh?i t?o b?ng cách s? d?ng ch? s? c?a m?t Workbook
Set wb = Application.Workbooks(2)
' 6. t?o ra m?t Workbook m?i
Set wb = Workbooks.Add


' khai báo d?i tu?ng wb
Dim wbInput As Workbook
Dim wbOutput As Workbook
' gán wbInput b?ng Workbook dang ho?t d?ng (ActiveWorkbook)
Set wbInput = Application.ActiveWorkbook
' gán wbOutput b?ng Workbooks.Open d? m? file D:\test\Output.xlsx
Set wbOutput = Application.Workbooks.Open("D:\test\Output.xlsx")

Sub SaveWorkbookExample1()
    Dim wb As Workbook
    Set wb = Workbooks.Add
    wb.Save
End Sub

Sub SaveWorkbookExample1()
    ActiveWorkbook.Save
End Sub


Sub SaveAsWorkbookExample()
    Dim wb As Workbook
    Set wb = Workbooks.Add
    wb.SaveAs Filename:="D:\test\Sample.xlsx"
End Sub



Workbooks(“Workbook Name”).SaveCopyAs([Filename])

Sub WorkbookSaveCopyAsExample()
    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\" & "ver1_" & ThisWorkbook.Name
End Sub



























Option Explicit
 
Public Const SHEET_INPUT = "input"
Public Const SHEET_OUTPUT = "output"
 
Sub ClickButton()
    Call ViDu1
End Sub
 
Sub ViDu1()
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim rowCount As Integer
    Dim colCount As Integer
    Dim i As Integer
    Dim j As Integer
     
    On Error GoTo ErrorProcess
 
    ' assign wb to active workbook
    Set wb = Application.ActiveWorkbook
     
    ' delete sheet "output" if existed
    For i = 1 To wb.Worksheets.Count
        If wb.Sheets(i).Name = SHEET_OUTPUT Then
            Application.DisplayAlerts = False
            wb.Sheets(i).Delete
            Application.DisplayAlerts = True
        End If
    Next i
     
    ' copy sheet "input" and change to "output"
    Set wsInput = wb.Sheets(SHEET_INPUT)
    wsInput.Copy After:=wsInput
    Set wsOutput = wb.Sheets(SHEET_INPUT & " (2)")
    wsOutput.Name = SHEET_OUTPUT
     
    ' count row
    rowCount = wsOutput.Range("A1", wsOutput.Range("A1").End(xlDown)).Rows.Count
    ' count col
    colCount = wsOutput.Range("A1", wsOutput.Range("A1").End(xlToRight)) _
        .Columns.Count
     
    ' if cell value equals empty, assign that cell to 0 value
    For i = 1 To rowCount
        For j = 1 To colCount
            If wsInput.Cells(i, j) = "" Then
                wsOutput.Cells(i, j) = 0
                wsOutput.Cells(i, j).Interior.ColorIndex = 37
            End If
        Next j
    Next i
    GoTo EndSub
ErrorProcess:
    MsgBox Err.Number & ": " & Err.Description
EndSub:
    Set wsOutput = Nothing
    Set wsInput = Nothing
    Set wb = Nothing
End Sub