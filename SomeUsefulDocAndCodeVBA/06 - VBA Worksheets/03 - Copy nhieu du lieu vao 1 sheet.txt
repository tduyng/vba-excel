Option Explicit

Public sh As Worksheet
Public rng As Range, shp As Shape, i As Integer, j As Integer
Sub COPY_DATAs_TO_ONE_SHEET()
    Dim sh1 As Worksheet, sh2 As Worksheet, sh3 As Worksheet
    Dim lastRow1 As Integer, lastRow2 As Integer
    Set sh1 = Sheets("chuot")
    Set sh2 = Sheets("ga")
    Set sh3 = Sheets("ngon")
    
    lastRow1 = sh1.Cells(Rows.count, "A").End(xlUp).Row
    lastRow2 = sh2.Cells(Rows.count, "A").End(xlUp).Row + lastRow1 - 1

    
    sh1.Range("Data").Copy Sheet2.Range("A1")
    sh2.Range("Data2").Copy Sheet2.Range("A" & lastRow1)
    sh3.Range("Data3").Copy Sheet2.Range("A" & lastRow2)

End Sub

Function SHEET_EXIST(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    SHEET_EXIST = Not sht Is Nothing
    
End Function

Function CHECK_IF_SHEET_EXIST(SheetName As String) As Boolean
    CHECK_IF_SHEET_EXISTE = False
    Dim ws As Worksheet
    For Each ws In workheets
        If SheetName = ws.Name Then
        CHECK_IF_SHEET_EXIST = True
        Exit Function
        End If
    Next
   
End Function

Function IS_SHEET_EXIST(shName As String) As Boolean
        IS_SHEET_EXIST = Not Evaluate("IsError(" & shName & "!1:1)")
End Function



Sub COPY_MULDATAS_TO_ONE_SHEET()
    Dim shAll As Worksheet
    Dim lastRow As Long, count As Integer
    Dim lastRowSheeti As Long
    If SHEET_EXIST("TONG HOP") = False Then
        Set shAll = Sheets.Add(after:=Worksheets(Worksheets.count))
        shAll.Name = "TONG HOP"
    Else
        Set shAll = Sheets("TONG HOP")
    End If
    count = 0
    lastRow = 0
    For Each sh In Worksheets
        If (sh.Name <> shAll.Name And sh.Name <> "Tirage_v1") Then
            count = count + 1
            If count = 1 Then
                sh.Range("A1:F" & sh.Range("A" & Rows.count).End(xlUp).Row).Copy shAll.Range("A1")
            ElseIf count = 2 Then
                sh.Range("A2:F" & sh.Range("A" & Rows.count).End(xlUp).Row).Copy shAll.Range("A" & lastRow + 1)
            Else
                sh.Range("A2:F" & sh.Range("A" & Rows.count).End(xlUp).Row).Copy shAll.Range("A" & lastRow)
            End If
            lastRow = lastRow + sh.Range("A" & Rows.count).End(xlUp).Row
            
        End If
    Next sh

End Sub
