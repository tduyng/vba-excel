Function getLastRowCol(choice As Long, sh As Worksheet, Optional InputCol As Variant, Optional InputRw As Long)
' 1 = getLastRowCol row
' 2 = getLastRowCol column
' 3 = getLastRowCol cell
' 4 = getLastRow of column given
' 5 = getLastCol of row given
    Dim lRw&, lCol&, i&
    Dim s As String, s1 As String

    With sh
        Select Case choice
        Case 1:
            On Error Resume Next
            getLastRowCol = .Cells.Find(What:="*", _
                            After:=.Cells(1, 1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
            On Error GoTo 0
    
        Case 2:
            On Error Resume Next
            getLastRowCol = .Cells.Find(What:="*", _
                            After:=.Cells(1, 1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            On Error GoTo 0
    
        Case 3:
            On Error Resume Next
            lRw = .Cells.Find(What:="*", _
                           After:=.Cells(1, 1), _
                           LookAt:=xlPart, _
                           LookIn:=xlFormulas, _
                           SearchOrder:=xlByRows, _
                           SearchDirection:=xlPrevious, _
                           MatchCase:=False).Row
            lCol = .Cells.Find(What:="*", _
                            After:=.Cells(1, 1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            getLastRowCol = Cells(lRw, lCol).Address(False, False)
            If Err.Number > 0 Then
                getLastRowCol = .Cells(1, 1).Address(False, False)
                Err.Clear
            End If
            On Error GoTo 0
        Case 4:
            On Error Resume Next
            If Not IsNumeric(InputCol) Then
                InputCol = sh.Columns(InputCol).Column
            End If
           
            getLastRowCol = .Columns(InputCol).Find(What:="*", _
                            After:=.Cells(1, InputCol), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
            On Error GoTo 0
        Case 5
           On Error Resume Next
          
            getLastRowCol = .Rows(InputRw).Find(What:="*", _
                            After:=.Cells(InputRw, 1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
            On Error GoTo 0
            
            
        End Select
    End With
End Function


'test lastrow cua 1 range
Sub test1()
    Dim arr()
    Dim rng As Range
    Dim countRw&, lstRwC&
    Set rng = ThisWorkbook.Sheets("MENU").Range("DIR_SER")
    rng.Select
    'Count Rw: dung cho tinh toan ca voi o trong vide
    countRw = rng.Count / rng.Column - rng.Resize(1, 1).Row + 1
    
    'Ta test voi gia tri cuoi cung cua range nay
    lstRwC = rng.End(xlDown).Row
    
    'Test gia tri cuoi cung cua cot dau tien
    lstRwC = rng.Resize(countRw, 1).End(xlDown).Row
    
    'Test gia tri cuoi cung cua cot 2
    lstRwC = rng.Offset(0, 1).Resize(countRw, 1).End(xlDown).Row
    arr = rng.Resize(countRw, 1)
End Sub



Sub DONG_CUOI_CUNG()
Dim sh As Worksheet
Dim rng As Range, lastRow As Integer
Set sh = Worksheets("Tirage")
'Tim dong cuoi cung su dung cua mot sheet
lastRow = sh.UsedRange.Rows.Count

'Source: https://blog.hocexcel.online/tim-dong-cuoi-cua-bang-table-trong-excel.html



'*********************************************************************************************
'*********************************************************************************************
'Phuong phap End(xlup) thong thuong

lastRow = sh.Cells(Rows.Count, "H").End(xlUp).Row

'----------------------------------------------------------------------------------------------
'Rows.Count: dem so hang
'End(xlUP).Row: tinh tu toan bo song cua bang tinh, tinh nguoc len phia tren toi dong gan nhat
'Viec ghi nhan dong cuoi theo cach nay duoc hieu la toi dng cuoi cua bang du lieu dang ap dung
'Khong phu thuoc vaof pia trong no co du lieu hay khong
'----------------------------------------------------------------------------------------------

'*********************************************************************************************
'*********************************************************************************************




'*********************************************************************************************
'*********************************************************************************************
'Phuong phap End(xlup), End(xlDown) dung rieng cho table
Debug.Print sh.Cells(Rows.Count, "H").End(xlUp).Row
Debug.Print sh.Range("TableTest").End(xlUp).Row
Debug.Print sh.Range("TableTest").End(xlDown).Row

'----------------------------------------------------------------------------------------------
'Rows.Count: dem so hang
'End(xlUP).Row: cho dong dau tien can tim trong table
'End(Down).Row: Cho dong cuoi cung can tim tron table
'Tuy nhien neu dong cuoi cos du lieu bi an di(chuc nang nang Hide Row) thi cau lenh nay se khong tinh
'cac dong bi an, va neu co dong nam xe giua cac noi dung trong bang thi cach nay chi tim den
'vi tri dong trong xen giua dong do
'----------------------------------------------------------------------------------------------

'*********************************************************************************************
'*********************************************************************************************






'*********************************************************************************************
'*********************************************************************************************
'Phuong phap Range.Find
'Day la phong phap duoc nhieu nguowif cos kinh nghiem su dung VBA ua thich
'Boi no khac phuc duoc hau het cac nhuwoc diem cua cac phuong phap tim dong cuoi khac

On Error Resume Next
Debug.Print sh.Range("Table2").Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
Debug.Print sh.Range("Table2").Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
On Error Resume Next
'----------------------------------------------------------------------------------------------
'Co the su dung doi tuong tim kiem la value, hoac formulas
'----------------------------------------------------------------------------------------------



'
' last row containing data (using Find in formulas)
'
'jLastRangeData = oSht.Range("NamedRange").Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
'
' last visible row containing data (using Find in values)
'
'jLastVisibleRangeData = oSht.Range("NamedRange").Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
'
' last row containing data (using Find in formulas)
'
'jLastTableData = oSht.Range("Table1").Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
'
' last visible row containing data (using Find in values)
'
'jLastVisibleTableData = oSht.Range("Table1").Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

'*********************************************************************************************
'*********************************************************************************************




End Sub



Sub Macro1
    Sheets("Sheet1").Select
    MsgBox "The last row found is: " & Last(1, ActiveSheet.Cells)
    MsgBox "The last column (R1C1) found is: " & Last(2, ActiveSheet.Cells)
    MsgBox "The last cell found is: " & Last(3, ActiveSheet.Cells)
    MsgBox "The last column (A1) found is: " & Last(4, ActiveSheet.Cells)
End Sub


