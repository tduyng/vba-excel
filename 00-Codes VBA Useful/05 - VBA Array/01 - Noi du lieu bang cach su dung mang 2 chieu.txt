Sub FILTER_DATA()
Dim T1 As Long, T2 As Long, last_Column As Integer
Dim Data As Variant, i As Long, j As Long, TMP As Variant, MA As String, T As Double
Data = SHEET1.Range("A2:C" & SHEET1.Cells(Rows.Count, "C").End(xlUp).Row).Value2

T1 = 0
T1 = GetTickCount
j = 0
ReDim TMP(1 To 3, 1 To 1)
For i = 1 To UBound(Data, 1)
    MA = CStr(Data(i, 1) & Data(i, 2))
    If MA <> vbNullString Then
        j = j + 1
        ReDim Preserve TMP(1 To 3, 1 To j)
        TMP(1, j) = Data(i, 1)
        TMP(2, j) = Data(i, 2)
        TMP(3, j) = Data(i, 3)
    Else
        TMP(3, j) = TMP(3, j) & " ; " & Data(i, 3)
    End If
    
Next
Sheet2.Range("A2:C1048576").ClearContents
'On Error Resume Next
Sheet2.Activate
ReDim Preserve TMP(1 To 3, 1 To j)
'Debug.Print j
Sheet2.Range("A2").Resize(3, UBound(TMP, 2)).Value = TMP
Sheet2.Range("A2:" & Sheet2.Cells(4, Columns.Count).End(xlToLeft).Address).Copy
Sheet2.Range("D10").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

'Debug.Print last_Column
'Sheet3.Range("A2").Resize(UBound(TMP, 2), 3).Value = Application.WorksheetFunction.Transpose(TMP)
T2 = GetTickCount
MsgBox ((T2 - T1) / 1000), vbInformation, "So giay thuc hien la "
 'On Error Resume Next
End Sub