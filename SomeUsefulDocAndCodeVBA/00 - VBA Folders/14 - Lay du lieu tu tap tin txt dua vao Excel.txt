Sub ReadTxtLines()
'Do dung khai bao tre, nen khong can phai tham chieu den thu vien Scripting Runtime trong reference
    Dim sht As Worksheet
    Dim fso As Object
    Dim fil As Object
    Dim txt As Object
    Dim strtxt As String
    Dim tmpLoc As Long
    'Lam viec voi Sheet hien tai
    Set sht = ActiveSheet
    'Xoa du lieu tren Sheet hien tai
    sht.UsedRange.ClearContents
    'Tao doi tuong File system chung ta can de lam viec voi tap tin
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Tap tin ban can mo
    Set fil = fso.GetFile("c:\service.log")
    'Mo tap tin nhu la mot  TextStream
    Set txt = fil.OpenAsTextStream(1)
    'Doc noi dung va dua vao bien kieu chuoi
    strtxt = txt.ReadAll
    'Dong textstream
    txt.Close
    'Tim ky tu xuong hang dau tien
    tmpLoc = InStr(1, strtxt, vbCrLf)
    'Dung vong lap de quet qua tung hang
    Do Until tmpLoc = 0
        sht.Cells(sht.Rows.Count, 1).End(xlUp).Offset(1).Value = _
        Left(strtxt, tmpLoc - 1)
        strtxt = Right(strtxt, Len(strtxt) - tmpLoc - 1)
        tmpLoc = InStr(1, strtxt, vbCrLf)
    Loop
    sht.Cells(sht.Rows.Count, 1).End(xlUp).Offset(1).Value = strtxt
    ' Giai phong bo nho
    Set fso = Nothing
End Sub