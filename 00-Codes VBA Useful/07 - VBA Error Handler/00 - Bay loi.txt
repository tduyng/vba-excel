Sub CAN_BAC_HAI()

Dim number As Variant
TINHCAN:
number = InputBox("Ban hay nhap mot so de tinh can", "Can bac hai")

If number = "" Then Exit Sub
If number < 0 Then
    MsgBox "Ban phai nhap so duong > 0", vbOKOnly
    GoTo TINHCAN
    'Exit Sub
End If
If Not IsNumeric(number) Then
    MsgBox "Du lieu nhap phai dang so", vbOKOnly
    GoTo TINHCAN
    'Exit Sub
End If

MsgBox Round(Sqr(number), 3)

End Sub



Sub CAN_BAC_HAI2()

    Dim number As Variant, thulai As Variant
    
TINHCAN:
    
    number = InputBox("Hay nhap du lieu de tinh can")
    
    On Error GoTo LOI
    If number = "" Then Exit Sub
    MsgBox "Ket qua la: " & Round(Sqr(number), 3), vbOKOnly
    Exit Sub
LOI:
    MsgBox "co loi xay ra voi du lieu ban da nhap", vbOKOnly
    thulai = MsgBox("Co tiep tuc khong", vbYesNo)
    If thulai = vbYes Then Resume TINHCAN
    
End Sub
