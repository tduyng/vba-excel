Sub TrangHoang()
Dim Dic As Object, KQ As Range
Dim TmpVao As Variant, TmpRa As Variant, Tmp As Variant, I As Long, J As Long
  Set Dic = CreateObject("Scripting.Dictionary") 'Kh?i t?o dictionary
With Sheet1
TmpRa = .Range("A3:A" & .Cells(.Rows.Count, "A").End(xlUp).Row).Value 'D? li?u s?n có ? c?t A
TmpVao = .Range("G3:G" & .Cells(.Rows.Count, "G").End(xlUp).Row).Value ' D? li?u ? C?t G
Set KQ = .Range("A" & .Cells(.Rows.Count, "A").End(xlUp).Row).Offset(1)' Ô c?n g?n k?t qu?
End With
For I = LBound(TmpRa, 1) To UBound(TmpRa, 1)' Vòng l?p ch?y qua c?t A
If Not Dic.exists(CStr(TmpRa(I, 1))) Then Dic.Add CStr(TmpRa(I, 1)), I' Add nh?ng ph?n t? không trùng vào Dic
Next
ReDim Tmp(0)' T?o m?ng m?t chi?u m?i d? luu k?t qu? sau khi d?i chi?u
J = -1 ' Do m?ng m?t chi?u b?t d?u t? ph?n t? 0 nên J = -1
For I = LBound(TmpVao, 1) To UBound(TmpVao, 1) ' Vòng l?p ch?y qua c?t G
If Not Dic.exists(CStr(TmpVao(I, 1))) Then ' N?u mã chua t?n t?i trong Dic
J = J + 1' Ðây là lý do t?i sao J = -1 ? dòng trên
ReDim Preserve Tmp(J)' M? r?ng m?ng và không làm m?t d? li?u cu
Tmp(J) = TmpVao(I, 1)'Ðua mã dã có trong c?t G mà chua có trong c?t A vào m?ng k?t qu?
    End If
Next

KQ.Resize(UBound(Tmp, 1) + 1, 1) = Application.Transpose(Tmp)' Dán k?t qu? vào b?ng tính
End Sub