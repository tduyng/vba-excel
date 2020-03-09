Function Quy_doi(k As Long) As String
    Dim vArr As Variant
    vArr = Split(Cells(1, k).Address(True, False), "$")
    Quy_doi = vArr(0)
End Function


'Hoac sd cong thuc
'=CHAR(64+ROW())