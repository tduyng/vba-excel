Public Function Coleter(ColIndex As Long) As String
    Coleter = Replace(Cells(1, ColIndex).Address(0, 0), 1, "")
End Function