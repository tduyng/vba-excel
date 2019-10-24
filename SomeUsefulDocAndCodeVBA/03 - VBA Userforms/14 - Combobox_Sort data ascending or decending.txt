
Public Sub SORT_LIST(cbx As ComboBox)
    Dim itemTemp As Variant
    Dim x&, y&
    
    With cbx
    For x = LBound(.list) To UBound(.list)
        For y = x To UBound(.list)
            If IsNumeric(.list(y, 0)) And IsNumeric(.list(x, 0)) Then
                If (.list(y, 0) + 0) < (.list(x, 0) + 0) Then
                    itemTemp = .list(x, 0)
                    .list(x, 0) = .list(y, 0)
                    .list(y, 0) = itemTemp
                End If
            Else
                If .list(y, 0) < .list(x, 0) Then
                    itemTemp = .list(x, 0)
                    .list(x, 0) = .list(y, 0)
                    .list(y, 0) = itemTemp
                End If
            End If
        Next y
'        .AddItem .list(x, 0)
    Next x
    End With
End Sub



Option Explicit

Private Sub cbSort_Click()

    Call SORT_LIST(ufSortCbx.cbx1)
    Call SORT_LIST(ufSortCbx.cbx2)

End Sub


Private Sub UserForm_Initialize()
    cbx1.list = Sheets("VEIF05").Range("A2:A27").Value
    cbx2.list = Sheets("VEIF05").Range("B2:B27").Value
End Sub

