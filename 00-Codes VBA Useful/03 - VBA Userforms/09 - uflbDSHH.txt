Option Explicit
Dim Dic As Object

Private Sub cbInsert_Click()
    Dim arr As Variant
    Dim lastRow As Long, i As Long
    Sheet1.Activate
    arr = Dic.items
    lastRow = Sheet1.Range("A" & Rows.Count).End(xlUp).Row + 1
    
    For i = 0 To Dic.Count - 1
        Sheet1.Cells(lastRow + i, "A").Resize(, 4) = Split(arr(i), ";")
        'Debug.Print Dic.Count - 1
    Next i
    Set Dic = Nothing
    Erase arr
    Unload Me
End Sub

Private Sub lbxMHH_Change()
    Dim id As Variant
    id = lbxMHH.ListIndex
    With Me.lbxMHH
        If .Selected(id) Then
            If Not Dic.Exists(.List(id, 0)) Then
                Dic.Add .List(id, 0), .List(id, 0) & ";" & .List(id, 1) & ";" & .List(id, 2) & ";" & .List(id, 3)
            End If
        Else
            Dic.Remove (.List(id, 0))
        End If
    End With
End Sub

Private Sub txtMHH_Change()

End Sub

Private Sub UserForm_Initialize()
Set Dic = CreateObject("Scripting.Dictionary")
    lbxMHH.List = SheetData.Range("A2:D" & SheetData.Range("A" & Rows.Count).End(xlUp).Row).Value
    lbxMHH.ColumnCount = 4
    lbxMHH.ListStyle = fmListStyleOption
End Sub
