Option Explicit

Private Sub cbAdd_Click()
    insertToSheet
    Unload Me
End Sub

Private Sub cbCancel_Click()
    Unload Me
End Sub
Private Sub insertToSheet()
    With Sheet2
        Dim currentRow As Long
        If .Range("A10") = "" Then
            currentRow = 1
        Else
            currentRow = .Range("A" & .Rows.Count).End(xlUp).Row + 1
        End If
        .Cells(currentRow, 1) = txtFirstName.Value
        .Cells(currentRow, 1).Offset(0, 1) = UCase(txtLastName.Value)
    End With
        
End Sub

