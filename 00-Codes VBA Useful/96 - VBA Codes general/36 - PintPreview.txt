Private Sub cbPreview_Click()
    Dim shtMenu As Worksheet

    On Error GoTo sheetInvalid
    Set shtModel = ThisWorkbook.Sheets("02-MODEL")
    With shtModel
        If Not IsSheetVisible(.Range("A1")) Then
        .Visible = True
        End If
        Me.Hide
        Application.Goto Reference:="Model_Facture"
        ActiveWindow.SelectedSheets.PrintPreview
        Me.Show
'        .PrintPreview
'        .Visible = xlHidden
    End With
sheetInvalid:
    If Err.Number = 9 Then
        Set shtModel = shModel
        Resume Next
    End If
    
End Sub