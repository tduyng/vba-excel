
Private iscancelled As Boolean

Private Sub cbCancel_Click()
    Me.Hide
End Sub

Private Sub cbOk_Click()
    'Sheet2.Range("F1") = Me.txtName.Text
    'Sheet2.Range("G1") = Me.txtAge.Text
    
    Sheet2.Range("F1") = getName
    Sheet2.Range("G1") = getAge
    
    'Unload Me
    'Me.unload
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    'Su kien khi khoi tao userform, thi dieu gi se xay ra
    With ufBHE
        .BackColor = rgbBlue
        .BorderColor = rgbCyan
        .ForeColor = rgbRed
    End With
End Sub

Public Property Get getAge() As Variant
    getAge = txtAge.Value
End Property

Public Property Get getName() As Variant
    getName = txtName.Value
End Property
Public Property Get Huy() As Variant
    Huy = iscancelled
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
        iscancelled = True
    End If
End Sub






Option Explicit
Sub selectInputFile()
    Dim fd As Office.FileDialog
    Dim SelectedFile As String
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
 
      .AllowMultiSelect = True
      .InitialFileName = Application.ActiveWorkbook.path & "\"
      .Filters.Clear
      .Filters.Add "Excel", "*.xl*"
      .Filters.Add "All Files", "*.*"
 
      If .Show = True Then
        SelectedFile = .SelectedItems(1)
      End If
    End With
End Sub

Sub CALL_UF()
    Dim frm As New ufBHE
    frm.Show False
    If frm.Huy Then
        Debug.Print "Userform bi huy"
    Else
       Debug.Print frm.getAge
    End If
'    Unload ufBHE
End Sub
