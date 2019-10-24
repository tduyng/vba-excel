I now use this code on every form that has [OK] and [Cancel] buttons, as it allows you to control the [OK], [Cancel] and [X] being clicked by the user, in the main code:

Option Explicit
Private cancelled As Boolean

Public Property Get IsCancelled() As Boolean
    IsCancelled = cancelled
End Property

Private Sub OkButton_Click()
    Hide
End Sub

Private Sub CancelButton_Click()
    OnCancel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    cancelled = True
    Hide
End Sub
You can then use the form as an instance as follows:

With New frm1    
    .Show
    If Not .IsCancelled Then
        ' do your stuff ...
    End If
End With
or you can use it as an object (variable?) as noted above such as:

Dim MyDialog As frm1

Set MyDialog = New frm1    'This fires Userform_Initialize




Private Sub cbOk_Click()
    'Sheet2.Range("F1") = Me.txtName.Text
    'Sheet2.Range("G1") = Me.txtAge.Text
    
    Sheet2.Range("F1") = getName
    Sheet2.Range("G1") = getAge
    
    'Unload Me
    'Me.unload
    Me.Hide
End Sub
Private Sub cbCancel_Click()
    Me.Hide
End Sub
Private iscancelled As Boolean
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
