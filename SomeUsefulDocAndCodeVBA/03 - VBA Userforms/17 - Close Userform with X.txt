Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Me.Hide
        CloseForm.Show
    End If
End Sub

Private Sub UserForm_Activate()
    Me.Caption = "You must Click me to kill me!"
End Sub

Private Sub UserForm_Click()
  Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevent user from closing with the Close box in the title bar.
    If CloseMode <> 1 Then Cancel = 1
    Me.Caption = "The Close box won't work! Click me!"
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim Rep As Byte
 
If CloseMode = vbFormControlMenu Then
    Rep = MsgBox("Etes-vous sûr de vouloir quitter l'accès aux Encours en Production?", vbYesNo + vbQuestion, "Quitter l'accès aux données techniques?")
    If Rep = vbYes Then
        Workbooks("EncoursProduction.xlsm").Close False
    Else
        Cancel = True
    End If
End If


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
Dim Rep As Byte
 
FmAccueilUtilisateur.Hide
If CloseMode = vbFormControlMenu Then
    Rep = MsgBox("Etes-vous sûr de vouloir quitter l'accès aux Encours en Production?", vbYesNo + vbQuestion, "Quitter l'accès aux données techniques?")
    If Rep = vbYes Then
        Workbooks("EncoursProduction.xlsm").Close False 'Fermeture du fichier EncoursProduction + enregistrement
        End                                             'Fin du programme
    End If
End If
FmAccueilUtilisateur.Show
 
End Sub