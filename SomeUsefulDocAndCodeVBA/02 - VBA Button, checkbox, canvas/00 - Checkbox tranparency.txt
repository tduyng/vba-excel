Option Explicit

Sub checkBox_Transapency()
Dim chk As CheckBox
'Call OptimzeCode_Begin
    For Each chk In ActiveSheet.CheckBoxes
        With chk.Fill
            If chk.Value = True Then
                .Visible = msoTrue
                '.ForeColor.RGB = RGB(255, 0, 0)
                .BackColor = RGB(0, 0, 0)
                
            Else
                .Visible = msoTrue
                '.ForeColor.RGB = RGB(128, 128, 128)
                .BackColor = RGB(128, 128, 128)
            End If
        End With
    Debug.Print chk.Name
    Next chk

'Call OptimizeCode_End
End Sub


Sub checkBox_Transapency1()
Dim chk As Object
Dim nameChk As String
Dim ws As Worksheet
Set ws = ActiveSheet
For Each chk In ws.CheckBoxes
        Debug.Print "---------------------------------------------"
        Debug.Print "Form control on sheet: " & ws.Name
        Debug.Print "Location on sheet: " & chk.TopLeftCell.Address
        Debug.Print "Name of the component: " & chk.Name
        Debug.Print "Object type: CheckBox"
Next chk
End Sub

Sub CheckboxLoop()
'PURPOSE: Loop through each Form Control Checkbox on the ActiveSheet
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim cb As Shape

'Loop through Checkboxes
  For Each cb In ActiveSheet.Shapes
    If cb.Type = msoFormControl Then
      If cb.FormControlType = xlCheckBox Then
        If cb.ControlFormat.Value = 1 Then
          'Do something if checked...
          cb.Fill.ForeColor.RGB = RGB(0, 0, 0)
        ElseIf cb.ControlFormat.Value = -4146 Then
          'Do something if not checked...
          cb.Fill.ForeColor.RGB = RGB(128, 128, 128)
        ElseIf cb.ControlFormat.Value = 2 Then
          'Do something if mixed...
          cb.Fill.ForeColor.RGB = RGB(128, 128, 128)
        End If
      End If
    End If
  Next cb

End Sub

