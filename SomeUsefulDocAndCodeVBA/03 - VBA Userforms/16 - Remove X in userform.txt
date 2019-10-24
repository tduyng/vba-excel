Private Const mcGWL_STYLE = (-16)
Private Const mcWS_SYSMENU = &H80000

'Windows API calls to handle windows
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If


Public Sub subRemoveCloseButton(frm As Object)
    Dim lngStyle As Long
    Dim lngHWnd As Long

    lngHWnd = FindWindow(vbNullString, frm.Caption)
    lngStyle = GetWindowLong(lngHWnd, mcGWL_STYLE)

    If lngStyle And mcWS_SYSMENU > 0 Then
        SetWindowLong lngHWnd, mcGWL_STYLE, (lngStyle And Not mcWS_SYSMENU)
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, _
  CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then
    Cancel = True
    MsgBox "Please use the Close Form button!"
  End If
End Sub

subRemoveCloseButton Me 