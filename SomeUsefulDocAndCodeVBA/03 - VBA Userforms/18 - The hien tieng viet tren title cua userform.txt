Function HWndOfUserForm(UF As MSForms.UserForm) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HWndOfUserForm
' This returns the window handle (HWnd) of the userform referenced
' by UF. It first looks for a top-level window, then a child
' of the Application window, then a child of the ActiveWindow.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim AppHWnd As Long
    Dim DeskHWnd As Long
    Dim WinHWnd As Long
    Dim UFHWnd As Long
    Dim Cap As String
    Dim WindowCap As String

    Cap = UF.Caption

    ' First, look in top level windows
    UFHWnd = FindWindow(C_USERFORM_CLASSNAME, Cap)
    If UFHWnd <> 0 Then
        HWndOfUserForm = UFHWnd
        Exit Function
    End If
    ' Not a top level window. Search for child of application.
    AppHWnd = Application.HWnd
    UFHWnd = FindWindowEx(AppHWnd, 0&, C_USERFORM_CLASSNAME, Cap)
    If UFHWnd <> 0 Then
        HWndOfUserForm = UFHWnd
        Exit Function
    End If
    ' Not a child of the application.
    ' Search for child of ActiveWindow (Excel's ActiveWindow, not
    ' Window's ActiveWindow).
    If Application.ActiveWindow Is Nothing Then
        HWndOfUserForm = 0
        Exit Function
    End If
    WinHWnd = WindowHWnd(Application.ActiveWindow)
    UFHWnd = FindWindowEx(WinHWnd, 0&, C_USERFORM_CLASSNAME, Cap)
    HWndOfUserForm = UFHWnd

End Function



Private Declare Function DefWindowProcW Lib "user32" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Public Sub SetUniText(UF As MSForms.UserForm, ByVal sUniText As String)
'
' Mo ta:        Unicode TitleBar, Frame, Button, CheckBox, Option
' Yeu cau:      Frame, Button, CheckBox, Option khong ho tro XP style
' Nguoi viet:  thuongall
' Email:        thuongall@yahoo.com
' Website:      www.caulacbovb.com
'
    Dim UFHWnd As Long
    Dim WinInfo As Long
    Dim r As Long

    UFHWnd = HWndOfUserForm(UF)
    If UFHWnd = 0 Then
        Exit Sub
    End If

    DefWindowProcW UFHWnd, WM_SETTEXT, &H0&, StrPtr(sUniText)
End Sub


Private Sub UserForm_Initialize()
    SetUniText Me, VNI("Coäng hoøa xaõ hoäi chuû nghóa vieät nam")
End Sub