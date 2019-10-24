'***************************************************************************
'*
'* MODULE NAME:     USERFORM WINDOW STYLES
'* AUTHOR:          STEPHEN BULLEN, Office Automation Ltd.
'*                  TIM CLEM
'*
'* CONTACT:         Stephen@oaltd.co.uk
'* WEB SITE:        http://www.oaltd.co.uk
'*
'* DESCRIPTION:     Changes userform's window styles to give different visual effects
'*
'* THIS MODULE:     Changes the userform's styles so it can be resized/maximised/minimized, etc.
'*                  The code was initially created by Tim Clem, and expanded by Stephen Bullen.
'*
'* UPDATES:
'*  DATE            COMMENTS
'*  11 Jan 2005     Changed the way 'ShowInTaskBar' works, fixing a bug found by Jamie Collins
'*
'***************************************************************************

Option Explicit

'Windows API calls to do all the dirty work!
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

'Lots of window styles for us to play with!
Private Const GWL_STYLE As Long = (-16)          'The offset of a window's style
Private Const GWL_EXSTYLE As Long = (-20)        'The offset of a window's extended style
Private Const WS_CAPTION As Long = &HC00000      'Style to add a titlebar
Private Const WS_SYSMENU As Long = &H80000       'Style to add a system menu
Private Const WS_THICKFRAME As Long = &H40000    'Style to add a sizable frame
Private Const WS_MINIMIZEBOX As Long = &H20000   'Style to add a Minimize box on the title bar
Private Const WS_MAXIMIZEBOX As Long = &H10000   'Style to add a Maximize box to the title bar
Private Const WS_EX_APPWINDOW As Long = &H40000  'Application Window: shown on taskbar
Private Const WS_EX_TOOLWINDOW As Long = &H80    'Tool Window: small titlebar

'Constant to identify the Close menu item
Private Const SC_CLOSE As Long = &HF060

'Constants for hide or show a window
Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5

'Constants for Windows messages
Private Const WM_SETICON = &H80

'Variables to store the various selections/options
Dim mbSizeable As Boolean, mbCaption As Boolean, mbIcon As Boolean
Dim mbMaximize As Boolean, mbMinimize As Boolean, mbSysMenu As Boolean, mbCloseBtn As Boolean
Dim mbAppWindow As Boolean, mbToolWindow As Boolean, mbModal As Boolean
Dim msIconPath As String
Dim moForm As Object
Dim mhWndForm As Long

'Set the class's initial properties to be those of a default userform
Private Sub Class_Initialize()
    mbCaption = True
    mbSysMenu = True
    mbCloseBtn = True
End Sub

'Allow the calling code to tell us which form to handle
Public Property Set Form(oForm As Object)

    'Get the userform's window handle
    If Val(Application.Version) < 9 Then
        mhWndForm = FindWindow("ThunderXFrame", oForm.Caption)    'XL97
    Else
        mhWndForm = FindWindow("ThunderDFrame", oForm.Caption)    'XL2000+
    End If

    'Remember the form for later
    Set moForm = oForm

    'Set the form's style
    SetFormStyle

    'Update the form's icon
    ChangeIcon

    'Update the taskbar visibility
    If mbAppWindow Then ShowTaskBarIcon = True

End Property

'***************************************************************
'* Property procedures to get and set the form's window styles
'***************************************************************

Public Property Let Modal(bModal As Boolean)
    mbModal = bModal

    'Make the form modal or modeless by enabling/disabling Excel itself
    EnableWindow FindWindow("XLMAIN", Application.Caption), Abs(CInt(Not mbModal))
End Property

Public Property Get Modal() As Boolean
    Modal = mbModal
End Property

Public Property Let Sizeable(bSizeable As Boolean)
    mbSizeable = bSizeable
    SetFormStyle
End Property

Public Property Get Sizeable() As Boolean
    Sizeable = mbSizeable
End Property

Public Property Let ShowCaption(bCaption As Boolean)
    mbCaption = bCaption
    SetFormStyle
End Property

Public Property Get ShowCaption() As Boolean
    ShowCaption = mbCaption
End Property

Public Property Let SmallCaption(bToolWindow As Boolean)
    mbToolWindow = bToolWindow
    SetFormStyle
End Property

Public Property Get SmallCaption() As Boolean
    SmallCaption = mbToolWindow
End Property

Public Property Let ShowMaximizeBtn(bMaximize As Boolean)
    mbMaximize = bMaximize
    SetFormStyle
End Property

Public Property Get ShowMaximizeBtn() As Boolean
    ShowMaximizeBtn = mbMaximize
End Property

Public Property Let ShowMinimizeBtn(bMinimize As Boolean)
    mbMinimize = bMinimize
    SetFormStyle
End Property

Public Property Get ShowMinimizeBtn() As Boolean
    ShowMinimizeBtn = mbMinimize
End Property

Public Property Let ShowSysMenu(bSysMenu As Boolean)
    mbSysMenu = bSysMenu
    SetFormStyle
End Property

Public Property Get ShowSysMenu() As Boolean
    ShowSysMenu = mbSysMenu
End Property

Public Property Let ShowCloseBtn(bCloseBtn As Boolean)
    mbCloseBtn = bCloseBtn
    SetFormStyle
End Property

Public Property Get ShowCloseBtn() As Boolean
    ShowCloseBtn = mbCloseBtn
End Property

Public Property Let ShowTaskBarIcon(bAppWindow As Boolean)

    mbAppWindow = bAppWindow

    'When showing/hiding the task bar icon, we have to hide and reshow the form
    'to get Windows to update the task bar
    If mhWndForm <> 0 Then
        'Freeze the form, to avoid flicker when hiding/showing it
        LockWindowUpdate mhWndForm

        'Enable the Excel window, so we don't lose focus
        EnableWindow FindWindow("XLMAIN", Application.Caption), True

        'Hide the form
        ShowWindow mhWndForm, SW_HIDE

        'Update the style bits
        SetFormStyle

        'Reshow the userform
        ShowWindow mhWndForm, SW_SHOW

        'Unfreeze the form
        LockWindowUpdate 0&

        'Set the Excel window's enablement to the correct choice
        EnableWindow FindWindow("XLMAIN", Application.Caption), Abs(CInt(Not mbModal))
    End If

End Property

Public Property Get ShowTaskBarIcon() As Boolean
    ShowTaskBarIcon = mbAppWindow
End Property

Public Property Let ShowIcon(bIcon As Boolean)
    mbIcon = Not bIcon
    ChangeIcon
    SetFormStyle
End Property

Public Property Get ShowIcon() As Boolean
    ShowIcon = (mbIcon <> 1)
End Property

Public Property Let IconPath(sNewPath As String)
    msIconPath = sNewPath
    ChangeIcon
    SetFormStyle
End Property

Public Property Get IconPath() As String
    IconPath = msIconPath
End Property