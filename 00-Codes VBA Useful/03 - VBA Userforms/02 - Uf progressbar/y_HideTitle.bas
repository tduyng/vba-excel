Attribute VB_Name = "y_HideTitle"
'Hide tile bar

Option Explicit
Option Private Module

Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
Public Const mcGWL_STYLE = (-16)
Public Const mcWS_SYSMENU = &H80000

#If VBA7 Then
    Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
                           
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
                           
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    
    Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
#Else
    Public Declare Function GetWindowLong _
                           Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
                           
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
    Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    
    Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
#End If
Sub HideTitleBar(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub


Public Sub subRemoveCloseButton(frm As Object)
    Dim lngStyle As Long
    Dim lngHWnd As Long

    lngHWnd = FindWindowA(vbNullString, frm.Caption)
    lngStyle = GetWindowLong(lngHWnd, mcGWL_STYLE)

    If lngStyle And mcWS_SYSMENU > 0 Then
        SetWindowLong lngHWnd, mcGWL_STYLE, (lngStyle And Not mcWS_SYSMENU)
    End If
End Sub
