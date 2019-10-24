VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mX As Integer
Private mY As Integer
Private mWidth As Integer
Private mHeight As Integer
Public Event Draw(X As Integer, Y As Integer)

Public Property Let X(ByVal vValue As Integer)
  If (vValue > 0) Then
     mX = vValue
  End If
End Property
Public Property Get X() As Integer
  X = mX
End Property

Public Property Let Y(ByVal vValue As Integer)
  mY = vValue
End Property
Public Property Get Y() As Integer
  Y = mY
End Property
Public Property Let Width(ByVal vValue As Integer)
  mWidth = vValue
End Property
Public Property Get Width() As Integer
  Width = mWidth
End Property
Public Property Let Height(ByVal vValue As Integer)
  mHeight = vValue
End Property
Public Property Get Height() As Integer
  Height = mHeight
End Property

'Public Sub DrawBox(Canvas As Object)
'  Canvas.Line (mX, mY)-(mX + mWidth, mY + mHeight), , B
'End Sub

Public Sub DrawBox(Canvas As Object, Optional fColor As Long)
  If IsMissing(fColor) Then
    Canvas.Line (mX, mY)-(mX + mWidth, mY + mHeight), , B
  Else
    Canvas.Line (mX, mY)-(mX + mWidth, mY + mHeight), fColor, B
  End If
  RaiseEvent Draw(mX, mY)
End Sub

Public Sub ClearBox(Canvas As Object)
  Canvas.Line (mX, mY)-(mX + mWidth, mY + mHeight), Canvas.BackColor, B
End Sub
