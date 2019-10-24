VERSION 5.00
Begin VB.Form frmClass 
   Caption         =   "Draw a Box"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdTwoBoxes 
      Caption         =   "Two Boxes"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton CmdAnimate 
      Caption         =   "Animate"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton CmdDraw 
      Caption         =   "Draw"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents myBox As clsBox
Attribute myBox.VB_VarHelpID = -1
Private Sub myBox_Draw(X As Integer, Y As Integer)
     MsgBox ("A box has been drawed at " & CStr(X) & "," & CStr(Y))
End Sub

Private Sub CmdAnimate_Click()
   Dim position As Integer
   Dim myBox As clsBox
   Set myBox = New clsBox
   With myBox
     .Y = 50
     .Width = 1000
     .Height = 800
     For position = 10 To 4000
  '     .ClearBox Me
       .DrawBox Me, Me.BackColor
       .X = position
       .DrawBox Me, vbRed
     Next
   End With

End Sub

Private Sub CmdDraw_Click()
   Set myBox = New clsBox
   With myBox
     .X = 50
     .Y = 50
     .Width = 1000
     .Height = 800
     .DrawBox Me
   End With
End Sub

Private Sub CmdTwoBoxes_Click()
   Dim position As Integer
   Dim XOffset As Integer
   Dim myBox As clsBox
   Set myBox = New clsBox
   Dim smallBox As clsBox
   Set smallBox = New clsBox
   With myBox
     .Y = 50
     .Width = 1000
     .Height = 800
     smallBox.Height = 500
     smallBox.Y = .Y + (.Height - smallBox.Height) \ 2
     smallBox.Width = 700
     XOffset = (.Width - smallBox.Width) \ 2
     For position = 10 To 4000
       .DrawBox Me, Me.BackColor
       .X = position
       .DrawBox Me, vbRed
       smallBox.DrawBox Me, Me.BackColor
       smallBox.X = .X + XOffset
       smallBox.DrawBox Me, vbBlue
     Next
   End With

End Sub

