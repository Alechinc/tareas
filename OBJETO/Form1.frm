VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4200
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4200
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DETENTE"
      Height          =   615
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VAMOS"
      Height          =   615
      Left            =   3600
      TabIndex        =   0
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Shape1.Left = 480 Then
Timer1.Enabled = True
ElseIf Shape1.Top = 5880 Then
Timer1.Enabled = True
ElseIf Shape1.Left = 7680 Then
Timer2.Enabled = True
ElseIf Shape1.Top = 300 Then
Timer2.Enabled = True
End If
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
x = Shape1.Top
x = x + 50
Shape1.Top = x
If Shape1.Top > 5880 Then
Shape1.Top = 5880
x = Shape1.Left
x = x + 50
Shape1.Left = x
End If
If Shape1.Left > 7680 Then
Shape1.Left = 7680
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
x = Shape1.Top
x = x - 50
Shape1.Top = x
If Shape1.Top < 300 Then
Shape1.Top = 300
x = Shape1.Left
x = x - 50
Shape1.Left = x
End If
If Shape1.Left < 480 Then
Shape1.Left = 480
Timer2.Enabled = False
Timer1.Enabled = True
End If
End Sub
