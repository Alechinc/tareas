VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "CLENING"
      Height          =   735
      Left            =   1080
      TabIndex        =   8
      Top             =   0
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   0
      List            =   "Form1.frx":000D
      TabIndex        =   7
      Text            =   "STYLE FONT"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0030
      Left            =   3960
      List            =   "Form1.frx":0043
      TabIndex        =   6
      Text            =   "SIZE"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "Form1.frx":005B
      Left            =   1680
      List            =   "Form1.frx":0068
      TabIndex        =   5
      Text            =   "Color de Fondo de Objeto"
      Top             =   3480
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form1.frx":0084
      Left            =   5640
      List            =   "Form1.frx":0091
      TabIndex        =   4
      Text            =   "COLOR FONT"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SUBRAYADO"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CURSIVA"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REMARCADO"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "10" Then
Text1.FontSize = "10"
End If
If Combo1.Text = "20" Then
Text1.FontSize = "20"
End If
If Combo1.Text = "30" Then
Text1.FontSize = "30"
End If
If Combo1.Text = "40" Then
Text1.FontSize = "40"
End If
If Combo1.Text = "50" Then
Text1.FontSize = "50"
End If
End Sub

Private Sub Combo2_Click()
If Combo2.Text = "Arial Black" Then
Text1.Font = "Arial "
End If
If Combo2.Text = "Arial" Then
Text1.Font = "ALGERIA"
End If
If Combo2.Text = "Agency FB" Then
Text1.Font = "Agency FB"
End If
End Sub

Private Sub Combo3_Click()
If Combo3.Text = "Rojo" Then
Text1.ForeColor = vbPINK
End If
If Combo3.Text = "Amarillo" Then
Text1.ForeColor = vbYellow
End If
If Combo3.Text = "Azul" Then
Text1.ForeColor = vbBlue
End If
End Sub

Private Sub Combo4_Click()
If Combo4.Text = "Verde" Then
Shape1.BackColor = vbGreen
End If
If Combo4.Text = "Naranja" Then
Shape1.BackColor = &H80FF&
End If
If Combo4.Text = "Rosado" Then
Shape1.BackColor = &HFF80FF
End If
End Sub

Private Sub Command1_Click()
Text1.FontBold = True
End Sub

Private Sub Command2_Click()
Text1.FontItalic = True
End Sub

Private Sub Command3_Click()
Text1.FontUnderline = True
End Sub

Private Sub Command4_Click()
Text1.FontBold = False
Text1.FontItalic = False
Text1.FontUnderline = False
End Sub
