VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   3525
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "清    除"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   360
      ScaleHeight     =   2595
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "畫直線"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3000
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.BackColor = RGB(255, 255, 255)
    Picture1.Line (0, 0)-(2000, 2000)
End Sub

Private Sub Command2_Click()
    Picture1.Cls
End Sub
