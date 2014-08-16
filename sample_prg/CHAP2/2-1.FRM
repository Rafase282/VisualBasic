VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   3810
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "清    除"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   480
      ScaleHeight     =   2595
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "畫    圓"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.Circle (1000, 1000), 1000, RGB(255, 0, 0)
End Sub

Private Sub Command2_Click()
    Picture1.Cls
End Sub

