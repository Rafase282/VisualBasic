VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   3270
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "清除 Form"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   360
      ScaleHeight     =   2475
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "畫    圓"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.Circle (1000, 1000), 1000, RGB(255, 0, 0)
    Form1.Circle (1000, 1000), 1000
End Sub

Private Sub Command2_Click()
    Form1.Cls
End Sub
