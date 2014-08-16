VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   4620
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "清  除  圖  片"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "顯  示  圖  片"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3240
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   120
      Picture         =   "3-1.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.Visible = True
End Sub

Private Sub Command2_Click()
    Picture1.Visible = False
End Sub
