VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4185
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "縮      小"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "放      大"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   1800
      Picture         =   "3-3.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Image1.Width = Image1.Width + 50
    Image1.Height = Image1.Height + 50
End Sub

Private Sub Command2_Click()
    Image1.Width = Image1.Width - 50
    Image1.Height = Image1.Height - 50
End Sub
