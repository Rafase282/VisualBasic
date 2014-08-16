VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6015
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   960
      ScaleHeight     =   1995
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "顯  示  圖  片"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "圖  檔  名"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.Picture = LoadPicture(Text1.Text)
End Sub

