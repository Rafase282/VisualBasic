VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4050
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "休息 10 秒"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   120
      MousePointer    =   2  '十字形狀
      ScaleHeight     =   1515
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2400
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Screen.MousePointer = 11
    st = Timer
    Do While Timer - st < 10
    Loop
    Screen.MousePointer = 0
End Sub
