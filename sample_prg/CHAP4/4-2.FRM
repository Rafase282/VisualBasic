VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   5250
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   2160
      ScaleHeight     =   1515
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "及    格    否"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "英文成績"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "學生姓名"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.Print Text1.Text; Tab(10); Text2.Text; Tab(17);
    If Text2.Text >= 60 Then
        Picture1.ForeColor = RGB(0, 0, 255)
        Picture1.Print "及  格"
    Else
        Picture1.ForeColor = RGB(255, 0, 0)
        Picture1.Print "不及格"
    End If
    Text1.Text = ""
    Text2.Text = ""
    Picture1.ForeColor = RGB(0, 0, 0)
End Sub
