VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   2400
   End
   Begin VB.Image Img 
      Height          =   480
      Index           =   1
      Left            =   1560
      Picture         =   "17-9.frx":0000
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img 
      Height          =   480
      Index           =   0
      Left            =   840
      Picture         =   "17-9.frx":0442
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    If Form1.WindowState = 1 Then
        Form1.Icon = Img(0).Picture
        Timer1.Enabled = True
        Timer1.Interval = 1000
    Else
        Timer1.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    Static flag As Boolean
    If flag Then
        Form1.Icon = Img(0).Picture
    Else
        Form1.Icon = Img(1).Picture
    End If
    flag = Not flag
End Sub
