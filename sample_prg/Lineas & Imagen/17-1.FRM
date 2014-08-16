VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   4620
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3435
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "畫    圖"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3960
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.Scale (-32, 20)-(32, -20)
    Picture1.Line (-32, 0)-(32, 0)
    Picture1.Line (0, -20)-(0, 20)
    For x = -30 To 30
        y = x * x * x / 500
        If x = -30 Then
            Picture1.PSet (x, y)
        Else
            Picture1.Line -(x, y)
        End If
    Next x
End Sub
