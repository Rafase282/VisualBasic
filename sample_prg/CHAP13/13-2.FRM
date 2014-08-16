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
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   360
      ScaleHeight     =   795
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    Picture1.Top = 0
    Picture1.Left = 0
    Picture1.Width = ScaleWidth
    Picture1.Height = ScaleHeight
    Picture1.Refresh
End Sub

Private Sub Picture1_Paint()
    If Picture1.ScaleHeight > Picture1.ScaleWidth Then
        r = Picture1.ScaleWidth / 2
    Else
        r = Picture1.ScaleHeight / 2
    End If
    Picture1.Circle (r, r), r
End Sub
