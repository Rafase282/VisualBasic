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
      Height          =   1095
      Left            =   2280
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mv(obj As Object)
    obj.Width = obj.Width * 1.1
    obj.Height = obj.Height * 1.1
End Sub

Private Sub Command1_Click()
    mv Command1
End Sub

Private Sub Form_Click()
    mv Form1
End Sub

Private Sub Picture1_Click()
    mv Picture1
End Sub
