VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   3960
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   600
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "畫    橫    線"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    For x = 300 To 2400 Step 300
        Picture1.Line (300, x)-(2400, x)
    Next x
End Sub
