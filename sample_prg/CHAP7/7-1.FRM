VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   3660
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2475
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "每次取2字顯示"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    msg = "Microsoft Visual Basic 6.0"
    For i = 2 To Len(msg) Step 2
        Picture1.Print Left(msg, i)
    Next i
End Sub
