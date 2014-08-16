VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   5625
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "清            除"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   4800
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   720
      ScaleHeight     =   4395
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "畫   鑽   石   圖"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.BackColor = RGB(255, 255, 255)
    rd = 3.14159 / 180
    For i = 0 To 14
        X1 = 2000 * Cos(i * 22.5 * rd) + 2000
        Y1 = 2000 * Sin(i * 22.5 * rd) + 2000
        For j = i + 1 To 21
            X2 = 2000 * Cos(j * 22.5 * rd) + 2000
            Y2 = 2000 * Sin(j * 22.5 * rd) + 2000
            Picture1.Line (X1, Y1)-(X2, Y2)
        Next j
    Next i
End Sub

Private Sub Command2_Click()
    Picture1.Cls
End Sub
