VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   3495
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "清    除"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   360
      ScaleHeight     =   2955
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "顯    示"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim x(5) As Integer
    Dim y(5) As Integer
    x(0) = 1000: y(0) = 1000: x(1) = 1200: y(1) = 2400: x(2) = 1400: y(2) = 1200
    x(3) = 1600: y(3) = 2200: x(4) = 1800: y(4) = 1600: x(5) = 2000: y(5) = 2000
    Picture1.BackColor = RGB(255, 255, 255)
    For i = 0 To 5
        If i = 0 Then
            Picture1.PSet (x(i), y(i))
        Else
            Picture1.Line -(x(i), y(i))
        End If
    Next i
End Sub

Private Sub Command2_Click()
    Picture1.Cls
End Sub
