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
      Top             =   2160
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   240
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Img(15) As Picture

Private Sub Form_Load()
    For i = 0 To 14
        Set Img(i) = LoadPicture("C:\Program Files\Microsoft Visual Studio\Common\Graphics\Icons\Comm\Phone" & Format(i + 1, "00") & ".ico")
    Next i
    Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()
    Static n As Integer
    Image1.Picture = Img(n)
    n = n + 1
    If n > 14 Then n = 0
End Sub
