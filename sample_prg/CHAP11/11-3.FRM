VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   3090
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   1200
      Picture         =   "11-3.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image Img 
      DragIcon        =   "11-3.frx":0442
      Height          =   480
      Index           =   2
      Left            =   2040
      Picture         =   "11-3.frx":0884
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Img 
      DragIcon        =   "11-3.frx":0CC6
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "11-3.frx":1108
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Img 
      DragIcon        =   "11-3.frx":154A
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "11-3.frx":198C
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dx As Single, dy As Single

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move X - dx, Y - dy
End Sub

Private Sub Form_Load()
    For i = 0 To 2
        Img(i).DragIcon = Img(i).Picture
    Next i
End Sub

Private Sub Img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Img(Index).Drag 1
    dx = X: dy = Y
End Sub

Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Visible = False
    Picture1.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Picture1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = 0 Then Picture1.BackColor = RGB(0, 0, 255)
    If State = 1 Then Picture1.BackColor = RGB(255, 255, 255)
End Sub
