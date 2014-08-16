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
   Begin VB.Image Img 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "14-3.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Img 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "14-3.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Img 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "14-3.frx":0884
      Top             =   120
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

Private Sub Img_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Source.Move Img(Index).Left + X - dx, Img(Index).Top + Y - dy
End Sub

Private Sub Img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Img(Index).Drag 1
    dx = X: dy = Y
End Sub
