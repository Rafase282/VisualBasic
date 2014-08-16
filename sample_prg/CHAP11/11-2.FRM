VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4425
   StartUpPosition =   3  '系統預設值
   Begin VB.Image Image1 
      DragIcon        =   "11-2.frx":0000
      Height          =   480
      Left            =   480
      Picture         =   "11-2.frx":0442
      Top             =   360
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

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move Image1.Left + X - dx, Image1.Top + Y - dy
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Drag 1
    dx = X: dy = Y
End Sub
