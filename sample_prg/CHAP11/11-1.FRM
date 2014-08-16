VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   6975
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DrawFlag As Boolean

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawFlag = True
    PSet (X, Y), RGB(0, 0, 255)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawFlag Then
        Line -(X, Y), RGB(0, 0, 255)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawFlag = False
End Sub
