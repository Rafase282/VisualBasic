VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldx As Single, oldy As Single
Dim flag As Boolean, mf As Integer
Dim sx As Single, sy As Single

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    flag = True
    mf = DrawMode
    DrawMode = 7
    Line (0, Y)-(Width, Y), RGB(255, 255, 255)
    Line (X, 0)-(X, Height), RGB(255, 255, 255)
    oldx = X: oldy = Y
    sx = X: sy = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If flag Then
        Line (0, Y)-(Width, Y), RGB(255, 255, 255)
        Line (0, oldy)-(Width, oldy), RGB(255, 255, 255)
        Line (X, 0)-(X, Height), RGB(255, 255, 255)
        Line (oldx, 0)-(oldx, Height), RGB(255, 255, 255)
        oldx = X: oldy = Y
    End If
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If flag Then
        flag = False
        Line (0, Y)-(Width, Y), RGB(255, 255, 255)
        Line (X, 0)-(X, Height), RGB(255, 255, 255)
        DrawMode = mf
        Line (sx, sy)-(X, Y)
    End If
End Sub


