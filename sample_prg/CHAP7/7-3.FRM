VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    msg = "123I愛妳"
    Print Len(msg), LenB(msg)
    Print msg
    Print RightB(msg, 8)
    Print Right(msg, 3)
    For i = 1 To LenB(msg)
        Print Hex(AscB(MidB(msg, i, 1))); " ";
    Next i
    Print
    For i = 1 To Len(msg)
        Print Hex(AscW(Mid(msg, i, 1))); " ";
    Next i
End Sub

