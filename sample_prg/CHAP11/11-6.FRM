VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1920
      Picture         =   "11-6.frx":0000
      Top             =   1200
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And 1) = 1 Then
        mx = 200: my = 200
    Else
        mx = 30: my = 30
    End If
    Select Case KeyCode
        Case &H26
            Image1.Move Image1.Left, Image1.Top - my
        Case &H28
            Image1.Move Image1.Left, Image1.Top + my
        Case &H25
            Image1.Move Image1.Left - mx
        Case &H27
            Image1.Move Image1.Left + mx
    End Select
End Sub

