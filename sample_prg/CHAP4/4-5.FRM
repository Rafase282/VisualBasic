VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   6975
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   1440
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "4-5.frx":0000
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Image1.Left = Image1.Left + 100
    If Image1.Left > Form1.ScaleWidth Then
        Image1.Left = 0
    End If
End Sub
