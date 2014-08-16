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
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2640
      Picture         =   "14-4.frx":0000
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "14-4.frx":0442
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Zoom(a As Image)
    a.Stretch = True
    a.Width = a.Width * 1.5
    a.Height = a.Height * 1.5
End Sub

Private Sub Image1_Click()
    Zoom Image1
End Sub

Private Sub Image2_Click()
    Zoom Image2
End Sub
