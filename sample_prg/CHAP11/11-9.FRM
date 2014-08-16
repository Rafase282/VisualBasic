VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   3510
   StartUpPosition =   3  '系統預設值
   Begin VB.Image Img 
      Height          =   720
      Index           =   5
      Left            =   1920
      Picture         =   "11-9.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   720
   End
   Begin VB.Image Img 
      Height          =   720
      Index           =   4
      Left            =   1440
      Picture         =   "11-9.frx":018A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   720
   End
   Begin VB.Image Img 
      Height          =   720
      Index           =   3
      Left            =   1200
      Picture         =   "11-9.frx":0314
      Stretch         =   -1  'True
      Top             =   960
      Width           =   720
   End
   Begin VB.Image Img 
      Height          =   720
      Index           =   2
      Left            =   1800
      Picture         =   "11-9.frx":0756
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image Img 
      Height          =   720
      Index           =   1
      Left            =   840
      Picture         =   "11-9.frx":0B98
      Stretch         =   -1  'True
      Top             =   360
      Width           =   720
   End
   Begin VB.Image Img 
      Height          =   720
      Index           =   0
      Left            =   1440
      Picture         =   "11-9.frx":0D22
      Stretch         =   -1  'True
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Img_Click(Index As Integer)
    Img(Index).ZOrder 0
End Sub

Private Sub Img_DblClick(Index As Integer)
    Img(Index).ZOrder 1
End Sub
