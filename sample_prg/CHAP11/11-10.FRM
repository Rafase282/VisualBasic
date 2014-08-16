VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   2790
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text_GotFocus(Index As Integer)
    Text(Index).BackColor = RGB(255, 0, 0)
End Sub

Private Sub Text_LostFocus(Index As Integer)
    Text(Index).BackColor = RGB(255, 255, 255)
End Sub
