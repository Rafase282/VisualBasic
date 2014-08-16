VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6225
   StartUpPosition =   3  '系統預設值
   Begin VB.VScrollBar VScroll1 
      Height          =   4335
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Image Image1 
         Height          =   480
         Left            =   840
         Picture         =   "6-2.frx":0000
         Top             =   480
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    HScroll1.Max = Picture1.ScaleWidth - Image1.Width
    HScroll1.LargeChange = HScroll1.Max / 30
    HScroll1.SmallChange = HScroll1.Max / 60
    VScroll1.Max = Picture1.ScaleHeight - Image1.Height
    VScroll1.LargeChange = VScroll1.Max / 30
    VScroll1.SmallChange = VScroll1.Max / 60
    Image1.Move 0, 0
End Sub

Private Sub HScroll1_Change()
    Image1.Left = HScroll1.Value
End Sub

Private Sub VScroll1_Change()
    Image1.Top = VScroll1.Value
End Sub
