VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   4275
   StartUpPosition =   3  '系統預設值
   Begin VB.VScrollBar VScroll1 
      Height          =   2655
      Left            =   3720
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  '沒有框線
      Height          =   2535
      Left            =   -120
      ScaleHeight     =   2535
      ScaleWidth      =   3615
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  '沒有框線
         Height          =   7200
         Left            =   0
         Picture         =   "17-10.frx":0000
         ScaleHeight     =   7200
         ScaleWidth      =   9600
         TabIndex        =   1
         Top             =   -600
         Width           =   9600
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ScaleMode = vbPixels
    Picture1.ScaleMode = vbPixels
    Picture1.Move 0, 0, ScaleWidth - VScroll1.Width, ScaleHeight - HScroll1.Height
    Picture2.Move 0, 0
    HScroll1.Top = Picture1.Height
    HScroll1.Left = 0
    HScroll1.Width = Picture1.Width
    VScroll1.Top = 0
    VScroll1.Left = Picture1.Width
    VScroll1.Height = Picture1.Height
    HScroll1.Max = Picture2.Width - Picture1.Width
    VScroll1.Max = Picture2.Height - Picture1.Height
    VScroll1.Visible = (Picture1.Height < Picture2.Height)
    HScroll1.Visible = (Picture1.Width < Picture2.Width)
End Sub

Private Sub HScroll1_Change()
    Picture2.Left = -HScroll1.Value
End Sub

Private Sub VScroll1_Change()
    Picture2.Top = -VScroll1.Value
End Sub
