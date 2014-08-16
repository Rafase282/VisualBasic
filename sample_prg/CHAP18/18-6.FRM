VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   2370
   ClientTop       =   2385
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   3900
   ScaleWidth      =   3765
   Begin VB.CommandButton Command1 
      Caption         =   "分    割"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.PictureBox Pic 
      Height          =   1215
      Index           =   3
      Left            =   1920
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Pic 
      Height          =   1215
      Index           =   2
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Pic 
      Height          =   1215
      Index           =   1
      Left            =   1920
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Pic 
      Height          =   1215
      Index           =   0
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   240
      Top             =   2880
      _ExtentX        =   1614
      _ExtentY        =   1614
      _Version        =   393216
      Picture         =   "18-6.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    PictureClip1.Rows = 2
    PictureClip1.Cols = 2
    PictureClip1.StretchX = Pic(0).ScaleWidth
    PictureClip1.StretchY = Pic(0).ScaleHeight
    For i = 0 To 3
        Pic(i).Picture = PictureClip1.GraphicCell(i)
    Next i
End Sub

Private Sub Form_Load()
    For i = 0 To 3
        Pic(i).ScaleMode = 3
    Next i
End Sub




