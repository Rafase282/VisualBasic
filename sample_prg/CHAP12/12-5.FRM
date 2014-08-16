VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   2415
   ScaleWidth      =   5535
   Begin VB.Image Img 
      BorderStyle     =   1  '單線固定
      Height          =   1215
      Index           =   2
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image Img 
      BorderStyle     =   1  '單線固定
      Height          =   1215
      Index           =   1
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image Img 
      BorderStyle     =   1  '單線固定
      Height          =   1215
      Index           =   0
      Left            =   480
      Picture         =   "12-5.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1215
   End
   Begin VB.Menu edit 
      Caption         =   "編    輯"
      Visible         =   0   'False
      Begin VB.Menu copy 
         Caption         =   "複      製"
      End
      Begin VB.Menu cut 
         Caption         =   "剪      下"
      End
      Begin VB.Menu paste 
         Caption         =   "貼      上"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer

Private Sub copy_Click()
    Clipboard.SetData Img(n).Picture
End Sub

Private Sub cut_Click()
    Clipboard.SetData Img(n).Picture
    Img(n).Picture = LoadPicture("")
End Sub


Private Sub Form_Load()
    Clipboard.Clear
End Sub

Private Sub Img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        n = Index
        PopupMenu Edit
    End If
End Sub


Private Sub paste_Click()
    Img(n).Picture = Clipboard.GetData()
End Sub


