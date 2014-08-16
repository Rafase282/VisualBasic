VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   2685
   ClientTop       =   1785
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   6345
   ScaleWidth      =   8280
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   6255
      Left            =   0
      ScaleHeight     =   6195
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
   Begin VB.Menu m1 
      Caption         =   "檔    案"
      Begin VB.Menu m11 
         Caption         =   "開啟舊檔"
      End
      Begin VB.Menu m12 
         Caption         =   "儲存檔案"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DrawFlag As Boolean

Private Sub Form_Load()
    Picture1.AutoRedraw = True
End Sub


Private Sub m11_Click()
     CommonDialog1.ShowOpen
    Picture1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub m12_Click()
    CommonDialog1.ShowSave
    SavePicture Picture1.Image, CommonDialog1.FileName
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawFlag = True
    Picture1.PSet (X, Y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawFlag Then
        Picture1.Line -(X, Y)
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawFlag = False
End Sub


