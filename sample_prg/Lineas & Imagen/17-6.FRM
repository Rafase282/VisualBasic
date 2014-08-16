VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   1590
   ClientTop       =   2085
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   6480
   ScaleWidth      =   8895
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   6255
      Left            =   2160
      ScaleHeight     =   6195
      ScaleWidth      =   6555
      TabIndex        =   3
      Top             =   120
      Width           =   6615
   End
   Begin VB.FileListBox File1 
      Height          =   2970
      Left            =   120
      Pattern         =   "*.ico;*.bmp"
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   2820
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    x = 20: y = 10
    Picture1.Cls
    f = File1.Path
    If Right(f, 1) = "\" Then f = Left(f, 2)
    For i = 0 To File1.ListCount - 1
        Image1.Picture = LoadPicture(f & "\" & File1.List(i))
        w = Image1.Width / Screen.TwipsPerPixelX
        h = Image1.Height / Screen.TwipsPerPixelY
        Picture1.PaintPicture Image1.Picture, x, y, 32, 32, 0, 0, w, h
        x = x + 40
        If x = 20 + 10 * 40 Then
            y = y + 40
            x = 20
        End If
    Next i
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Picture1.ScaleMode = 3
End Sub


