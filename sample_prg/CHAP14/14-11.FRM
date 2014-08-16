VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5175
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "顯    示"
      Height          =   735
      Left            =   2160
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   1530
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   1770
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
      BorderStyle     =   1  '單線固定
      Height          =   2175
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    f = File1.Path
    If Right(f, 1) = "\" Then f = Left(f, 2)
    Image1.Picture = LoadPicture(f & "\" & File1.FileName)
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    ftype = LCase(Right(File1.FileName, 3))
    If ftype = "ico" Or ftype = "bmp" Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
End Sub

