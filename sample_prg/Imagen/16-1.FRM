VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   6315
   StartUpPosition =   3  '系統預設值
   Begin VB.FileListBox File1 
      Height          =   1710
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1350
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '單線固定
      Height          =   3735
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    On Error GoTo ErrHandler
    f = File1.Path
    If Right(f, 1) = "\" Then f = Left(f, 2)
    Image1.Picture = LoadPicture(f & "\" & File1.FileName)
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, "錯誤"
End Sub

