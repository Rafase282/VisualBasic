VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   5910
   StartUpPosition =   3  '系統預設值
   Begin VB.FileListBox File1 
      Height          =   1530
      Left            =   120
      Pattern         =   "*.ico;*.bmp"
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   2190
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '單線固定
      Height          =   1215
      Left            =   2400
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    f = File1.Path
    If Right(f, 1) = "\" Then f = Left(f, 2)
    For i = 0 To File1.ListCount - 1
        Image1.Picture = LoadPicture(f & "\" & File1.List(i))
        DoEvents
        File1.ListIndex = i
        st = Timer
        Do While Timer - st < 1
        Loop
    Next i
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

