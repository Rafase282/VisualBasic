VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form2"
   ScaleHeight     =   2610
   ScaleWidth      =   4665
   StartUpPosition =   3  '系統預設值
   Begin VB.FileListBox File1 
      Height          =   2250
      Left            =   2160
      Pattern         =   "*.bmp;*.ico"
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   1770
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
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

Private Sub File1_DblClick()
    f = File1.Path
    If Right(f, 1) = "\" Then f = Left(f, 2)
    Form1.Image1.Picture = LoadPicture(f & "\" & File1.FileName)
    Unload Form2
End Sub

