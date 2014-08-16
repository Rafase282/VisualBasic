VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   4335
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   3240
      Picture         =   "15-2.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   4080
      Width           =   615
   End
   Begin VB.FileListBox File1 
      DragIcon        =   "15-2.frx":0442
      Height          =   2430
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1350
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "選擇檔案 :"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
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


Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    File1.Drag 1
End Sub


Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
    If Right(Source.Path, 1) = "\" Then
        Open Source.Path & Source.FileName For Input As #1
    Else
        Open Source.Path & "\" & Source.FileName For Input As #1
    End If
    Do While Not EOF(1)
        Line Input #1, buf
        Printer.Print buf
    Loop
    Printer.EndDoc
    Close #1
End Sub
