VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   5490
   StartUpPosition =   3  '系統預設值
   Begin VB.DirListBox Dir2 
      Height          =   1980
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.DriveListBox Drive2 
      Height          =   300
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      DragIcon        =   "9-6.frx":0000
      Height          =   2430
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1980
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "DragDrop 後"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "DragDrop 前"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
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

Private Sub Dir2_DragDrop(Source As Control, X As Single, Y As Single)
    Dim buf(1) As Byte
    f1 = Source.Path
    f2 = Dir2.Path
    If f1 <> f2 Then
        If Right(f1, 1) = "\" Then f1 = Left(f1, 2)
        Open f1 & "\" & Source.FileName For Binary As #1
        If Right(f2, 1) = "\" Then f2 = Left(f2, 2)
        Open f2 & "\" & Source.FileName For Binary As #2
        Get #1, , buf
        Do While Not EOF(1)
            Put #2, , buf
            Get #1, , buf
        Loop
        Close #1, #2
    End If
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub


Private Sub Drive2_Change()
    Dir2.Path = Drive2.Drive
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    File1.Drag 1
End Sub


