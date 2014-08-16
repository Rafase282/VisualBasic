VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6630
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox List1 
      Height          =   4380
      Left            =   2400
      TabIndex        =   3
      Top             =   480
      Width           =   3975
   End
   Begin VB.FileListBox File1 
      Height          =   2610
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
   Begin VB.Label Label2 
      Caption         =   "顯示檔案內容"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "請選擇檔案"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
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

Private Sub File1_DblClick()
    f = File1.Path
    If Right(f, 1) = "\" Then f = Left(f, 2)
    Open f & "\" & File1.FileName For Input As #1
    List1.Clear
    Do While Not EOF(1)
        Line Input #1, buf
        List1.AddItem buf
    Loop
    Close #1
End Sub

