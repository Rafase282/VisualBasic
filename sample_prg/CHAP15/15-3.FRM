VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   10785
   StartUpPosition =   3  '系統預設值
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   4800
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.FileListBox File1 
      Height          =   2250
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1350
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "選擇字型 :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "選擇檔案 :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    List1.FontName = Combo1.Text
End Sub

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

Private Sub Form_Load()
    For i = 0 To Screen.FontCount - 1
        Combo1.AddItem Screen.Fonts(i)
    Next i
    Combo1.Text = Screen.Fonts(0)
End Sub
