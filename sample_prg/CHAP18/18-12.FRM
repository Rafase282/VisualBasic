VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   3660
   ClientTop       =   2385
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   4200
   ScaleWidth      =   5520
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "複    製"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2610
      Left            =   2640
      MultiSelect     =   2  '進階多重選取
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
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
    n = 0
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) Then
            n = n + 1
        End If
    Next i
    ProgressBar1.Min = 0
    ProgressBar1.Max = n
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) Then
            Text1.Text = File1.List(i) & " --->a:"
            DoEvents
            FileCopy f & "\" & File1.List(i), "a:\" & File1.List(i)
            ProgressBar1.Value = ProgressBar1.Value + 1
        End If
    Next i
    Text1.Text = ""
    ProgressBar1.Value = 0
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub


