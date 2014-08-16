VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "列          印"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   930
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "選擇檔案 :"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '單線固定
      Height          =   2655
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Label1.Caption = File1.FileName
    Form2.Image1.Picture = Image1.Picture
    Form2.PrintForm
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    f = File1.Path
    If Right(f, 1) = "\" Then f = Left(f, 2)
    Image1.Picture = LoadPicture(f & "\" & File1.FileName)
End Sub
