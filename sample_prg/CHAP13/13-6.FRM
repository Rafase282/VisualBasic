VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   4965
   StartUpPosition =   3  '系統預設值
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2655
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4683
      _Version        =   393217
      TextRTF         =   $"13-6.frx":0000
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   3240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "尋找字串"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu m1 
      Caption         =   "檔    案"
      Begin VB.Menu m11 
         Caption         =   "開啟舊檔"
      End
      Begin VB.Menu m12 
         Caption         =   "尋        找"
      End
      Begin VB.Menu m13 
         Caption         =   "再次尋找"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer

Private Sub m11_Click()
    CommonDialog1.ShowOpen
    RichTextBox1.LoadFile CommonDialog1.FileName
End Sub

Private Sub m12_Click()
    n = InStr(RichTextBox1.Text, Text1.Text)
    If n > 0 Then
        RichTextBox1.SetFocus
        RichTextBox1.SelStart = n - 1
        RichTextBox1.SelLength = Len(Text1.Text)
    End If
End Sub

Private Sub m13_Click()
    n = InStr(n + Len(Text1.Text), RichTextBox1.Text, Text1.Text)
    If n > 0 Then
        RichTextBox1.SetFocus
        RichTextBox1.SelStart = n - 1
        RichTextBox1.SelLength = Len(Text1.Text)
    End If
End Sub
