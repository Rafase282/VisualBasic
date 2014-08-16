VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1931
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"12-2.frx":0000
   End
   Begin VB.Menu m1 
      Caption         =   "檔      案(&F)"
      Begin VB.Menu m11 
         Caption         =   "開啟新檔(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu m12 
         Caption         =   "儲存檔案(&S)"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    RichTextBox1.Width = ScaleWidth
    RichTextBox1.Height = ScaleHeight
End Sub

Private Sub m11_Click()
    CommonDialog1.ShowOpen
    RichTextBox1.LoadFile CommonDialog1.FileName
End Sub

Private Sub m12_Click()
    CommonDialog1.ShowSave
    RichTextBox1.SaveFile CommonDialog1.FileName
End Sub
