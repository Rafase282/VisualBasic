VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   9135
   StartUpPosition =   3  '系統預設值
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8493
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"18-2.frx":0000
   End
   Begin VB.Menu m1 
      Caption         =   "檔    案"
      Begin VB.Menu fopen 
         Caption         =   "開啟舊檔"
      End
      Begin VB.Menu fsave 
         Caption         =   "儲存檔案"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fopen_Click()
    CommonDialog1.ShowOpen
    RichTextBox1.LoadFile CommonDialog1.FileName
    Form1.Caption = CommonDialog1.FileTitle
End Sub

Private Sub fsave_Click()
    CommonDialog1.ShowSave
    RichTextBox1.SaveFile CommonDialog1.FileName
    Form1.Caption = CommonDialog1.FileTitle
End Sub
