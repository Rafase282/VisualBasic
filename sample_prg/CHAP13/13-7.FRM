VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4200
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5565
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '系統預設值
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu m1 
      Caption         =   "檔        案"
      WindowList      =   -1  'True
      Begin VB.Menu m11 
         Caption         =   "開啟新檔"
      End
      Begin VB.Menu m12 
         Caption         =   "開啟舊檔"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub m11_Click()
    Dim newform As New Form1
    newform.Show
End Sub

Private Sub m12_Click()
    CommonDialog1.ShowOpen
    ActiveForm.RichTextBox1.LoadFile CommonDialog1.FileName
    ActiveForm.Caption = CommonDialog1.FileTitle
End Sub

