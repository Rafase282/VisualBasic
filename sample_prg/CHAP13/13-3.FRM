VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '³Ì¤W¼h±±¨î¶µªº½Õ¦â½L
   ScaleHeight     =   4680
   ScaleWidth      =   6780
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"13-3.frx":0000
   End
   Begin VB.Menu m1 
      Caption         =   "ÀÉ  ®×"
      Begin VB.Menu m11 
         Caption         =   "¶}±ÒÂÂÀÉ"
      End
      Begin VB.Menu m12 
         Caption         =   "Àx¦sÀÉ®×•Û‘¶"
      End
      Begin VB.Menu m13 
         Caption         =   "µ²        §ô"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SaveFlag As Boolean

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If SaveFlag Then
        If MsgBox("¬O§_¯uªº­nµ²§ô", 36, "µ²§ô") = 7 Then
            Cancel = True
        Else
            m12_Click
        End If
    End If
End Sub


Private Sub m11_Click()
    CommonDialog1.ShowOpen
    RichTextBox1.LoadFile CommonDialog1.FileName
    Form1.Caption = CommonDialog1.FileTitle
    SaveFlag = False
End Sub


Private Sub m12_Click()
    CommonDialog1.ShowSave
    RichTextBox1.SaveFile CommonDialog1.FileName
    Form1.Caption = CommonDialog1.FileTitle
    SaveFlag = False
End Sub


Private Sub m13_Click()
    Unload Form1
End Sub

Private Sub RichTextBox1_Change()
    SaveFlag = True
End Sub

