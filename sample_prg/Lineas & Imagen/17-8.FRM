VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   960
      ScaleHeight     =   2955
      ScaleWidth      =   3555
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   495
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   240
      _ExtentX        =   476
      _ExtentY        =   661
      _Version        =   327681
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DrawFlag As Boolean

Private Sub Form_Load()
    Text1.Text = 1
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawFlag = True
    Picture1.PSet (X, Y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawFlag Then
        Picture1.Line -(X, Y)
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawFlag = False
End Sub

Private Sub UpDown1_DownClick()
    If Text1.Text > 1 Then
        Text1.Text = Val(Text1.Text) - 1
        Picture1.DrawWidth = Text1.Text
    End If
End Sub

Private Sub UpDown1_UpClick()
    Text1.Text = Val(Text1.Text) + 1
    Picture1.DrawWidth = Text1.Text
End Sub
