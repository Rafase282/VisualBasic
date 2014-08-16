VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   3645
   ClientTop       =   2295
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   4200
   ScaleWidth      =   6225
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   240
      _ExtentX        =   476
      _ExtentY        =   661
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   495
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"18-14.frx":0000
   End
   Begin VB.Label Label1 
      Caption         =   "字型大小"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Text1.Text = 11
    RichTextBox1.SelFontSize = 11
End Sub

Private Sub UpDown1_DownClick()
    If Text1.Text > 8 Then
        Text1.Text = Text1.Text - 1
        RichTextBox1.SelFontSize = Text1.Text
    End If
End Sub

Private Sub UpDown1_UpClick()
    If Text1.Text < 200 Then
        Text1.Text = Text1.Text + 1
        RichTextBox1.SelFontSize = Text1.Text
    End If
End Sub
