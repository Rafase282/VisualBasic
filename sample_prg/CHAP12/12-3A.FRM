VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   5505
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7435
      _Version        =   393217
      TextRTF         =   $"12-3a.frx":0000
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
