VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   3150
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Pic 
      Height          =   495
      Index           =   3
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic 
      Height          =   495
      Index           =   2
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.PictureBox Pic 
      Height          =   495
      Index           =   1
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.PictureBox Pic 
      Height          =   495
      Index           =   0
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "您可以任選一方格改變底色"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Pic_Click(Index As Integer)
    CommonDialog1.ShowColor
    Pic(Index).BackColor = CommonDialog1.Color
End Sub
