VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4905
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame2 
      Caption         =   "大小"
      Height          =   1695
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1335
      Begin VB.OptionButton Option6 
         Caption         =   "24"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "18"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "12"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "顏色"
      Height          =   1695
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1455
      Begin VB.OptionButton Option3 
         Caption         =   "藍    色"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "綠    色"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "紅    色"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "諸法皆空,自由自在"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()
    Label1.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Option2_Click()
    Label1.ForeColor = RGB(0, 255, 0)
End Sub

Private Sub Option3_Click()
    Label1.ForeColor = RGB(0, 0, 255)
End Sub

Private Sub Option4_Click()
    Label1.FontSize = 12
End Sub

Private Sub Option5_Click()
    Label1.FontSize = 18
End Sub

Private Sub Option6_Click()
    Label1.FontSize = 24
End Sub
