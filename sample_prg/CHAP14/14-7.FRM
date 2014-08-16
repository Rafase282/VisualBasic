VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "產     生     副     本"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "副本編號"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim fi As New Form1
    fi.Show
    fi.Move Left + 800, Top + 800
    fi.Text1.Text = n
    n = n + 1
End Sub

