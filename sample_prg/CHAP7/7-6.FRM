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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   1200
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "今    天    是"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "今    天    是"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "現在時間是"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Text1.Text = Now
    Text2.Text = Format(Now, "yyyy年m月d日")
    Text3.Text = Format(Now, "ggge年m月d日aaaa")
End Sub
