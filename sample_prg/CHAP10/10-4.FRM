VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "計算日本隊總分"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計算臺 灣 隊 總 分"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "日 本 隊 總 分"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "臺  灣  隊  總  分"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function sum(a() As Integer, n As Integer) As Integer
    s = 0
    For i = 0 To n - 1
        s = s + a(i)
    Next i
    sum = s
End Function

Private Sub Command1_Click()
    Dim taiwan(4) As Integer
    taiwan(0) = 90: taiwan(1) = 78: taiwan(2) = 82: taiwan(3) = 85: taiwan(4) = 96
    Text1.Text = sum(taiwan(), 5)
End Sub

Private Sub Command2_Click()
    Dim japan(4) As Integer
    japan(0) = 77: japan(1) = 68: japan(2) = 76: japan(3) = 91: japan(4) = 86
    Text2.Text = sum(japan(), 4)
End Sub
