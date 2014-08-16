VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   3660
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一張"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "前一張"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "張"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "第"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2520
      Width           =   255
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '單線固定
      Height          =   1935
      Left            =   600
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer

Private Sub Command1_Click()
    n = n - 1
    If n < 1 Then n = 15
    Image1.Picture = LoadPicture("c:\program Files\microsoft visual studio\common\Graphics\icons\comm\phone" & Format(n, "00") & ".ico")
    Text1.Text = n
End Sub

Private Sub Command2_Click()
    n = n + 1
    If n > 15 Then n = 1
    Image1.Picture = LoadPicture("c:\program Files\microsoft visual studio\common\Graphics\icons\comm\phone" & Format(n, "00") & ".ico")
    Text1.Text = n
End Sub

Private Sub Form_Load()
    n = 0
End Sub
