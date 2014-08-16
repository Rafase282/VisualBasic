VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   6915
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   5640
      TabIndex        =   12
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "顯   示   資   料"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登   錄   資   料"
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   4875
      TabIndex        =   8
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "筆"
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "共"
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "數學"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "英文"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "國文"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "姓  名:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim table(10, 3)
Dim n As Integer

Private Sub Command1_Click()
    table(n, 0) = Text1.Text
    table(n, 1) = Text2.Text
    table(n, 2) = Text3.Text
    table(n, 3) = Text4.Text
    n = n + 1
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = n
End Sub

Private Sub Command2_Click()
    Picture1.Cls
    Picture1.Print "姓  名"; Tab(15); "國  文"; Tab(30); "英  文"; Tab(45); "數  學"
    For i = 0 To n - 1
        For j = 0 To 3
            Picture1.Print Tab(j * 15); table(i, j);
        Next j
        Picture1.Print
    Next i
End Sub

Private Sub Form_Load()
    n = 0
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Sub
