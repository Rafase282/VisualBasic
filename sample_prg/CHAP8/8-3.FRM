VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5940
   StartUpPosition =   3  '╰参箇砞
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   2280
      ScaleHeight     =   2835
      ScaleWidth      =   3195
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   840
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  ﹚  Θ  罿  单  材"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Θ罿"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "﹎"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.Print "﹎", "Θ罿", "单材"
    t = Text2.Text
    If t >= 80 Then
        Picture1.ForeColor = RGB(0, 0, 255)
        Picture1.Print Text1.Text, t, "ヒ单"
    ElseIf t >= 70 Then
        Picture1.Print Text1.Text, t, "单"
    ElseIf t >= 60 Then
        Picture1.Print Text1.Text, t, "单"
    Else
        Picture1.ForeColor = RGB(255, 0, 0)
        Picture1.Print Text1.Text, t, "ぃの"
    End If
    Text1.Text = ""
    Text2.Text = ""
    Picture1.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
End Sub

