VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查詢就讀學校"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "就讀學校"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "姓名 :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stu_name(5) As String, school(5) As String

Private Sub Command1_Click()
    For i = 0 To 5
        If Text1.Text = stu_name(i) Then
            Text2.Text = school(i)
            Exit Sub
        End If
    Next i
    Text2.Text = "找不到您輸入的姓名"
End Sub

Private Sub Form_Load()
    stu_name(0) = "NIKE": stu_name(1) = "ADIDAS": stu_name(2) = "MISUNO"
    stu_name(3) = "ALLSTAR": stu_name(4) = "HENTENG": stu_name(5) = "KIMKO"
    school(0) = "臺灣大學": school(1) = "清華大學": school(2) = "臺灣科技大學"
    school(3) = "交通大學": school(4) = "成功大學": school(5) = "遠東科技大學"
    Text1.Text = ""
    Text2.Text = ""
End Sub
