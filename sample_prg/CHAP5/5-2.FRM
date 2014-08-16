VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4920
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "登    錄    資    料"
      Height          =   1215
      Left            =   720
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2760
      Left            =   2400
      TabIndex        =   6
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "性 別"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "年 齡"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "姓 名"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    List1.AddItem Text1.Text & " == " & Text2.Text & " == " & Text3.Text
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub
