VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   5955
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確          定"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox Check 
      Caption         =   "離   別   的   火   車   站"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox Check 
      Caption         =   "媽   媽   請   妳   也   保   重"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CheckBox Check 
      Caption         =   "悲   情   的   城   市"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "請選擇您喜歡的歌曲"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Text1.Text = ""
    For i = 0 To 2
        If Check(i).Value Then
            Text1.Text = Text1.Text + Check(i).Caption + Chr(13) + Chr(10)
        End If
    Next i
End Sub

Private Sub Form_Load()
    Text1.Text = ""
End Sub
