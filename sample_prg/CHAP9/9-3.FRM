VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "顯 示 檔 案 內 容"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "關閉檔案"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "存    檔"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "電    話"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "姓    名"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If FreeFile = 1 Then
        Open "tel.dat" For Append As #1
    End If
    If Text1.Text = "" Then
        Close #1
    Else
        Write #1, Text1.Text, Text2.Text
    End If
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Command3_Click()
    On Error GoTo Errheader
    Open "tel.dat" For Input As #1
    Do While Not EOF(1)
        Input #1, namae, tel
        List1.AddItem namae & ":" & tel
    Loop
    Close #1
    Exit Sub
Errheader:
   ret = MsgBox("請先關閉檔案", vbExclamation, "關閉檔案")
End Sub

Private Sub Command2_Click()
    Close #1
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
End Sub

