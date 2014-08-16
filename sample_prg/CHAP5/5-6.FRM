VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   5655
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton key 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "請用滑鼠 Click 鍵盤"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Text1.Text = ""
    key(0).Caption = "a"
    For i = 1 To 25
        Load key(i)
        If i < 13 Then
            dx = key(0).Width * i
            dy = 0
        Else
            dx = key(0).Width * (i - 13)
            dy = key(0).Height
        End If
        key(i).Move key(0).Left + dx, key(0).Top + dy
        key(i).Visible = True
        key(i).Caption = Chr(Asc("a") + i)
    Next i
End Sub

Private Sub key_Click(Index As Integer)
    Text1.Text = Text1.Text + key(Index).Caption
End Sub
