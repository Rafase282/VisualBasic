VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4230
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox List1 
      Height          =   2220
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "顯    示"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim girl() As String
Dim n As Integer, max As Integer

Private Sub Command1_Click()
    List1.Clear
    For i = 0 To n - 1
        List1.AddItem girl(i)
    Next i
End Sub

Private Sub Form_Load()
    n = 0: max = 10
    ReDim girl(max)
    Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If n > max Then
            max = max + 20
            ReDim Preserve girl(max)
        End If
        girl(n) = Text1.Text
        n = n + 1
        Text1.Text = ""
    End If
End Sub
