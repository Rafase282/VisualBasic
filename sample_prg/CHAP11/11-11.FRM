VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   2910
   ClientTop       =   2085
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '≥Ã§Wºh±±®Ó∂µ™∫Ω’¶‚ΩL
   ScaleHeight     =   1995
   ScaleWidth      =   5460
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox Text 
      Height          =   270
      IMEMode         =   10  'SBCS HANGUL
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   270
      IMEMode         =   2  '√ˆ≥¨
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   270
      IMEMode         =   10  'SBCS HANGUL
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "© ßO"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "¶~    ƒ÷"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "©m    ¶W"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer

Private Sub Form_Load()
    n = 0
    Text(0).Text = ""
    Text(1).Text = ""
    Text(2).Text = ""
End Sub

Private Sub Text_GotFocus(Index As Integer)
    If Index = 1 Then
        Text(n).IMEMode = 2
    Else
        Text(n).IMEMode = 10
    End If
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        n = (n + 1) Mod 3
        Text(n).SetFocus
        If Index = 2 Then
            List1.AddItem Text(0).Text & "ÅF" & Text(1).Text & "ÅF" & Text(2).Text
            Text(0).Text = ""
            Text(1).Text = ""
            Text(2).Text = ""
            n = 0
        End If
    End If
End Sub


