VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2925
   ClientLeft      =   2295
   ClientTop       =   3270
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   2925
   ScaleWidth      =   6840
   Begin VB.CommandButton Command2 
      Caption         =   "列          印"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2220
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登          錄"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "職  位"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "姓  名"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Singer(100) As String, Song(100) As String
Dim N As Integer


Private Sub Command1_Click()
    List1.AddItem Text1.Text & ":" & Text2.Text
    Singer(N) = Text1.Text
    Song(N) = Text2.Text
    N = N + 1
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Command2_Click()
    Printer.ScaleMode = 4
    CharX = Printer.TextWidth("壹")
    CharY = Printer.TextHeight("壹")
    M = 32
    Printer.CurrentY = 1
    Printer.Print Tab(10); "姓"; Tab(20); "名"; Tab(42); "職"; Tab(52); "位"
    For i = 0 To N
        Printer.Print Tab(3); Singer(i); Tab(34); Song(i)
    Next i
    
    Printer.DrawWidth = 4
    Printer.Line (1, 1)-(M * CharX, (N + 1) * CharY + 1), , B
    Printer.Line (1, CharY + 1)-(M * CharX, CharY + 1)
    Printer.Line (16 * CharX, 1)-(16 * CharX, (N + 1) * CharY + 1)
    Printer.DrawWidth = 2
    For i = 2 To N
        Printer.Line (1, i * CharY + 1)-(M * CharX, i * CharY + 1)
    Next i

    Printer.EndDoc

End Sub


Private Sub Form_Load()
    N = 0
End Sub


