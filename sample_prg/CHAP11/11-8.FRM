VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   2910
   ClientTop       =   2085
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   1935
   ScaleWidth      =   4740
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   2280
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "性別"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "年齡"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "姓名"
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
Private Sub Form_Load()
    Text(0).Text = ""
    Text(1).Text = ""
    Text(2).Text = ""
End Sub

Private Sub Text_LostFocus(Index As Integer)
    If Index = 2 Then
        List1.AddItem Text(0).Text & ":" & Text(1).Text & ":" & Text(2).Text
        Text(0).Text = ""
        Text(1).Text = ""
        Text(2).Text = ""
    End If
End Sub
