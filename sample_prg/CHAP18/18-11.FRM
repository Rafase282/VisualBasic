VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   3210
   ClientTop       =   2280
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   2805
   ScaleWidth      =   6930
   Begin ComctlLib.Slider Slider 
      Height          =   495
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   327682
   End
   Begin ComctlLib.Slider Slider 
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   327682
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   4440
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin ComctlLib.Slider Slider 
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   327682
   End
   Begin VB.Label Label4 
      Caption         =   "藍"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "綠"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "紅"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   " 0      1      2     3      4     5      6     7      8     9     10"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    For i = 0 To 2
        Slider(i).Min = 0
        Slider(i).Max = 255
        Slider(i).SmallChange = 1
        Slider(i).LargeChange = 10
        Slider(i).TickFrequency = Slider(i).Max / 10
    Next i
    Picture1.BackColor = RGB(0, 0, 0)
End Sub

Private Sub Slider_Scroll(Index As Integer)
    r = Slider(0).Value
    g = Slider(1).Value
    b = Slider(2).Value
    Picture1.BackColor = RGB(r, g, b)
End Sub
