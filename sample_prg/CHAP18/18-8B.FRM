VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  '單線固定
   Caption         =   "Text的設定"
   ClientHeight    =   3465
   ClientLeft      =   4950
   ClientTop       =   5340
   ClientWidth     =   5085
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   3465
   ScaleWidth      =   5085
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O K"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "字    型"
      TabPicture(0)   =   "18-8b.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(1)=   "Combo1"
      Tab(0).Control(2)=   "Label1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "位    置"
      TabPicture(1)   =   "18-8b.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   840
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   840
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   -71400
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   -74640
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "高度"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "寬度"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Size"
         Height          =   255
         Left            =   -71880
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Text1.FontName = Combo1.Text
    Form1.Text1.FontSize = Val(Text1.Text)
    Form1.Text1.Width = Val(Text2.Text)
    Form1.Text1.Height = Val(Text3.Text)
    Unload Form2
End Sub

Private Sub Command2_Click()
    Unload Form2
End Sub

Private Sub Form_Load()
    For i = 0 To Screen.FontCount - 1
        Combo1.AddItem Screen.Fonts(i)
    Next i
    Combo1.Text = Form1.Text1.FontName
    Text1.Text = Form1.Text1.FontSize
    Text2.Text = Form1.Text1.Width
    Text3.Text = Form1.Text1.Height
End Sub

