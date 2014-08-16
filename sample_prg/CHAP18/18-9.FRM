VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   3345
   ClientTop       =   2670
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   5070
   ScaleWidth      =   4050
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2640
      Width           =   3615
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4048
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "18-9.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "18-9.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "18-9.frx":0634
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msg(6) As String

Private Sub Form_Load()
    TreeView1.ImageList = ImageList1
    Dim nd As Node
    Set nd = TreeView1.Nodes.Add(, , , "遠東技術學院", 2)
        Set nd = TreeView1.Nodes.Add(1, tvwChild, , "企管科", 3)
        Set nd = TreeView1.Nodes.Add(1, tvwChild, , "工管科", 3)
        Set nd = TreeView1.Nodes.Add(1, tvwChild, , "資管科", 3)
    nd.EnsureVisible
    Set nd = TreeView1.Nodes.Add(, , , "台西技術學院", 2)
        Set nd = TreeView1.Nodes.Add(5, tvwChild, , "電機科", 3)
    nd.EnsureVisible
    
    Dim Lf As String
    Lf = Chr$(13) & Chr$(10)
    Msg(2) = "遠東技術學院的企業管理科" & Lf & "成立於民國81年" & Lf & "現有學生900人"
    Msg(3) = ""
    Msg(4) = ""
    Msg(6) = ""
End Sub

Private Sub TreeView1_Collapse(ByVal Node As Node)
    TreeView1.Nodes(Node.Index).Image = 1
End Sub

Private Sub TreeView1_Expand(ByVal Node As Node)
    TreeView1.Nodes(Node.Index).Image = 2
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As Node)
    Text1.Text = Msg(Node.Index)
End Sub


