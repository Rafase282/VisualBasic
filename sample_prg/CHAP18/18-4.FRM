VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   10095
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "取消重填"
      Height          =   735
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定登錄"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   720
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   720
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   720
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   2415
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "住址"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "電話"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "姓名"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Set lv = ListView1.ListItems.Add(, , Text1.Text)
    lv.SubItems(1) = Text2.Text
    lv.SubItems(2) = Text3.Text
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub

Private Sub Form_Load()
    Dim x As ColumnHeader
    w = ListView1.Width / 3.5
    Set x = ListView1.ColumnHeaders.Add(, , "姓名", w)
    Set x = ListView1.ColumnHeaders.Add(, , "電話", w)
    Set x = ListView1.ColumnHeaders.Add(, , "住址", w)
    ListView1.View = lvwReport
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub

