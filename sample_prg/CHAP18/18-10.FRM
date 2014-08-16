VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   2520
   ClientTop       =   1425
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   5310
   ScaleWidth      =   4605
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtext 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"18-10.frx":0000
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu m1 
      Caption         =   "檔        案"
      Begin VB.Menu fopen 
         Caption         =   "開啟舊檔"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Private Sub fopen_Click()
    On Error GoTo ErrHandler
    CommonDialog1.ShowOpen
    rtext(0).LoadFile CommonDialog1.FileName
    If n <> 1 Then
        TabStrip1.Tabs.Add , , "File" & n
    End If
    Load rtext(n)
    rtext(n).Visible = True
    rtext(n).LoadFile CommonDialog1.FileName
    rtext(n).Left = TabStrip1.ClientLeft
    rtext(n).Top = TabStrip1.ClientTop
    rtext(n).Height = TabStrip1.ClientHeight
    rtext(n).Width = TabStrip1.ClientWidth
    rtext(n).ZOrder 0
    TabStrip1.Tabs(n).Selected = True
    n = n + 1
ErrHandler:
End Sub

Private Sub Form_Load()
    TabStrip1.Tabs(1).Caption = "File1"
    CommonDialog1.CancelError = True
    n = 1
End Sub

Private Sub TabStrip1_Click()
    rtext(TabStrip1.SelectedItem.Index).ZOrder 0
End Sub


