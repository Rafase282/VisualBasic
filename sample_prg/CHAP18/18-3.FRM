VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   2730
   ClientTop       =   2175
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   3645
   ScaleWidth      =   5490
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DrawFlag As Boolean

Private Sub Form_Load()
    DrawFlag = False
    StatusBar1.Panels.Item(3).Text = Date
    StatusBar1.Panels.Item(4).Text = Time
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawFlag = True
    PSet (X, Y)
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawFlag Then
        Line -(X, Y)
    End If
    StatusBar1.Panels.Item(1).Text = "X=" & Format(X, "00000")
    StatusBar1.Panels.Item(2).Text = "Y=" & Format(Y, "00000")
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawFlag = False
End Sub



