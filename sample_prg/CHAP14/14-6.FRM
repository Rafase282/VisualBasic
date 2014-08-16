VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.Label Label1 
      Caption         =   "請按影像,它會移動 !"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image main 
      Height          =   495
      Left            =   120
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "14-6.frx":0000
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "14-6.frx":0442
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mv(a As Image)
    main.Top = a.Top
    main.Left = a.Left
    main.Picture = a.Picture
    Do While main.Left <= Form1.ScaleWidth - main.Width
        main.Move main.Left + 50
    Loop
End Sub
Private Sub Image1_Click()
    mv Image1
End Sub

Private Sub Image2_Click()
    mv Image2
End Sub
