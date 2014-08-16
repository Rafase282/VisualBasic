VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1890
   ClientLeft      =   3000
   ClientTop       =   3255
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   1890
   ScaleWidth      =   3630
   Begin VB.CommandButton Command1 
      Caption         =   "結    束"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Label1.Caption = "10 秒後結束"
    DoEvents
    st = Timer
    Do While Timer - st < 10
    Loop
    End
End Sub


