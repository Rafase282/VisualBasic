VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   3870
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "清        除"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "畫        圓"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   360
      ScaleHeight     =   2595
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    r = InputBox("請輸入圓的半徑", "輸入半徑")
    Picture1.BackColor = RGB(255, 255, 255)
    Picture1.Circle (Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2), r
End Sub


Private Sub Command2_Click()
    ret = MsgBox("您確定要清除畫面嗎?", vbOKCancel + vbExclamation, "清除畫面")
    If ret = 1 Then Picture1.Cls
End Sub


