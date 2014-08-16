VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4185
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox List1 
      Height          =   2220
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "輸     入"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    student = InputBox("請輸入學生姓名", "輸入姓名")
    Do While student <> ""
        List1.AddItem student
        student = InputBox("請輸入學生姓名", "輸入姓名")
    Loop
End Sub

