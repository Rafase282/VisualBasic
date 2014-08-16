VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   5730
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "畫正多角形"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    For n = 3 To 9
        SetAngle 0
        SetPoint 2000, 4500
        For j = 1 To n
            Movel 1200, RGB(0, 0, 255)
            Turn 360 / n
        Next j
    Next n
End Sub
