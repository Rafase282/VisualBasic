VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    For r = 100 To 1000 Step 100
        Circle (3000, 1500), r
    Next r
End Sub
   
Private Sub Form_Load()
    For r = 100 To 1000 Step 100
        Circle (1000, 1500), r
    Next r
End Sub


Private Sub Form_Paint()
    For r = 100 To 1000 Step 100
        Circle (5000, 1500), r
    Next r
End Sub
