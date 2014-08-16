VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.Menu m1 
      Caption         =   "應  用  程  式"
      Begin VB.Menu m11 
         Caption         =   "小  畫  家"
         Index           =   0
      End
      Begin VB.Menu m11 
         Caption         =   "筆  記  本"
         Index           =   1
      End
      Begin VB.Menu m11 
         Caption         =   "CD 播  放  程  式"
         Index           =   2
      End
      Begin VB.Menu m12 
         Caption         =   "-"
      End
      Begin VB.Menu m13 
         Caption         =   "結    束"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub m11_Click(Index As Integer)
    Select Case Index
        Case 0
            RET = Shell("pbrush.exe", 1)
        Case 1
            RET = Shell("notepad.exe", 1)
        Case 2
            RET = Shell("cdplayer.exe", 1)
    End Select
End Sub

Private Sub m13_Click()
    End
End Sub
