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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Action = MsgBox("Select Case 程式測試", vbAbortRetryIgnore + vbQuestion, "訊息對話方塊")
    Select Case Action
        Case 3
            Print "您按了[異常終止]鈕!!"
        Case 4
            Print "您按了[重試]鈕!!"
        Case 5
            Print "您按了[略過]鈕!!"
    End Select
End Sub

