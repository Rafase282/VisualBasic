VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   5640
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Caption = "LastEntry:" & GetSetting("RegTest", "Startup", "LastEntry", "?")
    Width = Val(GetSetting("RegTest", "Startup", "SizeW", "3000"))
    Height = Val(GetSetting("RegTest", "Startup", "SizeH", "3000"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "RegTest", "Startup", "Lastentry", Time
    SaveSetting "RegTest", "Startup", "SizeW", Width
    SaveSetting "RegTest", "Startup", "SizeH", Height
End Sub
