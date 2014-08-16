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
Private Sub Disp(Namae As String, Optional Age As Variant = "-", Optional Tel As String = "-")
    Print Namae, Age, Tel
End Sub

Private Sub Form_Click()
    Disp "Kobe"
    Disp "Tokyo", 16
    Disp "Kyoto", 22, "07-2343234"
End Sub

