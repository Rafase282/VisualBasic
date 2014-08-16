VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   4530
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetArray(a() As String, ParamArray dat())
    i = 0
    For Each x In dat
        a(i) = x
        i = i + 1
    Next
End Sub

Private Sub Form_Click()
    Dim car(5) As String
    SetArray car(), "Toyota", "Nissan", "Honda"
    For i = 0 To 5
        Print car(i)
    Next i
End Sub

