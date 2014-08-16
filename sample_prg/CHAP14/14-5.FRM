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
   Begin VB.CommandButton Command2 
      Caption         =   "顯    示  Form3"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "顯   示  Form2"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Form

Private Sub Command1_Click()
    If Not (f Is Nothing) Then
        Unload f
        Set f = Nothing
    End If
    Set f = Form2
    f.Show
    f.Print "我是 Form2 !!!"
End Sub

Private Sub Command2_Click()
    If Not (f Is Nothing) Then
        Unload f
        Set f = Nothing
    End If
    Set f = Form3
    f.Show
    f.Print "我是 Form3 !!!"
End Sub

