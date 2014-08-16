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
      Caption         =   "指  向  右  邊"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "指  向  左  邊"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1800
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mv(X As Integer, Y As Integer, fname As String)
    Image1.Left = X
    Image1.Top = Y
    Image1.Picture = LoadPicture(fname)
End Sub

Private Sub Command1_Click()
    mv Command1.Left, 400, "c:\Program Files\Microsoft Visual Studio\common\Graphics\icons\arrows\point02.ico"
End Sub

Private Sub Command2_Click()
    mv Command2.Left, 400, "c:\Program Files\Microsoft Visual Studio\common\Graphics\icons\arrows\point04.ico"
End Sub

