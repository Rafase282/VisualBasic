VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   4170
   StartUpPosition =   3  '╰参箇砞
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ゑ  耕    "
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "      计"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "ヒ      计"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function max(a As Integer, b As Integer) As String
    If a > b Then
        max = "ヒ    "
    Else
        max = "    ヒ"
    End If
End Function

Private Sub Command1_Click()
    Picture1.Cls
    Picture1.Print "ヒ计", "计", "ゑ耕挡狦"
    Dim x As Integer, y As Integer
    x = Text1.Text
    y = Text2.Text
    Picture1.Print x, y, max(x, y)
    Text1.Text = ""
    Text2.Text = ""
End Sub

