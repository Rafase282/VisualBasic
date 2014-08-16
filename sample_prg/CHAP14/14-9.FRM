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
      Caption         =   "顯示副本的屬性Caption"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "產      生    副      本"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   2400
      ScaleHeight     =   2835
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim fi As New Form1
    fi.Show
    fi.Move Left + 500, Top + 500
    fi.Caption = "instance" & n
    n = n + 1
End Sub

Private Sub Command2_Click()
    For i = 0 To Forms.Count - 1
        Picture1.Print Forms(i).Caption
    Next i
End Sub

