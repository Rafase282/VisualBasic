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
   Begin VB.ListBox List1 
      Height          =   960
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Menu m1 
      Caption         =   "控    制    項"
      Begin VB.Menu menuct 
         Caption         =   "a"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctArray(20) As Control

Private Sub Form_Load()
    Dim ct As Control
    n = 0
    For Each ct In Controls
        If Not (TypeOf ct Is Menu) Then
            If n <> 0 Then
                Load menuct(n)
            End If
            menuct(n).Caption = ct.Name
            Set ctArray(n) = ct
            n = n + 1
        End If
    Next
End Sub

Private Sub menuct_Click(Index As Integer)
    ctArray(Index).SetFocus
End Sub
