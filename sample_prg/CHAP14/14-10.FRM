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
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Drag 1
End Sub

Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
    Static objs As New Collection
    For Each o In objs
        If o Is Source Then
            Exit Sub
        End If
    Next
    objs.Add Source
    List1.AddItem Source.Name
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.Drag 1
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Drag 1
End Sub
