VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3975
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "顯    示"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2580
      Left            =   1320
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
Dim student(5) As String
Dim N As Integer

Option Compare Text

Private Sub Command1_Click()
    List1.Clear
    For k = 0 To N - 1
        If student(k) Like "[A-L]*" Then
            List1.AddItem student(k)
        End If
    Next k
End Sub

Private Sub Form_Load()
    N = 6
    student(0) = "Toyota": student(1) = "nissan": student(2) = "Honda"
    student(3) = "BENZ": student(4) = "bmw": student(5) = "Ford"
End Sub

