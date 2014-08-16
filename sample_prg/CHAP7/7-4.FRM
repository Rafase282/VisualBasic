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
      Caption         =   "排    序"
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
Dim Girl(5) As String
Dim N As Integer

Option Compare Text

Private Sub Command1_Click()
    For k = 0 To N - 2
        Min = Girl(k)
        s = k
        For j = k + 1 To N - 1
            If Girl(j) < Min Then
                Min = Girl(j)
                s = j
            End If
        Next j
        t = Girl(k): Girl(k) = Girl(s): Girl(s) = t
    Next k
    List1.Clear
    For k = 0 To N - 1
        List1.AddItem Girl(k)
    Next k
End Sub

Private Sub Form_Load()
    N = 6
    Girl(0) = "Amy": Girl(1) = "rolla": Girl(2) = "lucky"
    Girl(3) = "Nike": Girl(4) = "Lisa": Girl(5) = "Nancy"
End Sub

