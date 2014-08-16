VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   5325
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "請      選       擇"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "學   校   全   名"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sch_title(3) As String
Dim sch_name(3) As String

Private Sub Combo1_Click()
    Text1.Text = sch_name(Combo1.ListIndex)
End Sub

Private Sub Form_Load()
    sch_title(0) = "NTU"
    sch_name(0) = "國立臺灣大學"
    sch_title(1) = "NTNU"
    sch_name(1) = "國立臺灣師範大學"
    sch_title(2) = "NCKU"
    sch_name(2) = "國立成功大學"
    sch_title(3) = "NSYSU"
    sch_name(3) = "國立中山大學"
    For i = 0 To 3
        Combo1.AddItem sch_title(i)
    Next i
    Combo1.Text = sch_title(0)
End Sub
