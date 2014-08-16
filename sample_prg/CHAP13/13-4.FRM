VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   3615
   StartUpPosition =   3  '系統預設值
   Begin VB.Label Label1 
      Caption         =   "請 Double Click  Image1 以載入圖片"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '單線固定
      Height          =   2775
      Left            =   480
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_DblClick()
    Form2.Show 1
End Sub
