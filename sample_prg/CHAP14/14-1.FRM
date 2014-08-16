VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   4155
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "顯  示  汽  車  基  本  資  料"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "售價  ( 萬元 )"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "引擎  CC  數"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "車             名"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type toyota
    car_name As String
    car_cc As Integer
    car_price As Single
End Type

Dim car(2) As toyota

Private Sub Command1_Click()
    Static n As Integer
    Text1.Text = car(n).car_name
    Text2.Text = car(n).car_cc
    Text3.Text = car(n).car_price
    n = n + 1
    If n >= 3 Then
       n = 0
       ret = MsgBox("沒有資料了", vbAbortRetryIgnore + vbExclamation, "沒有資料了")
       Text1.Text = ""
       Text2.Text = ""
       Text3.Text = ""
    End If
End Sub

Private Sub Form_Load()
    car(0).car_name = "Tercel": car(0).car_cc = 1300: car(0).car_price = 30.2
    car(1).car_name = "Exsior": car(1).car_cc = 1600: car(1).car_price = 60.5
    car(2).car_name = "Camary": car(2).car_cc = 2000: car(2).car_price = 80.5
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub
