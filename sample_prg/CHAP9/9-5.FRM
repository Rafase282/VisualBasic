VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   3015
   ScaleWidth      =   3180
   Begin VB.CommandButton Command3 
      Caption         =   "結          束"
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "查    詢"
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "存      檔"
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   720
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   720
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   720
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   720
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1200
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "汽  車  編  號  "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "售  價"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "CC 數"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "車  名"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "廠  牌"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub textclr()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
End Sub

Private Sub Command1_Click()
    car.mark = Text2.Text
    car.car_name = Text3.Text
    car.cc = Val(Text4.Text)
    car.price = Val(Text5.Text)
    Put #1, Val(Text1.Text), car
    textclr
End Sub

Private Sub Command2_Click()
    If Val(Text1.Text) = 0 Then
        ret = MsgBox("請輸入汽車編號!", vbExclamation, "錯    誤")
    Else
        Get #1, Val(Text1.Text), car
        Text2.Text = car.mark
        Text3.Text = car.car_name
        Text4.Text = car.cc
        Text5.Text = car.price
    End If
End Sub

Private Sub Command3_Click()
    Close #1
    End
End Sub


Private Sub Form_Load()
    Open "car.rnd" For Random As #1 Len = Len(car)
    textclr
End Sub


