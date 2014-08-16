VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   7725
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   1680
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   4080
      ScaleHeight     =   1035
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   1080
      ScaleHeight     =   1035
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "14-8.frx":0000
      Top             =   720
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer

Private Sub Form_Load()
    flag = 1
    Image1.Top = Picture1.Top
End Sub

Private Sub Timer1_Timer()
    If flag = 1 And Image1.Left > Picture1.Left Then
        Set Image1.Container = Picture1
        Image1.Left = 0
        Image1.Top = 0
        flag = 2
    End If
    
    If flag = 2 And Image1.Left > Picture1.Width Then
        Set Image1.Container = Form1
        Image1.Left = Picture1.Left + Picture1.Width
        Image1.Top = Picture1.Top
        flag = 3
    End If
    
    If flag = 3 And Image1.Left > Picture2.Left Then
        Set Image1.Container = Picture2
        Image1.Left = 0
        Image1.Top = 0
        flag = 4
    End If
    
    If flag = 4 And Image1.Left > Picture2.Width Then
        Set Image1.Container = Form1
        Image1.Left = Picture2.Left + Picture2.Width
        Image1.Top = Picture2.Top
        flag = 5
    End If

    Image1.Left = Image1.Left + 100
End Sub
