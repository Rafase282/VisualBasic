VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "www.rafase282.tk By R.282"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C00000&
      Caption         =   "Enter the Chat"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   450
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "www.rafase282.tk"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Hide
Form2.Show
End Sub
