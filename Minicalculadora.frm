VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   Caption         =   "MiniCalculadora By R.282"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   735
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Text3.Text = Val(Text1.Text) + Val(Text2.Text)
Label1.Caption = "+"
End Sub

Private Sub Command2_Click()
Text3.Text = Val(Text1.Text) - Val(Text2.Text)
Label1.Caption = "-"
End Sub

Private Sub Command3_Click()
Text3.Text = Val(Text1.Text) * Val(Text2.Text)
Label1.Caption = "X"
End Sub

Private Sub Command4_Click()
Text3.Text = Val(Text1.Text) / Val(Text2.Text)
Label1.Caption = "/"
End Sub

Private Sub exit_Click()
End
End Sub
