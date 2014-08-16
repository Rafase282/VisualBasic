VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "www.rafase282.tk Chat By R.282"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   2955
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
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
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   3855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   1320
      ItemData        =   "Chat.frx":0000
      Left            =   4080
      List            =   "Chat.frx":0007
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
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
      Left            =   4200
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   1860
      ItemData        =   "Chat.frx":0017
      Left            =   120
      List            =   "Chat.frx":0019
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub
Private Sub Command2_Click()
a = "Rafasé_282 Dice: "
List1.AddItem (a) & (Text1.Text)
Text1.Text = ""
Text1.SetFocus
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
a = "Rafasé_282 Dice: "
List1.AddItem (a) & (Text1.Text)
Text1.Text = ""
Text1.SetFocus
End If
End Sub
