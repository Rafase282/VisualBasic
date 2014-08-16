VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   Caption         =   "Tabla En List By R.282"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "TABLAE~1.frx":0000
      Left            =   2400
      List            =   "TABLAE~1.frx":0002
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Numero"
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C00000&
      Caption         =   "Calcular"
      Height          =   495
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.Clear
Dim a, b, c As Integer
For b = 1 To 12
a = Val(Text1.Text)
c = a * b
x = "x"
d = "="
List1.AddItem (a) & (x) & (b) & (d) & (c)
Next b
End Sub


