VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Colores By R.282"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Height          =   405
      Index           =   2
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0"
      Top             =   2760
      Width           =   495
   End
   Begin VB.HScrollBar hsbColor 
      Height          =   375
      Index           =   2
      LargeChange     =   16
      Left            =   720
      Max             =   255
      TabIndex        =   12
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   1
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   2280
      Width           =   495
   End
   Begin VB.HScrollBar hsbColor 
      Height          =   375
      Index           =   1
      LargeChange     =   16
      Left            =   720
      Max             =   255
      TabIndex        =   9
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   1800
      Width           =   495
   End
   Begin VB.HScrollBar hsbColor 
      Height          =   375
      Index           =   0
      LargeChange     =   16
      Left            =   720
      Max             =   255
      TabIndex        =   5
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
      Begin VB.OptionButton optColor 
         Caption         =   "Texto"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Fondo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Azul"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label V 
      Caption         =   "Verde"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label R 
      Caption         =   "Rojo"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Rafasé_282 (R.282)"
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public brojo, bverde, bazul As Integer
Public frojo, fverde, fazul As Integer
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
brojo = 0
bverde = 0
bazul = 0
frojo = 255
fverde = 255
fazul = 255
Label1.BackColor = RGB(brojo, bverde, bazul)
Label1.ForeColor = RGB(frojo, fverde, fazul)
End Sub

Private Sub hsbColor_Change(Index As Integer)
If optColor(0).Value = True Then
Label1.BackColor = RGB(hsbColor(0).Value, hsbColor(1).Value, hsbColor(2).Value)
Dim i As Integer
For i = 0 To 2
Text(i).Text = hsbColor(i).Value
Next i
Else
Label1.ForeColor = RGB(hsbColor(0).Value, hsbColor(1).Value, hsbColor(2).Value)
For i = 0 To 2
Text(i).Text = hsbColor(i).Value
Next i
End If
End Sub

Private Sub Option_Click(Index As Integer)
If Index = 0 Then 'Se pasa a cambiar el fondo
frojo = hsbColor(0).Value
fverde = hsbColor(1).Value
fazul = hsbColor(2).Value
hsbColor(0).Value = brojo
hsbColor(1).Value = bverde
hsbColor(2).Value = bazul
Else
brojo = hsbColor(0).Value
bverde = hsbColor(1).Value
bazul = hsbColor(2).Value
hsbColor(0).Value = frojo
hsbColor(1).Value = fverde
hsbColor(2).Value = fazul
End If
End Sub
