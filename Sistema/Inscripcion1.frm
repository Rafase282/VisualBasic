VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form3"
   ScaleHeight     =   8595
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   255
      Left            =   1200
      TabIndex        =   23
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   1920
      TabIndex        =   22
      Top             =   7800
      Width           =   2415
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   1920
      TabIndex        =   20
      Top             =   7320
      Width           =   3735
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   6840
      Width           =   3735
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   6360
      Width           =   3735
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   5880
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   1920
      TabIndex        =   12
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   1920
      TabIndex        =   10
      Top             =   4920
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   4440
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3960
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3360
      Width           =   3735
   End
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   360
      Picture         =   "Inscripcion1.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   5115
      TabIndex        =   1
      Top             =   1560
      Width           =   5175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   2040
      Picture         =   "Inscripcion1.frx":03E4
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Telefono:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Direccion:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Lugar de Trabajo"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Profesion u Ocupacion"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Nombre dela Madre"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Telefono:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Direccion:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Lugar de Trabajo"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Profesion u Ocupacion"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre del Padre"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Datos Familiares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Hide
Form4.Show
End Sub
