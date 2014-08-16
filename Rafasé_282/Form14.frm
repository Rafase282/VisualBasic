VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H00000000&
   Caption         =   "Formula Cuadratica By R.282"
   ClientHeight    =   7305
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5325
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   Picture         =   "Form14.frx":628A
   ScaleHeight     =   7305
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Comandos"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   5055
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limpiar"
         Height          =   255
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Calcular"
         Height          =   255
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Soluciones"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   5055
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "X2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "X1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Datos"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   480
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Formula"
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   2655
         Begin VB.PictureBox Picture1 
            Height          =   735
            Left            =   120
            Picture         =   "Form14.frx":C6F2
            ScaleHeight     =   675
            ScaleWidth      =   2235
            TabIndex        =   2
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "C="
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "B="
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "A="
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu dgf 
         Caption         =   "Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu h 
         Caption         =   "-"
      End
      Begin VB.Menu gfgf 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu programs 
      Caption         =   "&Programs"
      Begin VB.Menu Taba 
         Caption         =   "Tablas"
         Begin VB.Menu tabla1 
            Caption         =   "Tabla del 1"
         End
         Begin VB.Menu tabla2 
            Caption         =   "Tabla de Cualquier Numero"
         End
      End
      Begin VB.Menu trh 
         Caption         =   "-"
      End
      Begin VB.Menu colores 
         Caption         =   "Colores"
      End
      Begin VB.Menu promedio 
         Caption         =   "Promedio"
      End
      Begin VB.Menu calculadora 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu calificaciones 
         Caption         =   "Calificaciones"
      End
      Begin VB.Menu adivinaelnumero 
         Caption         =   "Adivina el Numero"
      End
      Begin VB.Menu areatriangilo 
         Caption         =   "Area del Triangulo"
      End
      Begin VB.Menu areadelcuadrado 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu anunciocafam 
         Caption         =   "Anuncio del CAFAM"
      End
      Begin VB.Menu ordenelosnumeros 
         Caption         =   "Ordene los Numeros"
      End
      Begin VB.Menu entrada 
         Caption         =   "Entrada de  Estudiantes"
      End
      Begin VB.Menu formulac 
         Caption         =   "Formula C= t+h*d - m / 4 "
      End
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adivinaelnumero_Click()
Form14.Hide
Form4.Show
End Sub

Private Sub anunciocafam_Click()
Form14.Hide
Form7.Show
End Sub

Private Sub areadelcuadrado_Click()
Form14.Hide
Form10.Show
End Sub

Private Sub areatriangilo_Click()
Form14.Hide
Form2.Show
End Sub

Private Sub calculadora_Click()
Form14.Hide
Form15.Show
End Sub

Private Sub calificaciones_Click()
Form14.Hide
Form8.Show
End Sub

Private Sub colores_Click()
Form14.Hide
Form13.Show
End Sub

Private Sub Command1_Click()
Dim a, b, c, x, y As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
x = -b + Sqr(b * b) + (-4 * a * c) / (2 * a)
y = -b - Sqr(b * b) + (-4 * a * c) / (2 * a)
Text4.Text = (x)
Text5.Text = (y)
End Sub

Private Sub Command2_Click()
Cls
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text1.SetFocus
End Sub

Private Sub dgf_Click()
Form14.Hide
Form1.Show
End Sub

Private Sub entrada_Click()
Form14.Hide
Form9.Show
End Sub

Private Sub formulac_Click()
Form14.Hide
Form12.Show
End Sub

Private Sub gfgf_Click()
End
End Sub

Private Sub ordenelosnumeros_Click()
Form14.Hide
Form11.Show
End Sub

Private Sub promedio_Click()
Form14.Hide
Form5.Show
End Sub

Private Sub tabla1_Click()
Form14.Hide
Form3.Show
End Sub

Private Sub tabla2_Click()
Form14.Hide
Form6.Show
End Sub
