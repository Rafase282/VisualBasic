VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00000000&
   Caption         =   "Formula C=t+h*d-m/4 By R.282"
   ClientHeight    =   5340
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4620
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   Picture         =   "Form12.frx":628A
   ScaleHeight     =   5340
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Comandos"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3000
      TabIndex        =   11
      Top             =   4200
      Width           =   1455
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limpiar"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Calcular"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Datos"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   2775
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "C"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "M"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "D"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "H"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "T"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu foro 
         Caption         =   "&Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu h 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu pro 
      Caption         =   "&Programs"
      Begin VB.Menu tablas 
         Caption         =   "Tablas"
         Begin VB.Menu tabla1 
            Caption         =   "Tabla del 1"
         End
         Begin VB.Menu tabla2 
            Caption         =   "Tabla de Cualquier Numero"
         End
      End
      Begin VB.Menu fs 
         Caption         =   "-"
      End
      Begin VB.Menu color 
         Caption         =   "Colores"
      End
      Begin VB.Menu promedio 
         Caption         =   "Promedio"
      End
      Begin VB.Menu cal 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu cali 
         Caption         =   "Calificaciones"
      End
      Begin VB.Menu adi 
         Caption         =   "Adivina el Numero"
      End
      Begin VB.Menu tri 
         Caption         =   "Area del Triangulo"
      End
      Begin VB.Menu cua 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu forcu 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu cafam 
         Caption         =   "Anuncio del  CAFAM"
      End
      Begin VB.Menu orde 
         Caption         =   "Ordene los Numeros"
      End
      Begin VB.Menu intro 
         Caption         =   "Entrada de Estudiantes"
      End
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adi_Click()
Form12.Hide
Form4.Show
End Sub

Private Sub cafam_Click()
Form12.Hide
Form7.Show
End Sub

Private Sub cal_Click()
Form12.Hide
Form15.Show
End Sub

Private Sub cali_Click()
Form12.Hide
Form8.Show
End Sub

Private Sub color_Click()
Form12.Hide
Form13.Show
End Sub

Private Sub Command1_Click()
Dim t, h, d, m As Integer
t = Val(Text1.Text)
h = Val(Text2.Text)
d = Val(Text3.Text)
m = Val(Text4.Text)
c = (t + h) * (d - m) / 4
Text5.Text = (c)
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

Private Sub cua_Click()
Form12.Hide
Form10.Hide
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub forcu_Click()
Form12.Hide
Form14.Show
End Sub

Private Sub foro_Click()
Form12.Hide
Form1.Show
End Sub

Private Sub intro_Click()
Form12.Hide
Form9.Show
End Sub

Private Sub orde_Click()
Form12.Hide
Form11.Show
End Sub

Private Sub promedio_Click()
Form12.Hide
Form5.Show
End Sub

Private Sub tabla1_Click()
Form12.Hide
Form2.Show
End Sub

Private Sub tabla2_Click()
Form12.Hide
Form6.Show
End Sub

Private Sub tri_Click()
Form12.Hide
Form2.Show
End Sub
