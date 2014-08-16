VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   Caption         =   "Promedio del Estudiante By R.282"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4560
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":628A
   ScaleHeight     =   6420
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Comandos"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   11
      Top             =   5520
      Width           =   3855
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limpiar"
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Calcular"
         Height          =   375
         Left            =   480
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
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000012&
         Caption         =   "Promedio"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000012&
         Caption         =   "Nota 4"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000012&
         Caption         =   "Nota 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Nota 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Nota 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu foro 
         Caption         =   "&Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ta 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu programs 
      Caption         =   "&Programs"
      Begin VB.Menu tie 
         Caption         =   "&Tablas"
         Begin VB.Menu tabladel1 
            Caption         =   "&Tabla del 1"
         End
         Begin VB.Menu tablano 
            Caption         =   "&Tabla de Cualquier Numero"
         End
      End
      Begin VB.Menu gfrg 
         Caption         =   "-"
      End
      Begin VB.Menu colcio 
         Caption         =   "Colores"
      End
      Begin VB.Menu uhjrt 
         Caption         =   "&Calificaciones"
      End
      Begin VB.Menu calcio 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu adivina 
         Caption         =   "&Adivina el Numero"
      End
      Begin VB.Menu areatriangu 
         Caption         =   "&Area del Triangulo"
      End
      Begin VB.Menu dsfg 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu fro 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu HGK 
         Caption         =   "Anuncio del CAFAM"
      End
      Begin VB.Menu orden 
         Caption         =   "&Ordene los Numeros"
      End
      Begin VB.Menu fad 
         Caption         =   "Entrade de Estudiantes"
      End
      Begin VB.Menu rg 
         Caption         =   "Formula C=t+h*d-m/4 "
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adivina_Click()
Form5.Hide
Form4.Show
End Sub

Private Sub areatriangu_Click()
Form5.Hide
Form2.Show
End Sub

Private Sub calcio_Click()
Form5.Hide
form15.Show
End Sub

Private Sub colcio_Click()
Form5.Hide
Form13.Show
End Sub

Private Sub Command1_Click()
Dim a, b, c, d As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Text4.Text)
e = (a + b + c + d) / 4
Text5.Text = (e)
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

Private Sub dsfg_Click()
Form5.Hide
Form10.Show
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub fad_Click()
Form5.Hide
Form9.Show
End Sub

Private Sub foro_Click()
Form5.Hide
Form1.Show
End Sub

Private Sub fro_Click()
Form5.Hide
form14.Show
End Sub

Private Sub HGK_Click()
Form5.Hide
Form7.Show
End Sub

Private Sub orden_Click()
Form5.Hide
Form11.Show
End Sub

Private Sub rg_Click()
Form5.Hide
Form12.Show
End Sub

Private Sub tabladel1_Click()
Form5.Hide
Form3.Show
End Sub

Private Sub tablano_Click()
Form5.Hide
Form6.Show
End Sub

Private Sub uhjrt_Click()
Form5.Hide
Form8.Show
End Sub
