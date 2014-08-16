VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00000000&
   Caption         =   "Ordene los Numeros By R.282"
   ClientHeight    =   5625
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4200
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":628A
   ScaleHeight     =   5625
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "   Ni el del medio quede junto al ultimo o al  primero"
      Top             =   4440
      Width           =   3855
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "      Debes Ordenar los Numeros de tal forma que"
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Acciones"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   3855
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rendir"
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Numeros"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu forum 
         Caption         =   "Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu l 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu programs 
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
      Begin VB.Menu ll 
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
      Begin VB.Menu calificaciones 
         Caption         =   "Calificaciones"
      End
      Begin VB.Menu adivina 
         Caption         =   "Adivina el Numero"
      End
      Begin VB.Menu area 
         Caption         =   "Area del Triangulo"
      End
      Begin VB.Menu are 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu fyt 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu anuncio 
         Caption         =   "Anuncio del CAFAM"
      End
      Begin VB.Menu entrada 
         Caption         =   "Entrada de Estudiantes"
      End
      Begin VB.Menu formula 
         Caption         =   "Formula C= t+h*d-m/4 "
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adivina_Click()
Form11.Hide
Form4.Hide
End Sub

Private Sub anuncio_Click()
Form11.Hide
Form7.Show
End Sub

Private Sub are_Click()
Form11.Hide
Form10.Show
End Sub

Private Sub area_Click()
Form11.Hide
Form2.Show
End Sub

Private Sub cal_Click()
Form11.Hide
Form15.Show
End Sub

Private Sub calificaciones_Click()
Form11.Hide
Form8.Show
End Sub

Private Sub color_Click()
Form11.Hide
Form13.Show
End Sub

Private Sub Command1_Click()
Dim a, b, c, d, e, f, g As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Text4.Text)
e = Val(Text5.Text)
f = Val(Text6.Text)
g = Val(Text7.Text)
If a = 3 Or a = 1 And b = 1 Or b = 5 And c = 5 Or c = 3 And d = 7 And e = 4 Or e = 2 And f = 6 Or f = 4 And g = 2 Or g = 6 Then
MsgBox ("Felicidades, Combinacion Correcta")
Else
MsgBox ("La combinacion es Incorrecta")
End If
End Sub

Private Sub Command2_Click()
Text1.Text = 3
Text2.Text = 1
Text3.Text = 5
Text4.Text = 7
Text5.Text = 4
Text6.Text = 6
Text7.Text = 2
End Sub

Private Sub entrada_Click()
Form11.Hide
Form9.Show
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub formula_Click()
Form11.Hide
Form12.Show
End Sub

Private Sub forum_Click()
Form11.Hide
Form1.Show
End Sub

Private Sub fyt_Click()
Form11.Hide
Form14.Show
End Sub

Private Sub promedio_Click()
Form11.Hide
Form5.Show
End Sub

Private Sub tabla1_Click()
Form11.Hide
Form3.Show

End Sub

Private Sub tabla2_Click()
Form11.Hide
Form6.Show
End Sub
