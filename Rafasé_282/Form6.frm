VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Tabla de Cualquier #  By R.282"
   ClientHeight    =   4365
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4515
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":628A
   ScaleHeight     =   4365
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calcular"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Text            =   "Numero"
      Top             =   960
      Width           =   855
   End
   Begin VB.Menu file6 
      Caption         =   "&File"
      Begin VB.Menu forum6 
         Caption         =   "&Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sd 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu pro 
      Caption         =   "&Programs"
      Begin VB.Menu fhg 
         Caption         =   "&Tablas"
         Begin VB.Menu tabla 
            Caption         =   "&Tabla del 1"
         End
      End
      Begin VB.Menu erwtg 
         Caption         =   "-"
      End
      Begin VB.Menu color 
         Caption         =   "Colores"
      End
      Begin VB.Menu promedio 
         Caption         =   "&Promedio"
      End
      Begin VB.Menu calcu 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu rehy 
         Caption         =   "Calificaciones"
      End
      Begin VB.Menu adivina 
         Caption         =   "Adivina el Numero"
      End
      Begin VB.Menu area 
         Caption         =   "Area del Triangulo"
      End
      Begin VB.Menu ef 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu fdh 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu reyhre 
         Caption         =   "Anuncio del CAFAM"
      End
      Begin VB.Menu ds 
         Caption         =   "&Ordene los Numeros"
      End
      Begin VB.Menu ewf 
         Caption         =   "Entrada de Estudiantes"
      End
      Begin VB.Menu rfh 
         Caption         =   "Formula C=t+h*d-m/4"
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adivina_Click()
Form6.Hide
Form4.Show
End Sub

Private Sub area_Click()
Form6.Hide
Form2.Show
End Sub

Private Sub calcu_Click()
Form6.Hide
form15.Show
End Sub

Private Sub color_Click()
Form6.Hide
Form13.Show
End Sub

Private Sub Command1_Click()
Dim a As Integer
a = Val(Text1.Text)
For b = 1 To 12
c = a * b
Print a; "x"; b; "="; c
Next b
End Sub

Private Sub Command2_Click()
Cls
Text1.Text = "Numero"
End Sub

Private Sub ds_Click()
Form6.Hide
Form11.Show
End Sub

Private Sub ef_Click()
Form6.Hide
Form10.Show
End Sub

Private Sub ewf_Click()
Form6.Hide
Form9.Show
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub fdh_Click()
Form6.Hide
form14.Show
End Sub

Private Sub forum6_Click()
Form6.Hide
Form1.Show
End Sub

Private Sub promedio_Click()
Form6.Hide
Form5.Show
End Sub

Private Sub rehy_Click()
Form6.Hide
Form8.Show
End Sub

Private Sub reyhre_Click()
Form6.Hide
Form7.Show
End Sub

Private Sub rfh_Click()
Form6.Hide
Form12.Show
End Sub

Private Sub tabla_Click()
Form6.Hide
Form3.Show
End Sub
