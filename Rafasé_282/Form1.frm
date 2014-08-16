VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Rafasé_282  (R.282)"
   ClientHeight    =   2700
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3345
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":628A
   ScaleHeight     =   2700
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   1200
   End
   Begin VB.Image Image3 
      Height          =   465
      Left            =   240
      Picture         =   "Form1.frx":4C331
      Top             =   120
      Width           =   2835
   End
   Begin VB.Image Image2 
      Height          =   1635
      Left            =   720
      Picture         =   "Form1.frx":4CB2B
      Top             =   600
      Width           =   1785
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   600
      Picture         =   "Form1.frx":4D908
      Top             =   720
      Width           =   2130
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu programs 
      Caption         =   "&Programs"
      Begin VB.Menu req 
         Caption         =   "&Tablas"
         Begin VB.Menu tabla1 
            Caption         =   "&Tabla del 1"
         End
         Begin VB.Menu rhrew 
            Caption         =   "&Tabla de Cualquier Numero"
         End
      End
      Begin VB.Menu gfsdj 
         Caption         =   "-"
      End
      Begin VB.Menu color 
         Caption         =   "&Colores"
      End
      Begin VB.Menu ew 
         Caption         =   "&Promedio "
      End
      Begin VB.Menu calculadora 
         Caption         =   "&Calculadora"
      End
      Begin VB.Menu rewgt 
         Caption         =   "&Calificaciones"
      End
      Begin VB.Menu adivina 
         Caption         =   "&Adivina el Numero"
      End
      Begin VB.Menu basetriangulo 
         Caption         =   "&Area del Triangulo"
      End
      Begin VB.Menu fdsg 
         Caption         =   "&Area del Cuadrado"
      End
      Begin VB.Menu cuadr 
         Caption         =   "&Formula Cuadratica"
      End
      Begin VB.Menu drhgar 
         Caption         =   "&Anuncio del CAFAM"
      End
      Begin VB.Menu orden 
         Caption         =   "&Ordene los Numeros"
      End
      Begin VB.Menu ewtq 
         Caption         =   "&Entrada de Estudiante"
      End
      Begin VB.Menu formula 
         Caption         =   "&Formula C= t+h*d-m/4 "
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub adivina_Click()
Form1.Hide
Form4.Show
End Sub

Private Sub basetriangulo_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub calculadora_Click()
Form1.Hide
Form15.Show
End Sub

Private Sub color_Click()
Form1.Hide
Form13.Show
End Sub

Private Sub cuadr_Click()
Form1.Hide
Form14.Show
End Sub

Private Sub drhgar_Click()
Form1.Hide
Form7.Show
End Sub

Private Sub ew_Click()
Form1.Hide
Form5.Show
End Sub

Private Sub ewtq_Click()
Form1.Hide
Form9.Show
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub fdsg_Click()
Form1.Hide
Form10.Show
End Sub

Private Sub formula_Click()
Form1.Hide
Form12.Show
End Sub



Private Sub orden_Click()
Form1.Hide
Form11.Show
End Sub

Private Sub rewgt_Click()
Form1.Hide
Form8.Show
End Sub

Private Sub rhrew_Click()
Form1.Hide
Form6.Show
End Sub

Private Sub tabla1_Click()
Form1.Hide
Form3.Show
End Sub

Private Sub Timer1_Timer()
If Image2.Visible = False Then
Image2.Visible = True
ElseIf Image2.Visible = True Then
Image2.Visible = False
End If
End Sub
