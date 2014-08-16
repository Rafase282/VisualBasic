VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   Caption         =   "Adivina el Numero By R.282"
   ClientHeight    =   4785
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4080
   ForeColor       =   &H8000000E&
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":628A
   ScaleHeight     =   4785
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Comandos"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   2400
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rendir"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Calcular"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "  El No.1 x el No.2 / No.3 debe ser = a 16"
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Numeros"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Resultado"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000012&
         Caption         =   "Numero 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Numero 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Numero 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Menu file4 
      Caption         =   "&File"
      Begin VB.Menu forum4 
         Caption         =   "&Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu jk 
         Caption         =   "-"
      End
      Begin VB.Menu exit4 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu programs4 
      Caption         =   "&Programs"
      Begin VB.Menu wr 
         Caption         =   "&Tablas"
         Begin VB.Menu tabla14 
            Caption         =   "&Tabla del 1"
         End
         Begin VB.Menu tablani 
            Caption         =   "&Tabla de cualquier Numero"
         End
      End
      Begin VB.Menu gfjg 
         Caption         =   "-"
      End
      Begin VB.Menu color 
         Caption         =   "Colores"
      End
      Begin VB.Menu promedioestudiante 
         Caption         =   "&Promedio "
      End
      Begin VB.Menu cal 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu dfh 
         Caption         =   "&Calificaciones"
      End
      Begin VB.Menu area1 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu basetriangulo4 
         Caption         =   "&Area del Triangulo"
      End
      Begin VB.Menu forv 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu sdfg 
         Caption         =   "&Anuncio del CAFAM"
      End
      Begin VB.Menu orden 
         Caption         =   "&Ordene los Numeros"
      End
      Begin VB.Menu entrada 
         Caption         =   "Entrada de Estudiantes"
      End
      Begin VB.Menu rh 
         Caption         =   "Formula C=t+h*d-m/4 "
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub area1_Click()
Form4.Hide
Form10.Show
End Sub

Private Sub basetriangulo4_Click()
Form4.Hide
Form2.Show
End Sub

Private Sub cal_Click()
Form4.Hide
Form15.Show
End Sub

Private Sub color_Click()
Form4.Hide
Form13.Show
End Sub

Private Sub Command1_Click()
Dim b, c, a, d As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = (a * b) / c
Text5.Text = (d)
If d = 16 Then
MsgBox ("Felicidades el numero es correcto")
Else
MsgBox ("El Numero es Incorrecto")
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Text1.Text = "8"
Text2.Text = "6"
Text3.Text = "3"
End Sub

Private Sub dfh_Click()
Form4.Hide
Form8.Show
End Sub

Private Sub entrada_Click()
Form4.Hide
Form9.Show
End Sub

Private Sub exit4_Click()
End
End Sub

Private Sub forum4_Click()
Form4.Hide
Form1.Show
End Sub

Private Sub forv_Click()
Form4.Hide
Form14.Show
End Sub

Private Sub orden_Click()
Form4.Hide
Form11.Show
End Sub

Private Sub promedioestudiante_Click()
Form4.Hide
Form5.Show
End Sub

Private Sub rh_Click()
Form4.Hide
Form12.Show
End Sub

Private Sub sdfg_Click()
Form4.Hide
Form7.Show
End Sub

Private Sub tabla14_Click()
Form4.Hide
Form1.Show
End Sub

Private Sub tablani_Click()
Form4.Hide
Form6.Show
End Sub
