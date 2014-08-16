VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H80000012&
   Caption         =   "Entrada de Estudiantes By R.282"
   ClientHeight    =   5730
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4515
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":628A
   ScaleHeight     =   5730
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Comandos"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   4095
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limpiar"
         Height          =   195
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Entrar"
         Height          =   195
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Datos 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   4095
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Ult. Nom."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Ult Numero:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Datos 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Nombre:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu forum 
         Caption         =   "&Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu raya 
         Caption         =   "-"
      End
      Begin VB.Menu azfd 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu programs 
      Caption         =   "&Programs"
      Begin VB.Menu tablas 
         Caption         =   "Tablas"
         Begin VB.Menu tabladel1 
            Caption         =   "Tabla del 1"
         End
         Begin VB.Menu tablacualquier 
            Caption         =   "Tabla de Cualquier Numero"
         End
      End
      Begin VB.Menu p 
         Caption         =   "-"
      End
      Begin VB.Menu color 
         Caption         =   "Colores"
      End
      Begin VB.Menu jgf 
         Caption         =   "Promedio"
      End
      Begin VB.Menu dsfg 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu gf 
         Caption         =   "Calificaciones"
      End
      Begin VB.Menu adivina 
         Caption         =   "Adivina el Numero"
      End
      Begin VB.Menu area 
         Caption         =   "Area del Triangulo"
      End
      Begin VB.Menu areacuadrado 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu sfdg 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu gfj 
         Caption         =   "Anuncio del CAFAM"
      End
      Begin VB.Menu orde 
         Caption         =   "&Ordene los Numeros"
      End
      Begin VB.Menu tjh 
         Caption         =   "Formula C=t+h*d-m/4"
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adivina_Click()
Form9.Hide
Form4.Show
End Sub

Private Sub area_Click()
Form9.Hide
Form2.Show
End Sub

Private Sub areacuadrado_Click()
Form9.Hide
Form10.Show
End Sub

Private Sub azfd_Click()
End
End Sub

Private Sub color_Click()
Form9.Hide
Form13.Show
End Sub

Private Sub Command1_Click()
Dim a, b, c As Integer
a = 1
a = Val(Text2.Text) + a
Text2.Text = (a)
Text3.Text = Text1.Text
End Sub


Private Sub Command2_Click()
Cls
Text1.Text = " "
Text1.SetFocus
End Sub

Private Sub dsfg_Click()
Form9.Hide
form15.Show
End Sub

Private Sub forum_Click()
Form9.Hide
Form1.Show
End Sub

Private Sub gf_Click()
Form9.Hide
Form8.Show
End Sub

Private Sub gfj_Click()
Form9.Hide
Form7.Show
End Sub

Private Sub jgf_Click()
Form9.Hide
Form5.Show
End Sub

Private Sub orde_Click()
Form9.Hide
Form11.Show
End Sub

Private Sub sfdg_Click()
Form9.Hide
form14.Show
End Sub

Private Sub tablacualquier_Click()
Form9.Hide
Form6.Show
End Sub

Private Sub tabladel1_Click()
Form9.Hide
Form3.Show
End Sub

Private Sub tjh_Click()
Form9.Hide
Form12.Show
End Sub
