VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "La Tabla del 1 By R.282"
   ClientHeight    =   4560
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4530
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":628A
   ScaleHeight     =   4560
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   3000
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calcular"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Menu file3 
      Caption         =   "&File"
      Begin VB.Menu forum3 
         Caption         =   "&Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ty 
         Caption         =   "-"
      End
      Begin VB.Menu exit3 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu programs3 
      Caption         =   "&Programs"
      Begin VB.Menu rehy 
         Caption         =   "Tab&las"
         Begin VB.Menu dgfsadg 
            Caption         =   "&Tabla de cualquier Numero"
         End
      End
      Begin VB.Menu gjfg 
         Caption         =   "-"
      End
      Begin VB.Menu color 
         Caption         =   "Colores"
      End
      Begin VB.Menu promediestudiante3 
         Caption         =   "&Promedio "
      End
      Begin VB.Menu cal 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu rewg 
         Caption         =   "&Calificaciones"
      End
      Begin VB.Menu adivina 
         Caption         =   "&Adivina el Numero"
      End
      Begin VB.Menu areatriangulo3 
         Caption         =   "&Area del Triangulo"
      End
      Begin VB.Menu gfjfg 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu for 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu fdsh 
         Caption         =   "&Anuncio del CAFAM"
      End
      Begin VB.Menu orde 
         Caption         =   "&Ordene los Numeros"
      End
      Begin VB.Menu gfjgf 
         Caption         =   "Entrada de Estudiantes"
      End
      Begin VB.Menu dfs 
         Caption         =   "Formula C=t+h*d-m/4 "
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adivina_Click()
Form3.Hide
Form4.Show
End Sub

Private Sub areatriangulo3_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub cal_Click()
Form3.Hide
form15.Show
End Sub

Private Sub color_Click()
Form3.Hide
Form13.Show
End Sub

Private Sub Command1_Click()
Dim b, c As Integer
For a = 1 To 12
b = 1
c = b * a
Print b; "x"; a; "="; c
Next a
End Sub

Private Sub Command2_Click()
Cls
End Sub

Private Sub dfs_Click()
Form3.Hide
Form12.Show
End Sub

Private Sub dgfsadg_Click()
Form3.Hide
Form6.Show
End Sub

Private Sub exit3_Click()
End
End Sub

Private Sub fdsh_Click()
Form3.Hide
Form7.Show
End Sub

Private Sub for_Click()
Form3.Hide
form14.Show
End Sub

Private Sub forum3_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub gfjfg_Click()
Form3.Hide
Form10.Show
End Sub

Private Sub gfjgf_Click()
Form3.Hide
Form9.Show
End Sub

Private Sub orde_Click()
Form3.Hide
Form11.Show
End Sub

Private Sub promediestudiante3_Click()
Form3.Hide
Form5.Show
End Sub

Private Sub rewg_Click()
Form3.Hide
Form8.Show
End Sub
