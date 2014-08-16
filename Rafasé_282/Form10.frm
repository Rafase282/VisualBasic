VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H80000008&
   Caption         =   "Area del Cuadrado By R.282"
   ClientHeight    =   5190
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4500
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   Picture         =   "Form10.frx":628A
   ScaleHeight     =   5190
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "Comandos"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   4440
      Width           =   3375
      Begin VB.CommandButton Command2 
         Caption         =   "Limpiar"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calcular"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      Caption         =   "Datos"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   3360
      Width           =   3375
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Area  Cuadrado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Area de un Lado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Menu trwj 
      Caption         =   "&File"
      Begin VB.Menu ery 
         Caption         =   "&Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu fyu 
         Caption         =   "-"
      End
      Begin VB.Menu rtjht 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu ehy 
      Caption         =   "&Programs"
      Begin VB.Menu reger 
         Caption         =   "&Tablas"
         Begin VB.Menu rhyt 
            Caption         =   "Tabla del 1"
         End
         Begin VB.Menu trujhrt 
            Caption         =   "&Tabla de Cualquier Numero"
         End
      End
      Begin VB.Menu gfj 
         Caption         =   "-"
      End
      Begin VB.Menu gr 
         Caption         =   "Colores"
      End
      Begin VB.Menu tht 
         Caption         =   "&Promedio"
      End
      Begin VB.Menu ewg 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu rehy 
         Caption         =   "&Calificaciones"
      End
      Begin VB.Menu yre 
         Caption         =   "&Adivina el Numero"
      End
      Begin VB.Menu trhtr 
         Caption         =   "&Area del Triangulo"
      End
      Begin VB.Menu rg 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu tyht 
         Caption         =   "&Anuncio del CAFAM"
      End
      Begin VB.Menu orde 
         Caption         =   "&Ordene los Numeros"
      End
      Begin VB.Menu uoi 
         Caption         =   "Entrada de Estudiante"
      End
      Begin VB.Menu formula 
         Caption         =   "Formula C=t+h*d-m/4 "
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b As Integer
a = Val(Text1.Text)
b = a * 4
Text2.Text = (b)
End Sub

Private Sub Command2_Click()
Cls
Text1.Text = " "
Text2.Text = " "
Text1.SetFocus
End Sub

Private Sub ery_Click()
Form10.Hide
Form1.Show
End Sub

Private Sub ewg_Click()
Form10.Hide
Form15.Show
End Sub

Private Sub formula_Click()
Form10.Hide
Form12.Show
End Sub

Private Sub gr_Click()
Form10.Hide
Form13.Show
End Sub

Private Sub orde_Click()
Form10.Hide
Form11.Show
End Sub

Private Sub rehy_Click()
Form10.Hide
Form8.Show
End Sub

Private Sub rg_Click()
Form10.Hide
Form14.Show
End Sub

Private Sub rhyt_Click()
Form10.Hide
Form3.Hide
End Sub

Private Sub rtjht_Click()
End
End Sub

Private Sub tht_Click()
Form10.Hide
Form5.Show
End Sub

Private Sub trhtr_Click()
Form10.Hide
Form2.Show
End Sub

Private Sub trujhrt_Click()
Form10.Hide
Form6.Show
End Sub

Private Sub tyht_Click()
Form10.Hide
Form7.Show
End Sub

Private Sub uoi_Click()
Form10.Hide
Form9.Show
End Sub

Private Sub yre_Click()
Form10.Hide
Form4.Show
End Sub
