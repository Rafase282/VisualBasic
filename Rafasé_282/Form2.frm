VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   Caption         =   "Area del Triangulo By R.282"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4665
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   Palette         =   "Form2.frx":628A
   Picture         =   "Form2.frx":7AD0
   ScaleHeight     =   6000
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Comandos"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   2880
      TabIndex        =   7
      Top             =   480
      Width           =   1575
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limpiar"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Calcular"
         Height          =   255
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Valores"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Resultado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Valor H"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Valor B"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Menu file2 
      Caption         =   "&File"
      Begin VB.Menu forum 
         Caption         =   "&Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu k 
         Caption         =   "-"
      End
      Begin VB.Menu exit2 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu programs2 
      Caption         =   "&Programs"
      Begin VB.Menu ewt 
         Caption         =   "Tab&las"
         Begin VB.Menu tabladel12 
            Caption         =   "&Tabla del 1"
         End
         Begin VB.Menu r4ytre 
            Caption         =   "T&abla de Cualquier Numero"
         End
      End
      Begin VB.Menu gfj 
         Caption         =   "-"
      End
      Begin VB.Menu hgm 
         Caption         =   "Colores"
      End
      Begin VB.Menu promedio2 
         Caption         =   "&Promedio "
      End
      Begin VB.Menu cal 
         Caption         =   "&Calculadora"
      End
      Begin VB.Menu rewq 
         Caption         =   "&Calificaciones"
      End
      Begin VB.Menu adivina2 
         Caption         =   "&Adivina el Numero"
      End
      Begin VB.Menu areacu 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu dsb 
         Caption         =   "&Formula Cuadratica"
      End
      Begin VB.Menu trjst 
         Caption         =   "&Anuncio del CAFAM"
      End
      Begin VB.Menu orde 
         Caption         =   "&Ordene los Numeros"
      End
      Begin VB.Menu fdshgf 
         Caption         =   "Entrada de Estudiantes"
      End
      Begin VB.Menu dsfg 
         Caption         =   "Formula C=t+h*d-m/4 "
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adivina2_Click()
Form2.Hide
Form4.Show
End Sub

Private Sub areacu_Click()
Form2.Hide
Form10.Show
End Sub

Private Sub cal_Click()
Form2.Hide
Form15.Show
End Sub

Private Sub Command1_Click()
Dim b, h, a As Integer
b = Val(Text1.Text)
h = Val(Text2.Text)
a = b * h / 2
Text3.Text = (a)
End Sub

Private Sub Command2_Click()
Cls
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text1.SetFocus
End Sub

Private Sub dsb_Click()
Form2.Hide
Form14.Show
End Sub

Private Sub dsfg_Click()
Form2.Hide
Form12.Show
End Sub

Private Sub exit2_Click()
End
End Sub

Private Sub fdshgf_Click()
Form2.Hide
Form9.Show
End Sub

Private Sub forum_Click()
Form2.Hide
Form4.Show
End Sub

Private Sub hgm_Click()
Form2.Hide
Form13.Show
End Sub

Private Sub orde_Click()
Form2.Hide
Form11.Show
End Sub

Private Sub promedio2_Click()
Form2.Hide
Form5.Show
End Sub

Private Sub r4ytre_Click()
Form2.Hide
Form6.Show
End Sub

Private Sub rewq_Click()
Form2.Hide
Form8.Show
End Sub

Private Sub tabladel12_Click()
Form2.Hide
Form3.Show
End Sub

Private Sub trjst_Click()
Form2.Hide
Form7.Show
End Sub
