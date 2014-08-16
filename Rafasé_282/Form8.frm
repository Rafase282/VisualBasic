VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H80000007&
   Caption         =   "Calificaciones By R.282"
   ClientHeight    =   5490
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3660
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   Picture         =   "Form8.frx":628A
   ScaleHeight     =   5490
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Comandos"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   2640
      TabIndex        =   11
      Top             =   3360
      Width           =   975
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limpiar"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Calcular"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Datos"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   2535
      Begin VB.TextBox nota 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox ef 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox ep 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox tp 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox asi 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Nota Final"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Examen Final"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Examan Parcial"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Trabajo Pactico"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Asistencia"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Menu rey 
      Caption         =   "&File"
      Begin VB.Menu efe 
         Caption         =   "&Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ew 
         Caption         =   "-"
      End
      Begin VB.Menu reyry 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu rewg 
      Caption         =   "&Programs"
      Begin VB.Menu werg 
         Caption         =   "&Tablas"
         Begin VB.Menu resad 
            Caption         =   "&Tabla del 1"
         End
         Begin VB.Menu ewrt 
            Caption         =   "&Tabla de Cualquier No."
         End
      End
      Begin VB.Menu hg 
         Caption         =   "-"
      End
      Begin VB.Menu cdf 
         Caption         =   "Colores"
      End
      Begin VB.Menu tuse 
         Caption         =   "&Promedio"
      End
      Begin VB.Menu dsgvds 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu yj 
         Caption         =   "&Adivina el Numero"
      End
      Begin VB.Menu kyt 
         Caption         =   "&Area del Triangulo"
      End
      Begin VB.Menu o 
         Caption         =   "&Area del Cuadrado"
      End
      Begin VB.Menu uk 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu refdg 
         Caption         =   "&Anuncio del CAFAM"
      End
      Begin VB.Menu orde 
         Caption         =   "&Ordene los Numeros"
      End
      Begin VB.Menu jh 
         Caption         =   "Entrada de Estudiantes"
      End
      Begin VB.Menu reh 
         Caption         =   "Formula C=t+ h * d  -m /4"
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cdf_Click()
Form8.Hide
Form13.Show
End Sub

Private Sub Command1_Click()
Dim a, b, c, d As Integer
a = Val(asi.Text)
b = Val(tp.Text)
c = Val(ep.Text)
d = Val(ef.Text)
e = (a + b + c + d) / 4
If e >= 89 And e <= 100 Then
nota.Text = "A"
ElseIf e >= 79 And e < 89 Then
nota.Text = "B"
ElseIf e >= 75 And e < 79 Then
nota.Text = "C"
ElseIf e >= 70 And e < 75 Then
nota.Text = "D"
ElseIf e >= 59 And e < 70 Then
nota.Text = "Fc"
ElseIf e < 59 Then
nota.Text = "E"
End If
End Sub

Private Sub Command2_Click()
Cls
asi.Text = " "
tp.Text = " "
ep.Text = " "
ef.Text = " "
nota.Text = " "
asi.SetFocus
End Sub

Private Sub dsgvds_Click()
Form8.Hide
form15.Show
End Sub

Private Sub efe_Click()
Form8.Hide
Form1.Show
End Sub

Private Sub ewrt_Click()
Form8.Hide
Form6.Show
End Sub

Private Sub jh_Click()
Form8.Hide
Form9.Show
End Sub

Private Sub kyt_Click()
Form8.Hide
Form2.Show
End Sub

Private Sub o_Click()
Form8.Hide
Form10.Show
End Sub

Private Sub orde_Click()
Form8.Hide
Form11.Show
End Sub

Private Sub refdg_Click()
Form8.Hide
Form7.Show
End Sub

Private Sub reh_Click()
Form8.Hide
Form12.Show
End Sub

Private Sub resad_Click()
Form8.Hide
Form3.Hide
End Sub

Private Sub reyry_Click()
End
End Sub

Private Sub tuse_Click()
Form8.Hide
Form5.Show
End Sub

Private Sub uk_Click()
Form8.Hide
form14.Show
End Sub

Private Sub yj_Click()
Form8.Hide
Form4.Show
End Sub
