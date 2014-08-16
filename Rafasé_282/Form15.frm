VERSION 5.00
Begin VB.Form Form15 
   BackColor       =   &H00000000&
   Caption         =   "Calculadora By R.282"
   ClientHeight    =   4260
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4305
   Icon            =   "Form15.frx":0000
   LinkTopic       =   "Form15"
   Picture         =   "Form15.frx":628A
   ScaleHeight     =   4260
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Comandos"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   2520
      TabIndex        =   14
      Top             =   2520
      Width           =   1455
      Begin VB.CommandButton mod 
         Caption         =   "MOD"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command19 
         Caption         =   "OFF"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton igual 
         Caption         =   "="
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton divent 
         Caption         =   "\"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton divdec 
         Caption         =   "/"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton multi 
         Caption         =   "x"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton resta 
         Caption         =   "-"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton suma 
         Caption         =   "+"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Botones"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
      Begin VB.CommandButton Command12 
         Caption         =   "."
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "AC"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton N0 
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton N9 
         Caption         =   "9"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton N8 
         Caption         =   "8"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton N7 
         Caption         =   "7"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton N6 
         Caption         =   "6"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton N5 
         Caption         =   "5"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton N4 
         Caption         =   "4"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton N3 
         Caption         =   "3"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton N2 
         Caption         =   "2"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton N1 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.Menu edsgf 
      Caption         =   "&File"
      Begin VB.Menu dsgf 
         Caption         =   "Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu hgk 
         Caption         =   "-"
      End
      Begin VB.Menu jd 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu rfgh 
      Caption         =   "&Programs"
      Begin VB.Menu fas 
         Caption         =   "Tablas"
         Begin VB.Menu sdgf 
            Caption         =   "Tabla del 1"
         End
         Begin VB.Menu jh 
            Caption         =   "Tabla de Cualquier Numero"
         End
      End
      Begin VB.Menu jrt 
         Caption         =   "-"
      End
      Begin VB.Menu reg 
         Caption         =   "Colores"
      End
      Begin VB.Menu thst 
         Caption         =   "Promedio"
      End
      Begin VB.Menu thrt 
         Caption         =   "Calificaciones"
      End
      Begin VB.Menu th 
         Caption         =   "Adivina el Numero"
      End
      Begin VB.Menu htrtr 
         Caption         =   "Area del Triangulo"
      End
      Begin VB.Menu thth 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu tht 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu ghjf 
         Caption         =   "Anuncio del CAFAM"
      End
      Begin VB.Menu rgrrfg 
         Caption         =   "Ordene los Numeros"
      End
      Begin VB.Menu rgr 
         Caption         =   "Entrada de  Estudiantes"
      End
      Begin VB.Menu gfnjh 
         Caption         =   "Formula C= t + h * d - m / 4 "
      End
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OPERA As Byte
Dim NUM1, NUM2 As Double

Private Sub Command11_Click()
Text1.Text = " "
OPERA = 0
NUM1 = 0
NUM2 = 0

End Sub

Private Sub Command12_Click()
Text1.Text = Text1.Text + "."
End Sub

Private Sub Command19_Click()
End
End Sub

Private Sub divdec_Click()
NUM1 = Val(Text1.Text)
OPERA = 4
Text1.Text = " "
End Sub

Private Sub divent_Click()
NUM1 = Val(Text1.Text)
OPERA = 5
Text1.Text = " "
End Sub

Private Sub dsgf_Click()
Form15.Hide
Form1.Show
End Sub

Private Sub Form_Load()
NUM1 = 0
NUM2 = 0
End Sub

Private Sub gfnjh_Click()
Form15.Hide
Form12.Show
End Sub

Private Sub ghjf_Click()
Form15.Hide
Form7.Show
End Sub

Private Sub htrtr_Click()
Form15.Hide
Form2.Show
End Sub

Private Sub igual_Click()
RESP = 0
NUM2 = Val(Text1.Text)
If OPERA = 1 Then
RESP = NUM1 + NUM2
End If
If OPERA = 2 Then
RESP = NUM1 - NUM2
End If
If OPERA = 3 Then
RESP = NUM1 * NUM2
End If
If OPERA = 4 Then
RESP = NUM1 / NUM2
End If
If OPERA = 5 Then
RESP = NUM1 \ NUM2
End If
If OPERA = 6 Then
RESP = NUM1 Mod NUM2
End If
Text1.Text = RESP
End Sub

Private Sub jd_Click()
End

End Sub

Private Sub jh_Click()
Form15.Hide
Form6.Show
End Sub

Private Sub mod_Click()
NUM1 = Val(Text1.Text)
OPERA = 6
Text1.Text = " "
End Sub

Private Sub multi_Click()
NUM1 = Val(Text1.Text)
OPERA = 3
Text1.Text = " "
End Sub

Private Sub N0_Click()
Text1.Text = Text1.Text + Str(0)
End Sub

Private Sub N1_Click()
Text1.Text = Text1.Text + Str(1)
End Sub

Private Sub N2_Click()
Text1.Text = Text1.Text + Str(2)
End Sub

Private Sub N3_Click()
Text1.Text = Text1.Text + Str(3)
End Sub

Private Sub N4_Click()
Text1.Text = Text1.Text + Str(4)
End Sub

Private Sub N5_Click()
Text1.Text = Text1.Text + Str(5)
End Sub

Private Sub N6_Click()
Text1.Text = Text1.Text + Str(6)
End Sub

Private Sub N7_Click()
Text1.Text = Text1.Text + Str(7)
End Sub

Private Sub N8_Click()
Text1.Text = Text1.Text + Str(8)
End Sub

Private Sub N9_Click()
Text1.Text = Text1.Text + Str(9)
End Sub

Private Sub reg_Click()
Form15.Hide
Form13.Show
End Sub

Private Sub resta_Click()
NUM1 = Val(Text1.Text)
OPERA = 2
Text1.Text = " "
End Sub

Private Sub rgr_Click()
Form15.Hide
Form9.Show
End Sub

Private Sub rgrrfg_Click()
Form15.Hide
Form11.Show
End Sub

Private Sub sdgf_Click()
Form15.Hide
Form3.Show
End Sub

Private Sub suma_Click()
NUM1 = Val(Text1.Text)
OPERA = 1
Text1.Text = " "
End Sub

Private Sub th_Click()
Form15.Hide
Form4.Show
End Sub

Private Sub thrt_Click()
Form15.Hide
Form8.Show
End Sub

Private Sub thst_Click()
Form15.Hide
Form5.Show
End Sub

Private Sub tht_Click()
Form15.Hide
Form14.Show
End Sub

Private Sub thth_Click()
Form14.Hide
Form10.Show
End Sub
