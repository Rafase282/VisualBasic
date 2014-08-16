VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00000000&
   Caption         =   "Colores By R.282"
   ClientHeight    =   3675
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form13.frx":628A
   ScaleHeight     =   3675
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   " Colores "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.OptionButton Colores 
         BackColor       =   &H00000000&
         Caption         =   "Verde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Colores 
         BackColor       =   &H00000000&
         Caption         =   "Amarillo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Colores 
         BackColor       =   &H00000000&
         Caption         =   "Rojo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Colores 
         BackColor       =   &H00000000&
         Caption         =   "Azul"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Colores 
         BackColor       =   &H00000000&
         Caption         =   "Negro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   1935
      Left            =   2400
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu foro 
         Caption         =   "&Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu t 
         Caption         =   "-"
      End
      Begin VB.Menu ExIt 
         Caption         =   "&ExIt"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu prog 
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
      Begin VB.Menu tj 
         Caption         =   "-"
      End
      Begin VB.Menu pro 
         Caption         =   "Promedio"
      End
      Begin VB.Menu cal 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu cali 
         Caption         =   "Calificaciones"
      End
      Begin VB.Menu tr 
         Caption         =   "Area del Triangulo"
      End
      Begin VB.Menu ca 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu adi 
         Caption         =   "Adivina de Numero"
      End
      Begin VB.Menu for 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu ad 
         Caption         =   "Anuncio del CAFAM"
      End
      Begin VB.Menu orden 
         Caption         =   "Ordene los Numeros"
      End
      Begin VB.Menu da 
         Caption         =   "Entrada de Estudiantes"
      End
      Begin VB.Menu fo 
         Caption         =   "Formula C=t+h*d-m/4 "
      End
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ad_Click()
Form13.Hide
Form7.Show
End Sub

Private Sub adi_Click()
Form13.Hide
Form4.Show
End Sub

Private Sub ca_Click()
Form13.Hide
Form10.Show
End Sub

Private Sub cal_Click()
Form13.Hide
Form15.Show
End Sub

Private Sub cali_Click()
Form13.Hide
Form8.Show
End Sub

Private Sub Colores_Click(index As Integer)
Select Case index
Case 0
Shape2.BackColor = RGB(0, 0, 0) 'negro
Shape2.BackColor = RGB(0, 0, 0) 'negro
Case 1
Shape2.BackColor = RGB(0, 0, 255) 'azul
Shape2.BackColor = RGB(0, 0, 255) 'azul
Case 2
Shape2.BackColor = RGB(255, 0, 0) ' Rojo
Shape2.BackColor = RGB(255, 0, 0) ' Rojo
Case 3
Shape2.BackColor = RGB(255, 255, 0) 'Amarillo
Shape2.BackColor = RGB(255, 255, 0) 'Amarillo
Case 4
Shape2.BackColor = RGB(0, 255, 0) 'Verde
Shape2.BackColor = RGB(0, 255, 0) 'Verde
End Select
End Sub

Private Sub da_Click()
Form13.Hide
Form9.Show
End Sub

Private Sub Exit_Click()
End
End Sub


Private Sub fo_Click()
Form13.Hide
Form12.Show
End Sub

Private Sub for_Click()
Form13.Hide
Form14.Show
End Sub

Private Sub foro_Click()
Form13.Hide
Form1.Show
End Sub

Private Sub orden_Click()
Form13.Hide
Form11.Show
End Sub

Private Sub pro_Click()
Form13.Hide
Form5.Show
End Sub

Private Sub tabla1_Click()
Form13.Hide
Form3.Show
End Sub

Private Sub tabla2_Click()
Form13.Hide
Form6.Show
End Sub

Private Sub tr_Click()
Form13.Hide
Form2.Show
End Sub
