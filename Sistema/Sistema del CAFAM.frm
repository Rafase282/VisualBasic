VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "  Sistema del CAFAM By R.282"
   ClientHeight    =   3150
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4470
   Icon            =   "Sistema del CAFAM.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Sistema del CAFAM.frx":2372
   ScaleHeight     =   3150
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   0
      Picture         =   "Sistema del CAFAM.frx":349B
      ScaleHeight     =   3075
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Menu File 
      Caption         =   "&Archivo"
      Begin VB.Menu Salida 
         Caption         =   "&Salida"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu sistema 
      Caption         =   "&Sistema"
      Begin VB.Menu Inscripcion 
         Caption         =   "Inscripción"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Nota 
         Caption         =   "Nota"
         Begin VB.Menu nxmatria 
            Caption         =   "Por Materias"
            Shortcut        =   {F3}
         End
         Begin VB.Menu congeneral 
            Caption         =   "Consl. General"
            Shortcut        =   {F8}
         End
      End
      Begin VB.Menu Pago 
         Caption         =   "Pago"
         Begin VB.Menu anual 
            Caption         =   "Anual"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mensual 
            Caption         =   "Mensual"
            Shortcut        =   {F6}
         End
         Begin VB.Menu trimestral 
            Caption         =   "Trimestral"
            Shortcut        =   {F7}
         End
         Begin VB.Menu balance 
            Caption         =   "Consulta de Balance"
            Begin VB.Menu panual 
               Caption         =   "C. P. Anual"
               Shortcut        =   {F9}
            End
            Begin VB.Menu pmensual 
               Caption         =   "C. P. Mensual"
               Shortcut        =   {F11}
            End
            Begin VB.Menu ptrimestral 
               Caption         =   "C. P. Trimestral"
               Shortcut        =   {F12}
            End
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub anual_Click()
Form1.Hide
Form6.Show
End Sub

Private Sub congeneral_Click()
Form1.Hide
Form4.Show
End Sub

Private Sub Inscripcion_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub mensual_Click()
Form1.Hide
Form8.Show
End Sub

Private Sub nxmatria_Click()
Form1.Hide
Form5.Show
End Sub

Private Sub panual_Click()
Form1.Hide
Form6.Show
End Sub

Private Sub pmensual_Click()
Form1.Hide
Form8.Show
End Sub

Private Sub ptrimestral_Click()
Form1.Hide
Form7.Show
End Sub

Private Sub Salida_Click()
End
End Sub

Private Sub trimestral_Click()
Form1.Hide
Form7.Show
End Sub
