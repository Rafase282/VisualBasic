VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H8000000E&
   Caption         =   "Anuncio del CAFAM By R.282"
   ClientHeight    =   3360
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4095
   Icon            =   "Form7.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":628A
   ScaleHeight     =   3360
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   2400
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   3120
      Width           =   2625
   End
   Begin VB.Menu sdfg 
      Caption         =   "&File"
      Begin VB.Menu erg 
         Caption         =   "&Forum"
         Shortcut        =   {F1}
      End
      Begin VB.Menu hgj 
         Caption         =   "-"
      End
      Begin VB.Menu hfa 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu gfjks 
      Caption         =   "&Programs"
      Begin VB.Menu rg 
         Caption         =   "&Tablas"
         Begin VB.Menu reyeqr 
            Caption         =   "Tabla del 1"
         End
         Begin VB.Menu pufyp 
            Caption         =   "Tabla de cualquier numero"
         End
      End
      Begin VB.Menu fh 
         Caption         =   "-"
      End
      Begin VB.Menu color 
         Caption         =   "Colores"
      End
      Begin VB.Menu etgfwqe 
         Caption         =   "Promedio"
      End
      Begin VB.Menu calcio 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu ewtwer 
         Caption         =   "Calificaciones"
      End
      Begin VB.Menu tuwtu 
         Caption         =   "Area del Triangulo"
      End
      Begin VB.Menu eqwer 
         Caption         =   "Adivina el Numero"
      End
      Begin VB.Menu ew 
         Caption         =   "Area del Cuadrado"
      End
      Begin VB.Menu dsgd 
         Caption         =   "Formula Cuadratica"
      End
      Begin VB.Menu orn 
         Caption         =   "&Ordene los Numeros"
      End
      Begin VB.Menu ewf 
         Caption         =   "Entrada de Estudiantes"
      End
      Begin VB.Menu fhg 
         Caption         =   "Formula C = t+h*d-m /4 "
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub calcio_Click()
Form7.Hide
Form15.Show
End Sub

Private Sub color_Click()
Form7.Hide
Form13.Show
End Sub

Private Sub dsgd_Click()
Form7.Hide
Form14.Show
End Sub

Private Sub eqwer_Click()
Form7.Hide
Form4.Show
End Sub

Private Sub erg_Click()
Form7.Hide
Form1.Show
End Sub

Private Sub etgfwqe_Click()
Form7.Hide
Form5.Show
End Sub

Private Sub ew_Click()
Form7.Hide
Form10.Show
End Sub

Private Sub ewf_Click()
Form7.Hide
Form9.Show
End Sub

Private Sub ewtwer_Click()
Form7.Hide
Form8.Show
End Sub

Private Sub fhg_Click()
Form7.Hide
Form12.Show
End Sub

Private Sub hfa_Click()
End
End Sub

Private Sub orn_Click()
Form7.Hide
Form11.Show
End Sub

Private Sub pufyp_Click()
from7.Hide
Form6.Show
End Sub

Private Sub reyeqr_Click()
Form7.Hide
Form3.Show
End Sub
Private Sub Timer1_Timer()
If Label1.Caption = "" Then
Label1.Caption = "E"
ElseIf Label1.Caption = "E" Then
Label1.Caption = "El"
ElseIf Label1.Caption = "El" Then
Label1.Caption = "El "
ElseIf Label1.Caption = "El " Then
Label1.Caption = "El C"
ElseIf Label1.Caption = "El C" Then
Label1.Caption = "El Co"
ElseIf Label1.Caption = "El Co" Then
Label1.Caption = "El Col"
ElseIf Label1.Caption = "El Col" Then
Label1.Caption = "El Cole"
ElseIf Label1.Caption = "El Cole" Then
Label1.Caption = "El Coleg"
ElseIf Label1.Caption = "El Coleg" Then
Label1.Caption = "El Colegi"
ElseIf Label1.Caption = "El Colegi" Then
Label1.Caption = "El Colegio"
ElseIf Label1.Caption = "El Colegio" Then
Label1.Caption = "El Colegio "
ElseIf Label1.Caption = "El Colegio " Then
Label1.Caption = "El Colegio C"
ElseIf Label1.Caption = "El Colegio C" Then
Label1.Caption = "El Colegio CA"
ElseIf Label1.Caption = "El Colegio CA" Then
Label1.Caption = "El Colegio CAF"
ElseIf Label1.Caption = "El Colegio CAF" Then
Label1.Caption = "El Colegio CAFA"
ElseIf Label1.Caption = "El Colegio CAFA" Then
Label1.Caption = "El Colegio CAFAM"
ElseIf Label1.Caption = "El Colegio CAFAM" Then
Label1.Caption = "El Colegio CAFAM "
ElseIf Label1.Caption = "El Colegio CAFAM " Then
Label1.Caption = "El Colegio CAFAM e"
ElseIf Label1.Caption = "El Colegio CAFAM e" Then
Label1.Caption = "El Colegio CAFAM es"
ElseIf Label1.Caption = "El Colegio CAFAM es" Then
Label1.Caption = "El Colegio CAFAM est"
ElseIf Label1.Caption = "El Colegio CAFAM est" Then
Label1.Caption = "El Colegio CAFAM está"
ElseIf Label1.Caption = "El Colegio CAFAM está" Then
Label1.Caption = "El Colegio CAFAM está "
ElseIf Label1.Caption = "El Colegio CAFAM está " Then
Label1.Caption = "El Colegio CAFAM está e"
ElseIf Label1.Caption = "El Colegio CAFAM está e" Then
Label1.Caption = "El Colegio CAFAM está en"
ElseIf Label1.Caption = "El Colegio CAFAM está en" Then
Label1.Caption = "El Colegio CAFAM está en "
ElseIf Label1.Caption = "El Colegio CAFAM está en " Then
Label1.Caption = "El Colegio CAFAM está en p"
ElseIf Label1.Caption = "El Colegio CAFAM está en p" Then
Label1.Caption = "El Colegio CAFAM está en pr"
ElseIf Label1.Caption = "El Colegio CAFAM está en pr" Then
Label1.Caption = "El Colegio CAFAM está en pro"
ElseIf Label1.Caption = "El Colegio CAFAM está en pro" Then
Label1.Caption = "El Colegio CAFAM está en prog"
ElseIf Label1.Caption = "El Colegio CAFAM está en prog" Then
Label1.Caption = "El Colegio CAFAM está en progr"
ElseIf Label1.Caption = "El Colegio CAFAM está en progr" Then
Label1.Caption = "El Colegio CAFAM está en progre"
ElseIf Label1.Caption = "El Colegio CAFAM está en progre" Then
Label1.Caption = "El Colegio CAFAM está en progres"
ElseIf Label1.Caption = "El Colegio CAFAM está en progres" Then
Label1.Caption = "El Colegio CAFAM está en progreso"
ElseIf Label1.Caption = "El Colegio CAFAM está en progreso" Then
Label1.Caption = ""
ElseIf Label.Caption = "" Then
Form2.Show
End If
End Sub

Private Sub tuwtu_Click()
Form7.Hide
Form2.Show
End Sub
