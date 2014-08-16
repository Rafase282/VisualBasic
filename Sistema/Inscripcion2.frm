VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C00000&
   Caption         =   " Inscripcion"
   ClientHeight    =   6375
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6135
   Icon            =   "Inscripcion2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   6375
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Modificar"
      Height          =   255
      Left            =   3600
      TabIndex        =   24
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Forma de Pago"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   4440
      TabIndex        =   23
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C00000&
      Caption         =   "Forma de Pago"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4560
      TabIndex        =   19
      Top             =   4200
      Width           =   1455
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C00000&
         Caption         =   "Mensual"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Trimestral"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Anual"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text6 
      DataField       =   "Matricula"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Religion"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Vive con"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nuevo"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   6000
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "                            Inscripcion Page 2"
      Connect         =   "Access"
      DatabaseName    =   "E:\R.282\Julio\Sistema\Sistema del Cafam By R.282.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cafam"
      Top             =   5520
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "Religion"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   4215
      Begin VB.OptionButton Option13 
         BackColor       =   &H00C00000&
         Caption         =   "Otros"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   34
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00C00000&
         Caption         =   "Budista"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C00000&
         Caption         =   "Adbentista"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   32
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C00000&
         Caption         =   "Testigo de Jeova"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C00000&
         Caption         =   "Protestante"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C00000&
         Caption         =   "Catolica"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text3 
      DataField       =   "Lugar que Ocupa"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      DataField       =   "No de Hermanos"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nombre del Tutor (a)"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3720
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Vive Con:"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C00000&
         Caption         =   "Tutor (a)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C00000&
         Caption         =   "Madre"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C00000&
         Caption         =   "Padre"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C00000&
         Caption         =   "Ambos Padres"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      Picture         =   "Inscripcion2.frx":2372
      ScaleHeight     =   555
      ScaleWidth      =   5115
      TabIndex        =   1
      Top             =   1440
      Width           =   5175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   2160
      Picture         =   "Inscripcion2.frx":2756
      ScaleHeight     =   1275
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Caption         =   "Matricula:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "Lugar que Ocupa"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "No. de Hermanos"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Nombre del Tutor (a)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "El Estudiante vive con:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Menu menu 
      Caption         =   "&Menu"
      Begin VB.Menu inicio 
         Caption         =   "Inicio"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Hide
Form2.Show
End Sub
Private Sub Command2_Click()
Data1.Recordset.AddNew
Text1.SetFocus
End Sub
Private Sub Command3_Click()
Data1.Recordset.Update
Data1.Refresh
End Sub
Private Sub Command4_Click()
Data1.Recordset.Delete
Data1.Refresh
End Sub
Private Sub Command5_Click()
Data1.Recordset.Edit
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub
Private Sub inicio_Click()
Form3.Hide
Form1.Show
End Sub
Private Sub Option1_Click()
Option1.Value = True
Text7.Text = "Anual"
End Sub
Private Sub Option10_Click()
If Option10.Value = True Then
Text5.Text = "Testigo de Jeova"
End If
End Sub
Private Sub Option11_Click()
If Option11.Value = True Then
Text5.Text = "Adbentista"
End If
End Sub
Private Sub Option12_Click()
If Option12.Value = True Then
Text5.Text = "Budista"
End If
End Sub
Private Sub Option13_Click()
If Option13.Value = True Then
Text5.Text = "Otros"
End If
End Sub
Private Sub Option2_Click()
Option2.Value = True
Text7.Text = "Trimestral"
End Sub
Private Sub Option3_Click()
Option3.Value = True
Text7.Text = "Mensual"
End Sub
Private Sub Option4_Click()
If Option4.Value = True Then
Text4.Text = "Ambos Padres"
End If
End Sub
Private Sub Option5_Click()
If Option5.Value = True Then
Text4.Text = "Padre"
End If
End Sub
Private Sub Option6_Click()
If Option6.Value = True Then
Text4.Text = "Madre"
End If
End Sub
Private Sub Option7_Click()
If Option7.Value = True Then
Text4.Text = "Tutor(a)"
End If
End Sub
Private Sub Option8_Click()
If Option8.Value = True Then
Text5.Text = "Catolica"
End If
End Sub
Private Sub Option9_Click()
If Option9.Value = True Then
Text5.Text = "Protestante"
End If
End Sub
Private Sub Text4_Change()
If Text4.Text = "Ambos Padres" Then
Option4.Value = True
ElseIf Text4.Text = "Padre" Then
Option5.Value = True
ElseIf Text4.Text = "Madre" Then
Option6.Value = True
ElseIf Text4.Text = "Tutor(a)" Then
Option7.Value = True
End If
End Sub
Private Sub Text5_Change()
If Text5.Text = "Catolica" Then
Option8.Value = True
ElseIf Text5.Text = "Protestante" Then
Option9.Value = True
ElseIf Text5.Text = "Testigo de Jeova" Then
Option10.Value = True
ElseIf Text5.Text = "Adbentista" Then
Option11.Value = True
ElseIf Text5.Text = "Budista" Then
Option12.Value = True
ElseIf Text5.Text = "Otros" Then
Option13.Value = True
End If
End Sub
Private Sub Text7_Change()
If Text7.Text = "Anual" Then
Option1.Value = True
Option2.Value = False
Option3.Value = False
ElseIf Text7.Text = "Trimestral" Then
Option2.Value = True
Option1.Value = False
Option3.Value = False
ElseIf Text7.Text = "Mensual" Then
Option3.Value = True
Option2.Value = False
Option1.Value = False
End If
End Sub
