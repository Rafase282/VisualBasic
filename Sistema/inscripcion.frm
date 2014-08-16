VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C00000&
   Caption         =   " Inscripcion"
   ClientHeight    =   7050
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6975
   Icon            =   "inscripcion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Modificar"
      Height          =   255
      Left            =   4200
      TabIndex        =   32
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox Text11 
      DataField       =   "Celular"
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
      Height          =   360
      Left            =   5400
      TabIndex        =   31
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      DataField       =   "Beeper"
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
      Height          =   360
      Left            =   3480
      TabIndex        =   30
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      DataField       =   "Telefono"
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
      Height          =   360
      Left            =   960
      TabIndex        =   29
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      DataField       =   "Direcion2"
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
      Left            =   120
      TabIndex        =   28
      Top             =   5160
      Width           =   6735
   End
   Begin VB.TextBox Text7 
      DataField       =   "Direcion Actual"
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
      Left            =   1320
      TabIndex        =   27
      Top             =   4680
      Width           =   5535
   End
   Begin VB.TextBox Text6 
      DataField       =   "Fecha de Nacimiento"
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
      Left            =   5520
      TabIndex        =   26
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      DataField       =   "Lugar de Nacimiento"
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
      Left            =   1680
      TabIndex        =   25
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "Nacionalidad"
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
      Left            =   5280
      TabIndex        =   24
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      DataField       =   "Edad"
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
      Left            =   720
      TabIndex        =   23
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text2 
      DataField       =   "Apellidos"
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
      Left            =   3960
      TabIndex        =   22
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nombre"
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
      Left            =   720
      TabIndex        =   21
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "                                                   Inscripcion"
      Connect         =   "Access"
      DatabaseName    =   "E:\R.282\Julio\Sistema\Sistema del Cafam By R.282.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cafam"
      Top             =   6120
      Width           =   6615
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Sexo"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nuevo"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   6600
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C00000&
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   1080
      Picture         =   "inscripcion.frx":2372
      ScaleHeight     =   675
      ScaleWidth      =   5115
      TabIndex        =   15
      Top             =   1560
      Width           =   5175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   2640
      Picture         =   "inscripcion.frx":2756
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   14
      Top             =   0
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Sexo"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   3360
      Width           =   2535
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Femenino"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Masculino"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C00000&
      Caption         =   "Celular:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Bee 
      BackColor       =   &H00C00000&
      Caption         =   "Beeper:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C00000&
      Caption         =   "Telefono:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C00000&
      Caption         =   "Direccion Actual"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C00000&
      Caption         =   "Fecha de Nacimiento:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "Lugar de Nacimiento"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Caption         =   "Nacionalidad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "Edad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "Apellidos:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Datos Personales del Estudiante:"
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
      Left            =   720
      TabIndex        =   0
      Top             =   2400
      Width           =   5415
   End
   Begin VB.Menu menu 
      Caption         =   "&Menu"
      Begin VB.Menu inicio 
         Caption         =   "Inicio"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Form3.Show
End Sub

Private Sub Command2_Click()
Data1.Recordset.AddNew
Text12.Text = ""
Option1.Value = False
Option2.Value = False
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
Form2.Hide
Form1.Show
End Sub

Private Sub Option1_Click()
Option1.Value = True
Text12.Text = "Masculino"
End Sub

Private Sub Option2_Click()
Option2.Value = True
Text12.Text = "Femenino"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Text1.Text = "" And KeyAscii = 13 Then
MsgBox "Nombre Requerido"
End If
End Sub

Private Sub Text12_Change()
If Text12.Text = "Masculino" Then
Option1.Value = True
Else
Option2.Value = True
End If
End Sub
