VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C00000&
   Caption         =   " Pago o Consulta Anual"
   ClientHeight    =   5640
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5415
   Icon            =   "Consulta general de Pago Anual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   ScaleHeight     =   5640
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   5160
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2280
      TabIndex        =   18
      Top             =   2520
      Width           =   1935
      Begin VB.CommandButton Command4 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Pago cgpa"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Forma de Pago Anual"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "                               Pago Anual"
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
      Top             =   4680
      Width           =   5055
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "Pagado"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C00000&
      Caption         =   "Forma de Pago"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   2880
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Tarjeta"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Efectivo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "Pago Anual"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   5175
      Begin VB.TextBox Text3 
         DataField       =   "Monto Unico"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         DataField       =   "FechaCGP"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Monto Unico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C00000&
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      Picture         =   "Consulta general de Pago Anual.frx":2372
      ScaleHeight     =   675
      ScaleWidth      =   5115
      TabIndex        =   3
      Top             =   1560
      Width           =   5175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   1560
      Picture         =   "Consulta general de Pago Anual.frx":2756
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Matricula"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
      Begin VB.TextBox Text1 
         DataField       =   "Matricula"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Menu menu 
      Caption         =   "&Menu"
      Begin VB.Menu ini 
         Caption         =   "Inicio"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Check1.Value = Checked
Text5.Text = "Pagado"
End Sub

Private Sub Command1_Click()
Data1.Recordset.AddNew
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Data1.Recordset.Update
Data1.Refresh
End Sub

Private Sub Command3_Click()
Data1.Recordset.Delete
Data1.Refresh
End Sub

Private Sub Command4_Click()
Data1.RecordSource = "select * from cafam  where Matricula = '" & Text6.Text & "'"
Data1.Refresh
If Text6.Text = "" Then
MsgBox "Esa MAtricula no Existe"
End If
End Sub

Private Sub Command5_Click()
Data1.Recordset.Edit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub ini_Click()
Form6.Hide
Form1.Show
End Sub

Private Sub Option1_Click()
Option1.Value = True
Text4.Text = "efectivo"
End Sub

Private Sub Option2_Click()
Option2.Value = True
Text4.Text = "tarjeta"
End Sub

Private Sub Text4_Change()
If Text4.Text = "efectivo" Then
Option1.Value = True
Else
Option2.Value = True
End If
End Sub

Private Sub Text5_Change()
If Text5.Text = "Pagado" Then
Check1.Value = Checked
End If
End Sub
