VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00C00000&
   Caption         =   " Pago o Consulta de Pago Trimestral"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5385
   ForeColor       =   &H00000000&
   Icon            =   "Consulta pago trimestral.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   ScaleHeight     =   8190
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4200
      TabIndex        =   43
      Top             =   7800
      Width           =   855
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C00000&
      Caption         =   "Consulta"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   40
      Top             =   2400
      Width           =   1935
      Begin VB.CommandButton Command4 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   1200
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text14 
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
         TabIndex        =   41
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Pagado trimestral3"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Forma de pago trimestral3"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Pagado trimestral2"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "forma de pago trimestral2"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Pagado trimestral"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Forma de pago trimestral"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   3000
      TabIndex        =   33
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1680
      TabIndex        =   32
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "                  Consulta de Pago Trimestral"
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
      Top             =   7320
      Width           =   5055
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C00000&
      Caption         =   "Tercer Trimestre"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   22
      Top             =   5880
      Width           =   5175
      Begin VB.Frame Frame7 
         BackColor       =   &H00C00000&
         Caption         =   "Forma de Pago"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   2880
         TabIndex        =   27
         Top             =   240
         Width           =   2175
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C00000&
            Caption         =   "Pagado"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C00000&
            Caption         =   "Efectivo"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1080
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C00000&
            Caption         =   "Tarjeta"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox Text7 
         DataField       =   "Monto trimesteral3"
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
         TabIndex        =   26
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         DataField       =   "fecha trimestral3"
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
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto a Pagar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C00000&
      Caption         =   "Segndo Trimestre"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   5175
      Begin VB.Frame Frame5 
         BackColor       =   &H00C00000&
         Caption         =   "Forma de pago"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   2880
         TabIndex        =   18
         Top             =   240
         Width           =   2175
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C00000&
            Caption         =   "Pagado"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C00000&
            Caption         =   "Efectivo"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C00000&
            Caption         =   "Tarjeta"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox Text5 
         DataField       =   "Monto a pagar trimestrasl2"
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
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         DataField       =   "Fecha Pago trimestral2"
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
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto a pagar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "PrimerTrimestre"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   5175
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Pagado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C00000&
         Caption         =   "Forma de Pago"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   2055
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C00000&
            Caption         =   "Efectivo"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C00000&
            Caption         =   "Tarjeta"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox Text3 
         DataField       =   "monto a pagar trimestreal"
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
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         DataField       =   "fecha pago trimestral"
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
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Monto a Pagar"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Matricula"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
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
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C00000&
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      Picture         =   "Consulta pago trimestral.frx":2372
      ScaleHeight     =   675
      ScaleWidth      =   5115
      TabIndex        =   1
      Top             =   1560
      Width           =   5175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   1680
      Picture         =   "Consulta pago trimestral.frx":2756
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu menu 
      Caption         =   "&Menu"
      Begin VB.Menu fes 
         Caption         =   "Inicio"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
Check1.Value = Checked
Text9.Text = "Pagado"
End Sub

Private Sub Check2_Click()
Check2.Value = Checked
Text11.Text = "Pagado"
End Sub

Private Sub Check3_Click()
Check3.Value = Checked
Text13.Text = "Pagado"
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
Data1.RecordSource = " select * from cafam where Matricula = '" & Text14.Text & "'"
Data1.Refresh
If Text14.Text = "" Then
MsgBox "Esa Matricula no Existe"
End If
End Sub

Private Sub Command5_Click()
Data1.Recordset.Edit
End Sub

Private Sub fes_Click()
Form7.Hide
Form1.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub
Private Sub Option1_Click()
Option1.Value = True
Text8.Text = "efectivo"
End Sub

Private Sub Option2_Click()
Option2.Value = True
Text8.Text = "tarjeta"
End Sub

Private Sub Option3_Click()
Option3.Value = True
Text10.Text = "efectivo"
End Sub

Private Sub Option5_Click()
Option5.Value = True
Text12.Text = "efectivo"
End Sub

Private Sub Option6_Click()
Option6.Value = False
Text12.Text = "tarjeta"
End Sub

Private Sub Text10_Change()
If Text10.Text = "efectivo" Then
Option3.Value = True
ElseIf Text10.Text = "tarjeta" Then
Option4.Value = True
End If
End Sub

Private Sub Text11_Change()
If Text11.Text = "Pagado" Then
Check2.Value = Checked
End If
End Sub

Private Sub Text12_Change()
If Text12.Text = "efectivo" Then
Option5.Value = True
ElseIf Text12.Text = "tarjeta" Then
Option6.Value = True
End If
End Sub

Private Sub Text13_Change()
If Text13.Text = "Pagado" Then
Check3.Value = Checked
End If
End Sub

Private Sub Text8_Change()
If Text8.Text = "efectivo" Then
Option1.Value = True
Else
Option2.Value = True
End If
End Sub

Private Sub Text9_Change()
If Text9.Text = "Pagado" Then
Check1.Value = Checked
End If
End Sub
