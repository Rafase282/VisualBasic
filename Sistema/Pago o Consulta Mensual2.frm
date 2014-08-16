VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00C00000&
   Caption         =   "  Pago o Consulta Mensual"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5460
   Icon            =   "Pago o Consulta Mensual2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   ScaleHeight     =   8190
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   2880
      TabIndex        =   45
      Top             =   7680
      Width           =   855
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C00000&
      Caption         =   "Busacar"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2040
      TabIndex        =   42
      Top             =   2520
      Width           =   1935
      Begin VB.CommandButton Command6 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   1200
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Forma de pago mensual4"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Pagado mensual 4"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Formadepagomensual5"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "Text10"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Pagadomensual5"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Forma de pago mensual 6"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      DataField       =   "Pagado mensual 6"
      DataSource      =   "Data1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Next"
      Height          =   375
      Left            =   4560
      TabIndex        =   35
      Top             =   7680
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
      Height          =   375
      Left            =   3840
      TabIndex        =   34
      Top             =   7680
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   2040
      TabIndex        =   33
      Top             =   7680
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1080
      TabIndex        =   32
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   7680
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "                   Pago o Consulta Mensual"
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
      Top             =   7200
      Width           =   5175
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C00000&
      Caption         =   "Diciembre"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   22
      Top             =   5880
      Width           =   5175
      Begin VB.TextBox Text6 
         DataField       =   "Fecha pagomensual6"
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
         TabIndex        =   28
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         DataField       =   "Monto apagar mensual6"
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
         TabIndex        =   27
         Top             =   720
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C00000&
         Caption         =   "Forma de Pago"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   2880
         TabIndex        =   23
         Top             =   240
         Width           =   2055
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C00000&
            Caption         =   "Pagado"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C00000&
            Caption         =   "Tarjeta"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C00000&
            Caption         =   "Efectivo"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   24
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Monto a Pagar"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "Noviembre"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   5175
      Begin VB.Frame Frame3 
         BackColor       =   &H00C00000&
         Caption         =   "Forma de Pago"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   2880
         TabIndex        =   16
         Top             =   240
         Width           =   2055
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C00000&
            Caption         =   "Efectivo"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C00000&
            Caption         =   "Tarjeta"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C00000&
            Caption         =   "Pagado"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.TextBox Text5 
         DataField       =   "Monto a pagar Mensual5"
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
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         DataField       =   "Fecha a Pagarmensual5"
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
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Monto a Pagar"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   20
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
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
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
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   1920
      Picture         =   "Pago o Consulta Mensual2.frx":2372
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   10
      Top             =   0
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C00000&
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      Picture         =   "Pago o Consulta Mensual2.frx":58E6
      ScaleHeight     =   675
      ScaleWidth      =   5115
      TabIndex        =   9
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C00000&
      Caption         =   "Octubre"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   5175
      Begin VB.TextBox Text2 
         DataField       =   "Fecha pago mensual 4"
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
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         DataField       =   "Moto a pagar mensual4"
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
         Top             =   720
         Width           =   1575
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C00000&
         Caption         =   "Forma de Pago"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Width           =   2055
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C00000&
            Caption         =   "Pagado"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C00000&
            Caption         =   "Tarjeta"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C00000&
            Caption         =   "Efectivo"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
         Caption         =   "Monto a Pagar"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "&Menu"
      Begin VB.Menu inicio 
         Caption         =   "Inicio"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Check2.Value = Checked
Text11.Text = "Pagado"
End Sub

Private Sub Check2_Click()
Check3.Value = Checked
Text13.Text = "Pagado"
End Sub

Private Sub Check5_Click()
Check1.Value = Checked
Text9.Text = "Pagado"
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
Form9.Hide
Form8.Show
End Sub

Private Sub Command5_Click()
Form9.Hide
form10.Show
End Sub

Private Sub Command6_Click()
Data1.RecordSource = "select * from cafam where Matricula = '" & Text14.Text & "'"
Data1.Refresh
If Text14.Text = "" Then
MsgBox "Esa Matricula no Existe"
End If
End Sub

Private Sub Command7_Click()
Data1.Recordset.Edit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub

Private Sub inicio_Click()
Form9.Hide
Form1.Show
End Sub

Private Sub Option1_Click()
 Option1.Value = True
Text8.Text = "Efectivo"
End Sub

Private Sub Option2_Click()
Option2.Value = True
Text10.Text = "Tarjeta"
End Sub

Private Sub Option3_Click()
Option3.Value = True
Text10.Text = "Efectivo"
End Sub

Private Sub Option4_Click()
Option4.Value = True
Text10.Text = "Tarjeta"
End Sub

Private Sub Option5_Click()
 Option5.Value = True
Text12.Text = "Efectivo"
End Sub

Private Sub Option6_Click()
Option6.Value = True
Text12.Text = "Tarjeta"
End Sub

Private Sub Text10_Change()
If Text10.Text = "Efectivo" Then
Option3.Value = True
Else
Option4.Value = True
End If
End Sub

Private Sub Text11_Change()
If Text11.Text = "Pagado" Then
Check2.Value = Checked
End If
End Sub

Private Sub Text12_Change()
If Text12.Text = "Efectivo" Then
Option5.Value = True
Else
Option6.Value = True
End If
End Sub

Private Sub Text13_Change()
If Text13.Text = "Pagado" Then
Check3.Value = Checked
End If
End Sub

Private Sub Text6_Change()
If Text8.Text = "Efectivo" Then
Option1.Value = True
Else
Option2.Value = True
End If
End Sub

Private Sub Text8_Change()
If Text8.Text = "Efectivo" Then
Option1.Value = True
Else
Option2.Value = True
End If
End Sub

Private Sub Text9_Change()
If Text12.Text = "Efectivo" Then
Option5.Value = True
Else
Option6.Value = True
End If
End Sub
