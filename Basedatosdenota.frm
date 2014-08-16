VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404080&
   Caption         =   "Calificacion Base By R.282"
   ClientHeight    =   4680
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      DataField       =   "Calificaciones"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "             Datos"
      Connect         =   "Access"
      DatabaseName    =   "E:\R.282\Julio\Nota.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Nota"
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      DataField       =   "Examen Final"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      DataField       =   "Examen Parcial"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      DataField       =   "Trabajo Practico"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      DataField       =   "Asistencia"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "Materia"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404080&
      Caption         =   "Calificacion Final"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404080&
      Caption         =   "Nombre"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404080&
      Caption         =   "Examen Final (30 Ptos.)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404080&
      Caption         =   "Examen Parcial (30 Ptos.)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      Caption         =   "Trabajo Practico (20 Ptos.)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404080&
      Caption         =   "Asistencia (20 Ptos)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Materia 
      BackColor       =   &H00404080&
      Caption         =   "Materia"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.Menu ydjd 
      Caption         =   "&File"
      Begin VB.Menu sdg 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, d, e As Integer


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


Private Sub sdg_Click()
End
End Sub

Private Sub Text6_Click()
a = Val(Text2.Text)
b = Val(Text3.Text)
c = Val(Text4.Text)
d = Val(Text5.Text)
e = (a + b + c + d)
Text6.Text = e
End Sub

