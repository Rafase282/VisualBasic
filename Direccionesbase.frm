VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Direciones By R.282"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text5 
      DataField       =   "Celular"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      DataField       =   "Telefono"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      DataField       =   "Apellidos"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "Matricula"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   615
   End
   Begin VB.Data Data1 
      Caption         =   "             Consulta"
      Connect         =   "Access"
      DatabaseName    =   "E:\R.282\Julio\Direcciones.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Agenda"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Celular"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telefono"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Apellidos"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Matricula"
      ForeColor       =   &H80000004&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Data1.Recordset.Update
Data1.Refresh
End Sub

Private Sub Command3_Click()
Data1.Recordset.Edit
End Sub

Private Sub Command4_Click()
Data1.Recordset.Delete
Data1.Refresh
End Sub

Private Sub Command5_Click()
Data1.RecordSource = "select * from Agenda where Matricula = '" & Text6.Text & "'"
Data1.Refresh
If Text6.Text = " " Then
MsgBox "Esa Matricula no Existe"
End If
End Sub
