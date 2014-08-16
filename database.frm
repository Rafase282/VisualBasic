VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   Caption         =   "Base de Datos By R.282"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "End"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   4920
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "                          Datos"
      Connect         =   "Access"
      DatabaseName    =   "E:\R.282\Julio\R.282.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Raydatabase"
      Top             =   3480
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      DataField       =   "SN"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text8 
      DataField       =   "TD"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   16
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text7 
      DataField       =   "ISR"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text6 
      DataField       =   "SS"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   14
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text5 
      DataField       =   "SM"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   13
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "sueldo"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text3 
      DataField       =   "apellido"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      DataField       =   "nombre"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "codigo"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SN"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   225
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TD"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   225
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISR"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   270
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SS"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   210
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SM"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sueldo"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellidos"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   495
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
Data1.Recordset.Delete
Data1.Refresh
End Sub

Private Sub Command4_Click()
End

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
isrr = Val(Text7.Text)
sueldo = Val(Text4.Text)
sueldom = Val(Text5.Text)
seguros = Val(Text6.Text)
totald = sueldom + seguros + isrr
Text8.Text = Val(totald)
sueldon = sueldo - totald
Text9.Text = Val(sueldon)
End If
End Sub



