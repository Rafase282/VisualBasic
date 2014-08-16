VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C00000&
   Caption         =   " Consulta Por Materia"
   ClientHeight    =   7755
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5235
   FillColor       =   &H00C00000&
   Icon            =   "Consulta por Materia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   ScaleHeight     =   7755
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Ver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   82
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Modificar"
      Height          =   255
      Left            =   3960
      TabIndex        =   81
      Top             =   7440
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C00000&
      Caption         =   "Consulta"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3360
      TabIndex        =   78
      Top             =   1800
      Width           =   1815
      Begin VB.CommandButton Command4 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   1080
         TabIndex        =   80
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text52 
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
         TabIndex        =   79
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   2760
      TabIndex        =   77
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   255
      Left            =   1440
      TabIndex        =   76
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   255
      Left            =   120
      TabIndex        =   75
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "                         Consulta por Materia"
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
      Top             =   6960
      Width           =   5055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "Notas"
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Width           =   4215
      Begin VB.TextBox Text4 
         DataField       =   "01"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   57
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text5 
         DataField       =   "02"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   56
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text6 
         DataField       =   "03"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   55
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text7 
         DataField       =   "04"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   54
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text8 
         DataField       =   "05"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   53
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text9 
         DataField       =   "06"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   52
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text10 
         DataField       =   "07"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   51
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text11 
         DataField       =   "08"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   50
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text12 
         DataField       =   "09"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   49
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text13 
         DataField       =   "010"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   48
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text14 
         DataField       =   "011"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   47
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text15 
         DataField       =   "012"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   46
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H80000004&
         DataField       =   "013"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   45
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00FFFFFF&
         DataField       =   "014"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   44
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H00FFFFFF&
         DataField       =   "015"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   43
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text19 
         DataField       =   "016"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   42
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text20 
         DataField       =   "017"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   41
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text21 
         DataField       =   "018"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   40
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text22 
         DataField       =   "019"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   39
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text23 
         DataField       =   "020"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   38
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text24 
         DataField       =   "021"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   37
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text25 
         DataField       =   "022"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   36
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text26 
         DataField       =   "023"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   35
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text27 
         DataField       =   "024"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   34
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text28 
         DataField       =   "025"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   33
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text29 
         DataField       =   "026"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   32
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text30 
         DataField       =   "027"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   31
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text31 
         DataField       =   "028"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   30
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text32 
         DataField       =   "029"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text33 
         DataField       =   "030"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text34 
         DataField       =   "031"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   27
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text35 
         DataField       =   "032"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   26
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text36 
         DataField       =   "033"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text37 
         DataField       =   "034"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text38 
         DataField       =   "035"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   23
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text39 
         DataField       =   "036"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   22
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text40 
         DataField       =   "037"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text41 
         DataField       =   "038"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   20
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text42 
         DataField       =   "039"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   19
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text43 
         DataField       =   "040"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   18
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text44 
         DataField       =   "041"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text45 
         DataField       =   "042"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   16
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text46 
         DataField       =   "043"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text47 
         DataField       =   "044"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   14
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text48 
         DataField       =   "045"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text49 
         DataField       =   "046"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text50 
         DataField       =   "047"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         TabIndex        =   11
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text51 
         DataField       =   "048"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   10
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Materia"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota1"
         Height          =   255
         Left            =   1560
         TabIndex        =   73
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota2"
         Height          =   255
         Left            =   2160
         TabIndex        =   72
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota3"
         Height          =   255
         Left            =   2760
         TabIndex        =   71
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota4"
         Height          =   255
         Left            =   3360
         TabIndex        =   70
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lengua Española"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Matematica"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C.Naturales"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C.Sociales"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F.Humana"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E.Artisticas"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E. Fisica"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "English"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Frances"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Informatica"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orientacion"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E. Sexual"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   3120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Materia"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
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
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         DataField       =   "Materia"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         DataField       =   "Cuatrimestre"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Matricula"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Materia"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuatrimestre"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   1560
      Picture         =   "Consulta por Materia.frx":2372
      ScaleHeight     =   1155
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C00000&
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   0
      Picture         =   "Consulta por Materia.frx":58E6
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Menu menu 
      Caption         =   "&Menu"
      Begin VB.Menu rg 
         Caption         =   "Inicio"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form5"
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
Data1.RecordSource = "select * from cafam  where Matricula = '" & Text52.Text & "'"
Data1.Refresh
If Text52.Text = "" Then
MsgBox "Esa MAtricula no Existe"
End If
End Sub
Private Sub Command5_Click()
Data1.Recordset.Edit
End Sub

Private Sub Command6_Click()
If Text2.Text = "" Or Text2.Text = " " Then
MsgBox "No a Digitado la Materia"
Text2.SetFocus
ElseIf Text2.Text = "Lengua Española" Then
Text4.ForeColor = &H80000012
Text5.ForeColor = &H80000012
Text6.ForeColor = &H80000012
Text7.ForeColor = &H80000012
ElseIf Text2.Text = "Matematica" Then
Text8.ForeColor = &H80000012
Text9.ForeColor = &H80000012
Text10.ForeColor = &H80000012
Text11.ForeColor = &H80000012
ElseIf Text2.Text = "C.Naturales" Then
Text12.ForeColor = &H80000012
Text13.ForeColor = &H80000012
Text14.ForeColor = &H80000012
Text15.ForeColor = &H80000012
ElseIf Text2.Text = "C.Sociales" Then
Text16.ForeColor = &H80000012
Text17.ForeColor = &H80000012
Text18.ForeColor = &H80000012
Text19.ForeColor = &H80000012
ElseIf Text2.Text = "F.Humana" Then
Text20.ForeColor = &H80000012
Text21.ForeColor = &H80000012
Text22.ForeColor = &H80000012
Text23.ForeColor = &H80000012
ElseIf Text2.Text = "E.Artisticas" Then
Text24.ForeColor = &H80000012
Text25.ForeColor = &H80000012
Text26.ForeColor = &H80000012
Text27.ForeColor = &H80000012
ElseIf Text2.Text = "E. Fisica" Then
Text28.ForeColor = &H80000012
Text29.ForeColor = &H80000012
Text30.ForeColor = &H80000012
Text31.ForeColor = &H80000012
ElseIf Text2.Text = "English" Then
Text32.ForeColor = &H80000012
Text33.ForeColor = &H80000012
Text34.ForeColor = &H80000012
Text35.ForeColor = &H80000012
ElseIf Text2.Text = "Frances" Then
Text36.ForeColor = &H80000012
Text37.ForeColor = &H80000012
Text38.ForeColor = &H80000012
Text39.ForeColor = &H80000012
ElseIf Text2.Text = "Informatica" Then
Text40.ForeColor = &H80000012
Text41.ForeColor = &H80000012
Text42.ForeColor = &H80000012
Text43.ForeColor = &H80000012
ElseIf Text2.Text = "Orientacion" Then
Text44.ForeColor = &H80000012
Text45.ForeColor = &H80000012
Text46.ForeColor = &H80000012
Text47.ForeColor = &H80000012
ElseIf Text2.Text = "E. Sexual" Then
Text48.ForeColor = &H80000012
Text49.ForeColor = &H80000012
Text50.ForeColor = &H80000012
Text51.ForeColor = &H80000012
Else: MsgBox "Materia Incorrecta"
Text2.Text = ""
Text2.SetFocus
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub
Private Sub rg_Click()
Form5.Hide
Form1.Show
End Sub
