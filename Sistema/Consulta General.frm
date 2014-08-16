VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C00000&
   Caption         =   " Consulta General"
   ClientHeight    =   7620
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5310
   Icon            =   "Consulta General.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   7620
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Modificar"
      Height          =   255
      Left            =   3960
      TabIndex        =   98
      Top             =   7320
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3120
      TabIndex        =   95
      Top             =   1920
      Width           =   1815
      Begin VB.CommandButton Command4 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   1080
         TabIndex        =   97
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text63 
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
         TabIndex        =   96
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   2640
      TabIndex        =   94
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   255
      Left            =   1320
      TabIndex        =   93
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   255
      Left            =   120
      TabIndex        =   92
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "                        Consulta General"
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
      Top             =   6840
      Width           =   5055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "Introducion"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   2895
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
         Left            =   1200
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text2 
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
         Left            =   1200
         TabIndex        =   26
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Matricula"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Cuatrimestre"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   1560
      Picture         =   "Consulta General.frx":2372
      ScaleHeight     =   1275
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      Picture         =   "Consulta General.frx":58E6
      ScaleHeight     =   435
      ScaleWidth      =   5115
      TabIndex        =   1
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Notas"
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   4695
      Begin VB.TextBox Text62 
         DataField       =   "60"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   91
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text61 
         DataField       =   "59"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   90
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text60 
         DataField       =   "58"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   89
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text59 
         DataField       =   "57"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   88
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text58 
         DataField       =   "56"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   87
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text57 
         DataField       =   "55"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   86
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text56 
         DataField       =   "54"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   85
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text55 
         DataField       =   "53"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   84
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text54 
         DataField       =   "52"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   83
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text53 
         DataField       =   "51"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   82
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text52 
         DataField       =   "50"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   81
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text51 
         DataField       =   "49"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   80
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text50 
         DataField       =   "48"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   79
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text49 
         DataField       =   "47"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   78
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text48 
         DataField       =   "46"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   77
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text47 
         DataField       =   "45"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   76
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text46 
         DataField       =   "44"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   75
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text45 
         DataField       =   "43"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   74
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text44 
         DataField       =   "42"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   73
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text43 
         DataField       =   "41"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   72
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text42 
         DataField       =   "40"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   71
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text41 
         DataField       =   "39"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   70
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text40 
         DataField       =   "38"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   69
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text39 
         DataField       =   "37"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   68
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text38 
         DataField       =   "36"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   67
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text37 
         DataField       =   "35"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   66
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text36 
         DataField       =   "34"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   65
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text35 
         DataField       =   "33"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   64
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Text34 
         DataField       =   "32"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   63
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text33 
         DataField       =   "31"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   62
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text32 
         DataField       =   "30"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   61
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text31 
         DataField       =   "29"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   60
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text30 
         DataField       =   "28"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   59
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text29 
         DataField       =   "27"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   58
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text28 
         DataField       =   "26"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   57
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text27 
         DataField       =   "25"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   56
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text26 
         DataField       =   "24"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   55
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text25 
         DataField       =   "23"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   54
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text24 
         DataField       =   "22"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   53
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text23 
         DataField       =   "21"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   52
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text22 
         DataField       =   "20"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   51
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text21 
         DataField       =   "19"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   50
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text20 
         DataField       =   "18"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   49
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text19 
         DataField       =   "17"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   48
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text18 
         DataField       =   "16"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   47
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text17 
         DataField       =   "15"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   46
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text16 
         DataField       =   "14"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   45
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text15 
         DataField       =   "13"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   44
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text14 
         DataField       =   "12"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   43
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text13 
         DataField       =   "11"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   42
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text12 
         DataField       =   "10"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   41
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text11 
         DataField       =   "9"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   40
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text10 
         DataField       =   "8"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   39
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text9 
         DataField       =   "7"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   38
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text8 
         DataField       =   "6"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   37
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text7 
         DataField       =   "5"
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
         Height          =   285
         Left            =   3840
         TabIndex        =   36
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text6 
         DataField       =   "4"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   35
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text5 
         DataField       =   "3"
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
         Height          =   285
         Left            =   2640
         TabIndex        =   34
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text4 
         DataField       =   "2"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   33
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text3 
         DataField       =   "1"
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
         Height          =   285
         Left            =   1440
         TabIndex        =   32
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Mes"
         Height          =   255
         Left            =   3840
         TabIndex        =   31
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Dic"
         Height          =   255
         Left            =   3840
         TabIndex        =   30
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label30 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E.Sexual"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orientacion"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Informatica"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Frances"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "English"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E.Fisicas"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Nov"
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E. Artisticas"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Mes"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F.Humana"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Mes"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C.sociales"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Octubre"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Sept"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Agosto"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Mes"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C.Naturales"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Matematica"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lengua Espaola:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Mes"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "    Materias"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Menu menu 
      Caption         =   "&Menu"
      Begin VB.Menu inicu 
         Caption         =   "Inicio"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form4"
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
Data1.RecordSource = "select * from cafam  where Matricula = '" & Text63.Text & "'"
Data1.Refresh
If Text63.Text = "" Then
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

Private Sub inicu_Click()
Form4.Hide
Form1.Show
End Sub


