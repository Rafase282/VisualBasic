VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
