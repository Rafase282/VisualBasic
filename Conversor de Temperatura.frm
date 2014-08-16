VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   Caption         =   "Conversor de Temperatura By R.282"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2895
      LargeChange     =   10
      Left            =   2160
      Max             =   -100
      Min             =   100
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "0"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Grados Fahrenheit"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label G 
      BackColor       =   &H000000FF&
      Caption         =   "Grados Centrigrados"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu exit 
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
Option Explicit
Private Sub exit_Click()
Beep
End
End Sub


Private Sub VScroll1_Change()
Text1.Text = VScroll1.Value
Text2.Text = 32 + 1.8 * VScroll1.Value
End Sub
