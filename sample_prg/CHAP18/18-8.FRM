VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   4035
   ClientTop       =   2160
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   3765
   ScaleWidth      =   3150
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "18-8.frx":0000
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Text1.Text = "人生怨嘆無路用,誰人沒家庭."
End Sub

Private Sub Text1_DblClick()
    Form2.Show 1
End Sub
