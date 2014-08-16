VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   3990
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text2 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Menu edit 
      Caption         =   "編    輯"
      Begin VB.Menu cut 
         Caption         =   "剪    下"
      End
      Begin VB.Menu copy 
         Caption         =   "複    製"
      End
      Begin VB.Menu paste 
         Caption         =   "貼    上"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub copy_Click()
    Clipboard.SetText ActiveControl.SelText
End Sub

Private Sub cut_Click()
    Clipboard.SetText ActiveControl.SelText
    ActiveControl.SelText = ""
End Sub

Private Sub paste_Click()
    ActiveControl.SelText = Clipboard.GetText()
End Sub
