VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   2415
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox List1 
      Height          =   2580
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetArray(a() As String, dat As String)
    m = 1
    For i = LBound(a, 1) To UBound(a, 1)
        n = InStr(m, dat, ",")
        a(i) = Mid(dat, m, n - m)
        m = n + 1
    Next i
End Sub

Private Sub Form_Load()
    Dim girl(4) As String, dat1 As String
    Dim boy(3) As String, dat2 As String
    dat1 = "Amy,Rolla,Nikey,ada,Nancy,"
    dat2 = "Bear,Joe,Milk,Tom,"
    SetArray girl(), dat1
    SetArray boy(), dat2
    
    For i = 0 To 4
        List1.AddItem girl(i)
    Next i
    For i = 0 To 3
        List1.AddItem boy(i)
    Next i
End Sub
