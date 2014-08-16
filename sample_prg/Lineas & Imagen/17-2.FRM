VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   4650
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "顯  示  迴  旋  體"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim y(10) As Single, r(10) As Single
    y(0) = 200: y(1) = 180: y(2) = 160: y(3) = 140: y(4) = 120: y(5) = 100: y(6) = 80: y(7) = 60: y(8) = 40: y(9) = 20: y(10) = 0
    r(0) = 10: r(1) = 20: r(2) = 30: r(3) = 40: r(4) = 50: r(5) = 60: r(6) = 50: r(7) = 40: r(8) = 30: r(9) = 20: r(10) = 10
    rd = 3.1415927 / 180
    ax = 10 * rd
    ay = 0
    az = 10 * rd
    Picture1.Scale (-200, 300)-(200, -100)
    For k = 0 To 10
        For n = 0 To 360 Step 10
            x = r(k) * Cos(n * rd)
            z = r(k) * Sin(n * rd)
            X1 = x * Cos(ay) + z * Sin(ay)
            Y1 = y(k)
            Z1 = -x * Sin(ay) + z * Cos(ay)
            X2 = X1
            Y2 = Y1 * Cos(ax) - Z1 * Sin(ax)
            px = X2 * Cos(az) - Y2 * Sin(az)
            py = X2 * Sin(az) + Y2 * Cos(az)
            If n = 0 Then
                Picture1.PSet (px, py)
            Else
                Picture1.Line -(px, py)
            End If
        Next n
    Next k
    For n = 0 To 360 Step 60
        For k = 0 To 10
            x = r(k) * Cos(n * rd)
            z = r(k) * Sin(n * rd)
            X1 = x * Cos(ay) + z * Sin(ay)
            Y1 = y(k)
            Z1 = -x * Sin(ay) + z * Cos(ay)
            X2 = X1
            Y2 = Y1 * Cos(ax) - Z1 * Sin(ax)
            px = X2 * Cos(az) - Y2 * Sin(az)
            py = X2 * Sin(az) + Y2 * Cos(az)
            If k = 0 Then
                Picture1.PSet (px, py)
            Else
                Picture1.Line -(px, py)
            End If
        Next k
    Next n
End Sub
