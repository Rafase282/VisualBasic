VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "我的 CD 播放程式"
   ClientHeight    =   2130
   ClientLeft      =   7005
   ClientTop       =   2235
   ClientWidth     =   4725
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   2130
   ScaleWidth      =   4725
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   540
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2175
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    Text1.Text = ""
    MMControl1.UpdateInterval = 1000

    MMControl1.DeviceType = "CDAudio"
    MMControl1.Command = "Open"
    MMControl1.TimeFormat = mciFormatTmsf
    For i = 1 To MMControl1.Tracks
        List1.AddItem "Track" & Right("0" & i, 2)
    Next i
End Sub

Private Sub List1_Click()
    t = Val(Right(List1.List(List1.ListIndex), 2))
    MMControl1.From = t
    MMControl1.Track = t
    MMControl1.To = MMControl1.TrackLength * 256 + t
    MMControl1.Command = "Play"
End Sub

Private Sub MMControl1_StatusUpdate()
    f = Right("00000000" & Hex(MMControl1.Position), 8)
    Text1.Text = "[" & Val("&h" & Right(f, 2)) & "]:" & Val("&h" & Mid(f, 5, 2)) & ":" & Val("&h" & Mid(f, 3, 2))
End Sub



