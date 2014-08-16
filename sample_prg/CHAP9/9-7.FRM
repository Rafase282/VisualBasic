VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   2295
   ClientTop       =   2250
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   4335
   ScaleWidth      =   8805
   Begin ComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7435
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   4020
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    ChDir Dir1.Path
    ListView1.ListItems.Clear
    p = Dir("*.*", 2 + 4 + 16)
    Do While p <> ""
        If Not (p = "." Or p = "..") Then
            If (GetAttr(p) And vbDirectory) = vbDirectory Then
                Set lv = ListView1.ListItems.Add(, , p)
                lv.SubItems(1) = "資料夾"
            Else
                Set lv = ListView1.ListItems.Add(, , p)
                lv.SubItems(1) = (FileLen(p) + 1023) \ 1024 & "KB"
            End If
            lv.SubItems(2) = FileDateTime(p)
        End If
        p = Dir
    Loop
End Sub

Private Sub Drive1_Change()
    ChDrive Drive1.Drive
    Dir1.Path = Drive1.Drive
    Dir1_Change
End Sub

Private Sub Form_Load()
    Dim x As ColumnHeader
    Set x = ListView1.ColumnHeaders.Add(, , "檔案名稱", ListView1.Width / 3)
    Set x = ListView1.ColumnHeaders.Add(, , "SIZE", ListView1.Width / 5, 1)
    Set x = ListView1.ColumnHeaders.Add(, , "更新日期", ListView1.Width / 4, 1)
    ListView1.View = lvwReport
    Dir1_Change
End Sub

