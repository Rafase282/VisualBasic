VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   8595
   StartUpPosition =   3  '系統預設值
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   2
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8070
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"18-7.frx":0000
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "18-7.frx":03C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "18-7.frx":04DB
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
        Case 1
            CommonDialog1.ShowOpen
            RichTextBox1.LoadFile CommonDialog1.FileName
        Case 2
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName
    End Select
End Sub
