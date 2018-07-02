VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtest 
   AutoRedraw      =   -1  'True
   Caption         =   "Test form"
   ClientHeight    =   2130
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmtest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":030A
            Key             =   ""
            Object.Tag             =   "&New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":0466
            Key             =   ""
            Object.Tag             =   "&Options"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":05C2
            Key             =   ""
            Object.Tag             =   "Sub1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":071E
            Key             =   ""
            Object.Tag             =   "Sub2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":087A
            Key             =   ""
            Object.Tag             =   "&Open"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":09D6
            Key             =   ""
            Object.Tag             =   "&Save"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":0B32
            Key             =   ""
            Object.Tag             =   "&Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":0C8E
            Key             =   ""
            Object.Tag             =   "&Cut"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":0DEA
            Key             =   ""
            Object.Tag             =   "&Paste"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":0F46
            Key             =   ""
            Object.Tag             =   "&Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":10A2
            Key             =   ""
            Object.Tag             =   "Sub3"
         EndProperty
      EndProperty
   End
   Begin VB.Label LblFont 
      Caption         =   "LblFont"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open ..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "Save &as..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Options"
      Begin VB.Menu mnuSub 
         Caption         =   "&Options"
         Begin VB.Menu mnuSub1 
            Caption         =   "Sub1"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuSub2 
            Caption         =   "Sub2"
            Enabled         =   0   'False
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuLine2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSub3 
            Caption         =   "Sub3"
            Shortcut        =   ^D
         End
      End
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  
  'declared in Public module
  Set CoolMenuObj = New CoolMenu
   
  Call CoolMenuObj.Install(Me.hwnd, ImageList, True, True)

  LblFont.Caption = Me.FontName
  LblFont.Font = Me.Font
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call CoolMenuObj.Install(0&)
  
  Set CoolMenuObj = Nothing
End Sub

Private Sub Label1_Click()

End Sub

Private Sub mnuCopy_Click()
  mnuPaste.Enabled = True
  mnuCopy.Enabled = False
  mnuCut.Enabled = False
End Sub

Private Sub mnuCut_Click()
  mnuPaste.Enabled = True
  mnuCopy.Enabled = False
  mnuCut.Enabled = False
End Sub

Private Sub mnuPaste_Click()
  mnuPaste.Enabled = False
  mnuCopy.Enabled = True
  mnuCut.Enabled = True

End Sub

Private Sub mnuQuit_Click()
  Unload Me
End Sub

Private Sub mnuSub3_Click()
  mnuSub3.Checked = Not mnuSub3.Checked
End Sub

Private Sub mnuSub4_Click()
  mnuSub4.Checked = Not mnuSub4.Checked

End Sub
