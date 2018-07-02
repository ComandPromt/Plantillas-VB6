VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtest 
   AutoRedraw      =   -1  'True
   Caption         =   "Test form"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000080&
   Icon            =   "frmtest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Show test form II"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
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
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":05C2
            Key             =   ""
            Object.Tag             =   "Sub&1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":071E
            Key             =   ""
            Object.Tag             =   "Sub&2"
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
            Object.Tag             =   "Sub&3"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7752
            Text            =   "Menu help text ( kMain )"
            TextSave        =   "Menu help text ( kMain )"
            Key             =   "kMain"
         EndProperty
      EndProperty
   End
   Begin VB.Label LblForm2 
      Alignment       =   2  'Center
      Caption         =   $"frmtest.frx":11FE
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "|Creates a new file|&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "|Open an existing file|&Open ..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "|Save the current file|&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "|Save the current file|Save &as..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "|Quit the application|&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "|Cut selected object|&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "|Copy selected object|&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "|Paste an object from the clipboard|&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Options"
      Begin VB.Menu mnuEmbossedColor 
         Caption         =   "#|Draws disabled images in color|&Embossed in color"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuComplexChecks 
         Caption         =   "#|Draws complex checks boxes and radio buttons|&Complex checks"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuFullSelect 
         Caption         =   "#|Draws a full selection bar|&Full selection"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-Color selection"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "*|Set the color to red|&Red"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuColor 
         Caption         =   "*|Set the color to green|&Green"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuColor 
         Caption         =   "*|Set the color to blue|&Blue"
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuLine31 
         Caption         =   "-Apply color..."
      End
      Begin VB.Menu mnuColorSel 
         Caption         =   "*|Will apply the next selected color to the menu font|... to menu caption"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuColorSel 
         Caption         =   "*|Will apply the next selected color to the menu selection|... to menu selection"
         Index           =   1
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-Sub Menu"
      End
      Begin VB.Menu mnuSub 
         Caption         =   "&Options"
         Begin VB.Menu mnuSub1 
            Caption         =   "|This is example ""Sub1""|Sub&1"
         End
         Begin VB.Menu mnuSub2 
            Caption         =   "|This is example ""Sub2""|Sub&2"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuLine2 
            Caption         =   "-More Checks"
         End
         Begin VB.Menu mnuSub3 
            Caption         =   "|This is example ""Sub3""|Sub&3"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSub4 
            Caption         =   "#|This is example ""Sub4""|Sub&4"
            Checked         =   -1  'True
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSub5 
            Caption         =   "*|This is example ""Sub5""|Sub&5"
            Checked         =   -1  'True
            Enabled         =   0   'False
         End
      End
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
''  frmTest Form
''
''  Copyright Olivier Martin 2000
''
''  martin.olivier@bigfoot.com
''
''  This form tests CoolMenu's functionality
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private WithEvents HelpObj As HelpCallBack
Attribute HelpObj.VB_VarHelpID = -1

Private Sub Command1_Click()
  frmTest2.Show
End Sub

Private Sub Form_Load()
  Set HelpObj = New HelpCallBack

  Call mCoolMenu.Install(Me.hWnd, HelpObj, ImageList, True, True)
  
'Any property function must be used AFTER
'installation

'If the FontName property is nothing,
'CoolMenu uses the form's text style and size
'If you set FontName to something, default size
'and color will be used.
'Setting size without FontName as no effect
'  Call mCoolMenu.FontName(Me.hWnd, "Tahoma")
'  Call mCoolMenu.FontSize(Me.hWnd, 8)
'  Call mCoolMenu.ForeColor(Me.hWnd, &H80)

'This is yet to be resolved: bright colors on
'selection bar should print text in dark color
'  Call mCoolMenu.SelectColor(Me.hWnd, vbWhite)

  mnuColor(0).Checked = True
  mnuColorSel(0).Checked = True
  
  mnuComplexChecks.Checked = mCoolMenu.ComplexChecks(Me.hWnd)
  mnuEmbossedColor.Checked = mCoolMenu.ColorEmbossed(Me.hWnd)
  mnuFullSelect.Checked = mCoolMenu.FullSelect(Me.hWnd)

  StatusBar.Panels("kMain").Text = ""
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Me.PopupMenu mnuEdit, 0, X, Y
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Call mCoolMenu.Uninstall(Me.hWnd)
  
  Set HelpObj = Nothing
End Sub

Private Sub HelpObj_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
  If Enabled Then
    StatusBar.Panels("kMain").Text = MenuHelp$
  Else
    StatusBar.Panels("kMain").Text = ""
  End If

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

Private Sub mnuEmbossedColor_Click()
  mnuEmbossedColor.Checked = Not mnuEmbossedColor.Checked
  Call mCoolMenu.ColorEmbossed(Me.hWnd, mnuEmbossedColor.Checked)
End Sub

Private Sub mnuFullSelect_Click()
  mnuFullSelect.Checked = Not mnuFullSelect.Checked
  Call mCoolMenu.FullSelect(Me.hWnd, mnuFullSelect.Checked)
End Sub

Private Sub mnuComplexChecks_Click()
  mnuComplexChecks.Checked = Not mnuComplexChecks.Checked
  Call mCoolMenu.ComplexChecks(Me.hWnd, mnuComplexChecks.Checked)
End Sub

Private Sub mnuQuit_Click()
  Unload Me
End Sub

Private Sub mnuColor_Click(Index As Integer)
  On Error Resume Next
  
  Dim i As Integer
  For i = 0 To 2
    mnuColor(i).Checked = (i = Index)
  Next i
  
  Dim Color As Long
  Color = CLng("&H80" + String(Index * 2, "0"))
  
  If mnuColorSel(0).Checked Then _
    Call mCoolMenu.ForeColor(Me.hWnd, Color&)
  
  If mnuColorSel(1).Checked Then _
    Call mCoolMenu.SelectColor(Me.hWnd, Color&)
    
End Sub

Private Sub mnuColorSel_Click(Index As Integer)
  Dim i As Integer
  For i = 0 To 1
    mnuColorSel(i).Checked = (i = Index)
  Next i
End Sub

Private Sub mnuSub3_Click()
  mnuSub3.Checked = Not mnuSub3.Checked
End Sub

Private Sub mnuSub4_Click()
  mnuSub4.Checked = Not mnuSub4.Checked
End Sub
