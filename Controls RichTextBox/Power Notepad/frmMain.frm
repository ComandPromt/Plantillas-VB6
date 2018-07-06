VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Power Notepad Example by nk - nkillaz.com [New File]"
   ClientHeight    =   5310
   ClientLeft      =   2265
   ClientTop       =   2235
   ClientWidth     =   7635
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7635
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   8281
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":030A
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   3885
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer UndoT 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   2070
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3585
      Top             =   2415
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   3885
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4286
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "10/17/00"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "1:49 PM"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox TextTS 
      Height          =   300
      Left            =   3060
      TabIndex        =   2
      Top             =   2280
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   529
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":03D3
   End
   Begin MSComctlLib.ImageList IL 
      Left            =   3120
      Top             =   960
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
            Picture         =   "frmMain.frx":049C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1198
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3080
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":395C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":420C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox color1 
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4660
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4772
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4884
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4996
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4AA8
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4BBA
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CCC
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4DDE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4EF0
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5002
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5114
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5226
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5338
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":544A
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":555C
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":566E
            Key             =   "Justify"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "Create a new file."
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open existing file."
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save current file"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print current file."
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut selected text."
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy to clipboard."
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste from clipboard."
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Description     =   "Delete selected text."
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Description     =   "Bold font."
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Description     =   "Italic font."
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
            Style           =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Description     =   "Underline font."
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
            Style           =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Description     =   "Find text in document."
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Description     =   "Align text left."
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Description     =   "Align text center."
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Description     =   "Align text right."
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   720
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label8 
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label10 
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label12 
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu newhtmldoc 
         Caption         =   "&New HTML Document"
      End
      Begin VB.Menu open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Save As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu minimize 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu sa 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu undo 
         Caption         =   "&Undo"
         Shortcut        =   {F4}
      End
      Begin VB.Menu clear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu linme243234 
         Caption         =   "-"
      End
      Begin VB.Menu find 
         Caption         =   "&Find"
      End
   End
   Begin VB.Menu fonts 
      Caption         =   "&Fonts"
      Begin VB.Menu font 
         Caption         =   "&Font"
      End
      Begin VB.Menu color 
         Caption         =   "&Color"
      End
      Begin VB.Menu line99 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontsBold 
         Caption         =   "&Bold"
      End
      Begin VB.Menu mnuFontsItalic 
         Caption         =   "&Italic"
      End
      Begin VB.Menu mnuFontsUnderline 
         Caption         =   "&Underlined"
      End
   End
   Begin VB.Menu view5 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolBar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusbar 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu piewviewhtml 
         Caption         =   "&Preview HTML"
      End
      Begin VB.Menu line2323 
         Caption         =   "-"
      End
      Begin VB.Menu spellchecker 
         Caption         =   "&Spell Checker"
         Shortcut        =   {F6}
      End
      Begin VB.Menu line100 
         Caption         =   "-"
      End
      Begin VB.Menu totalnumberofwords 
         Caption         =   "&Total Number of Words"
      End
      Begin VB.Menu fd 
         Caption         =   "&Total Number of Characters"
      End
   End
   Begin VB.Menu insert 
      Caption         =   "&Insert"
      Begin VB.Menu tags 
         Caption         =   "&HTML Tags"
         Begin VB.Menu bold 
            Caption         =   "&Bold"
         End
         Begin VB.Menu italic 
            Caption         =   "&Italic"
         End
         Begin VB.Menu underline 
            Caption         =   "&Underline"
         End
         Begin VB.Menu strikethru 
            Caption         =   "&Strikethru"
         End
         Begin VB.Menu hr 
            Caption         =   "&Horizontal Rule"
         End
         Begin VB.Menu paragraph 
            Caption         =   "&Paragraph"
         End
         Begin VB.Menu space 
            Caption         =   "&Space"
         End
      End
      Begin VB.Menu align 
         Caption         =   "&Align"
         Begin VB.Menu paraghraphiccenter 
            Caption         =   "&Paragraph Center"
         End
         Begin VB.Menu left 
            Caption         =   "&Paragraph Left"
         End
         Begin VB.Menu pr 
            Caption         =   "&Parahraph Right"
         End
         Begin VB.Menu center 
            Caption         =   "&Center"
         End
      End
      Begin VB.Menu insertlink 
         Caption         =   "&HTML Insert Link"
      End
      Begin VB.Menu insertimage 
         Caption         =   "&HTML Insert Image"
      End
      Begin VB.Menu table 
         Caption         =   "&Table"
      End
      Begin VB.Menu line343 
         Caption         =   "-"
      End
      Begin VB.Menu timeanddate 
         Caption         =   "&Time and Date"
         Shortcut        =   {F7}
      End
      Begin VB.Menu insertpictre 
         Caption         =   "&Insert Picture"
      End
   End
   Begin VB.Menu hwelp 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DocChanged As Boolean
Private Sub FileOpen()
Dim Directory As String
Dim TextT As String
Start:
cmDialog.FileName = ""
cmDialog.DialogTitle = "Open"
cmDialog.InitDir = App.Path
cmDialog.Filter = "SDI Documents *.sdi|*.SDI|Text Files *.txt|*.TXT|HTML Files *.html|*.HTML|HTM Files *.htm|*.HTM|All files|*.*"
cmDialog.ShowOpen
If cmDialog.FileName <> "" Then
Directory$ = cmDialog.FileName
Else
Exit Sub
End If
If FileExists(Directory$) = False Then
MsgBox "The file you specified does not exist.", 48, "Error"
GoTo Start
End If
OpenFile:
frmMain.TextTS.Text = ""
frmMain.TextTS.LoadFile Directory$, rtfText
frmMain.Text1.Text = TextTS.Text

     frmMain.Caption = "Power Notepad Example by nk - nkillaz.com [" & cmDialog.FileName & "]"
   
End Sub

Private Sub save_two()
Dim Directory As String
Start:
cmDialog.FileName = ""
cmDialog.DialogTitle = "Save As..."
cmDialog.InitDir = App.Path
cmDialog.Filter = "SDI Documents *.sdi|*.SDI|Text Files *.txt|*.TXT|HTML Files *.html|*.HTML|HTM Files *.htm|*.HTM|All files|*.*"
cmDialog.ShowSave
If cmDialog.FileName <> "" Then
Directory$ = cmDialog.FileName
Else
Exit Sub
End If
If FileExists(Directory$) = True Then
  If MsgBox("Do you want to overwrite the previous file?", 48 + vbYesNo, "SDI Word") = vbYes Then
  GoTo SaveFile
  Else
  GoTo Start
  End If
End If
SaveFile:
TextTS.Text = Text1.Text
TextTS.SaveFile Directory$, 1
 frmMain.Caption = "SDI Word 1.0 BETA 1 [" & cmDialog.FileName & "]"
 End
End Sub





Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub bold_Click()
Text1.SelText = "<b>" & Text1.SelText & "</b>"

End Sub

Private Sub center_Click()
Text1.SelText = "<center>" & Text1.SelText & "</center>"

End Sub

Private Sub clear_Click()
Text1.Text = " "
End Sub



Public Sub RefreshM()
If Text1.SelBold = True Then
TB.Buttons.Item(6).Value = tbrPressed
Else
TB.Buttons.Item(6).Value = tbrUnpressed
End If
If Text1.SelItalic = True Then
TB.Buttons.Item(7).Value = tbrPressed
Else
TB.Buttons.Item(7).Value = tbrUnpressed
End If
If Text1.SelUnderline = True Then
TB.Buttons.Item(8).Value = tbrPressed
Else
TB.Buttons.Item(8).Value = tbrUnpressed
End If
End Sub






Private Sub color_Click()
On Error GoTo Error_Event:
CommonDialog1.ShowColor
color1.BackColor = CommonDialog1.color
 Text1.SelColor = color1.BackColor
Error_Event:
    Exit Sub
   
End Sub







Private Sub copy_Click()
  Clipboard.SetText Text1.SelText
End Sub

Private Sub cut_Click()

  Clipboard.clear
  Clipboard.SetText Text1.SelText
 Text1.SelText = ""
End Sub



Private Sub exit_Click()

If DocChanged Then
    
    Select Case MsgBox("The file has changed." & vbCr & vbCr & _
            "Do you wish to save your changes?", _
            vbExclamation + vbYesNoCancel, frmMain.Caption)
    
    Case vbYes
    Call save_two
    Case vbNo
      End
    Case vbCancel
        Cancel = True
    
    End Select

Else
End
End If

End Sub

Private Sub fd_Click()
MsgBox "There are " & Label8.Caption & " characters.", vbInformation, "Word Count"

End Sub

Private Sub find_Click()
frmFind.Show
End Sub

Private Sub font_Click()
CDL1.Flags = cdlCFBoth Or cdlCFEffects
CDL1.ShowFont

With Text1
    .SelFontName = CDL1.FontName
    .SelFontSize = CDL1.FontSize
    .SelBold = CDL1.FontBold
    .SelItalic = CDL1.FontItalic
    .SelStrikeThru = CDL1.FontStrikethru
    .SelUnderline = CDL1.FontUnderline
    .SelColor = CDL1.color
End With

End Sub

Private Sub Form_Load()

Dim intFonts As Integer



  
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
Text1.Width = Me.Width - 100
Text1.Height = Me.Height - 1400

End Sub





Private Sub Form_Unload(Cancel As Integer)
   

If DocChanged Then
    
    Select Case MsgBox("The file has changed." & vbCr & vbCr & _
            "Do you wish to save your changes?", _
            vbExclamation + vbYesNoCancel, frmMain.Caption)
    
    Case vbYes
    save_Click
    Case vbNo
      End
    Case vbCancel
        Cancel = True
    
    End Select

ElseIf DocChanged = False Then
End
End If

End Sub

Private Sub hr_Click()
Text1.SelText = "<hr>" & Text1.SelText

End Sub

Private Sub insertimage_Click()
Text1.SelText = "<img src = ""http://"">" & Text1.SelText

End Sub

Private Sub insertlink_Click()
Text1.SelText = "<a href = ""http://""" & Text1.SelText

End Sub




Private Sub insertpictre_Click()
frmInsert.Show
End Sub

Private Sub italic_Click()
Text1.SelText = "<i>" & Text1.SelText & "</i>"

End Sub

Private Sub left_Click()
Text1.SelText = "<p align = ""left"">" & Text1.SelText

End Sub

Private Sub minimize_Click()
WindowState = 1
End Sub

Private Sub mnuFontsBold_Click()
If Text1.SelBold Then
    Text1.SelBold = False
    mnuFontsBold.Checked = False
Else
    Text1.SelBold = True
    mnuFontsBold.Checked = True
End If
End Sub

Private Sub mnuFontsItalic_Click()
If Text1.SelItalic Then
   Text1.SelItalic = False
    mnuFontsItalic.Checked = False
Else
    Text1.SelItalic = True
    mnuFontsItalic.Checked = True
End If
End Sub

Private Sub mnuFontsUnderline_Click()
If Text1.SelUnderline Then
    Text1.SelUnderline = False
    mnuFontsUnderline.Checked = False
Else
    Text1.SelUnderline = True
    mnuFontsUnderline.Checked = True
End If
End Sub



Private Sub mnuViewStatusbar_Click()
mnuViewStatusbar.Checked = Not mnuViewStatusbar.Checked
sbStatusBar.Visible = mnuViewStatusbar.Checked
End Sub

Private Sub mnuViewToolBar_Click()
  
mnuViewToolBar.Checked = Not mnuViewToolBar.Checked
tbToolBar.Visible = mnuViewToolBar.Checked


If tbToolBar.Visible = False Then
    Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 270
Else
  Text1.Move 0, 420, Me.ScaleWidth, Me.ScaleHeight - 690
End If
End Sub

Private Sub new_Click()
A = MsgBox("Are you sure you want to create a new file? The Previous file will not be saved.", vbYesNo, "Warning")
If A = vbYes Then
Text1.Text = ""
Picture1.Visible = False
  Text1.SelBold = False
    mnuFontsBold.Checked = False
      Text1.SelBold = False
    mnuFontsItalic.Checked = False
      Text1.SelBold = False
    mnuFontsUnderline.Checked = False
frmMain.Caption = "SDI Word 1.0 BETA 1 [New File]"
Else
Exit Sub
End If
End Sub

Private Sub newhtmldoc_Click()
A = MsgBox("Are you sure you want to create a new file? The Previous file will not be saved.", vbYesNo, "Warning")
If A = vbYes Then
Text1.Text = ""
  Text1.SelBold = False
    mnuFontsBold.Checked = False
      Text1.SelBold = False
    mnuFontsItalic.Checked = False
      Text1.SelBold = False
    mnuFontsUnderline.Checked = False
Picture1.Visible = False
frmMain.Caption = "SDI Word 1.0 BETA 1 [New File]"
Else
Exit Sub
End If
End Sub

Private Sub open_Click()
Call FileOpen
End Sub

Private Sub paraghraphiccenter_Click()
Text1.SelText = "<p align = ""center"">" & Text1.SelText

End Sub

Private Sub paragraph_Click()
Text1.SelText = "<p>" & Text1.SelText & "</p>"


End Sub

Private Sub paste_Click()
Text1.SelText = Clipboard.GetText()
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nader&
    ReleaseCapture
    nader& = SendMessage(Picture1.hWnd, &H112, &HF012, 0)
End Sub

Private Sub piewviewhtml_Click()
Open App.Path & "\preview.html" For Output As #1
Print #1, Text1.Text
Close #1
Load Browser
Browser.Show
Browser.Web.Navigate App.Path & "\preview.html"
End Sub

Private Sub pr_Click()
Text1.SelText = "<p align = ""right"">" & Text1.SelText

End Sub

Private Sub print_Click()
Dim bcancel As Boolean
Dim ncopy As Integer
On Error GoTo errorhandler

bcancel = False

CDL1.Flags = cdlPDHidePrintToFile Or _
        cdlPDNoSelection Or cdlPDNoPageNums _
        Or cdlPDCollate

CDL1.CancelError = True
CDL1.PrinterDefault = True
CDL1.Copies = 1
CDL1.ShowPrinter

If bcancel = False Then
    PrintRTF Text1, 1440, 1440, 1440, 1440
    For ncopy = 1 To CDL1.Copies
    Next ncopy
End If

Exit Sub

errorhandler:
If Err.Number = cdlCancel Then
bcancel = True
Resume Next
End If

End Sub

Private Sub sa_Click()
Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
End Sub



Private Sub save_Click()
Dim Directory As String
Start:
cmDialog.FileName = ""
cmDialog.DialogTitle = "Save As..."
cmDialog.InitDir = App.Path
cmDialog.Filter = "SDI Documents *.sdi|*.SDI|Text Files *.txt|*.TXT|HTML Files *.html|*.HTML|HTM Files *.htm|*.HTM|All files|*.*"
cmDialog.ShowSave
If cmDialog.FileName <> "" Then
Directory$ = cmDialog.FileName
Else
Exit Sub
End If
If FileExists(Directory$) = True Then
  If MsgBox("Do you want to overwrite the previous file?", 48 + vbYesNo, "SDI Word") = vbYes Then
  GoTo SaveFile
  Else
  GoTo Start
  End If
End If
SaveFile:
TextTS.Text = Text1.Text
TextTS.SaveFile Directory$, 1
 frmMain.Caption = "SDI Word 1.0 BETA 1 [" & cmDialog.FileName & "]"
End Sub

Private Sub space_Click()
Text1.SelText = "<BR>" & Text1.SelText

End Sub

Private Sub spellchecker_Click()
frmChecker.Show
End Sub

Private Sub strikethru_Click()
Text1.SelText = "<s>" & Text1.SelText & "</s>"

End Sub

Private Sub table_Click()
Text1.SelText = "<td>" & Text1.SelText & "</td>"

End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 1
  new_Click
  Case 2
  open_Click
  Case 3
  save_Click
  Case 4
  print_Click
  Case 6
    If Text1.SelBold = True Then
    Text1.SelBold = False
    Else
    Text1.SelBold = True
    End If
  Case 7
    If Text1.SelItalic = True Then
    Text1.SelItalic = False
    Else
    Text1.SelItalic = True
    End If
  Case 8
    If Text1.SelUnderline = True Then
    Text1.SelUnderline = False
    Else
    Text1.SelUnderline = True
    End If
  Case 10
  cmDialog.color = Text1.SelColor
  cmDialog.DialogTitle = "Select Color"
  cmDialog.ShowColor
  Text1.SelColor = cmDialog.color
  Case 11
  cmDialog.FontBold = Text1.SelBold
  cmDialog.FontItalic = Text1.SelItalic
  cmDialog.FontSize = Text1.SelFontSize
  cmDialog.FontName = Text1.SelFontName
  cmDialog.FontStrikethru = Text1.SelStrikeThru
  cmDialog.FontUnderline = Text1.SelUnderline
  font_Click
  Text1.SelBold = cmDialog.FontBold
  Text1.SelItalic = cmDialog.FontItalic
  Text1.SelFontSize = cmDialog.FontSize
  Text1.SelFontName = cmDialog.FontName
  Text1.SelStrikeThru = cmDialog.FontStrikethru
  Text1.SelUnderline = cmDialog.FontUnderline
  RefreshM
  Case 13
  TB.Customize
  Case 14
  about_Click
End Select
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
On Error Resume Next
  
    Select Case Button.Key
        
        Case "New"
           new_Click
        
        Case "Open"
            open_Click
        
        Case "Save"
            save_Click
        
        Case "Print"
           print_Click
        
        Case "Cut"
            cut_Click
        
        Case "Copy"
           copy_Click
        
        Case "Paste"
            paste_Click
        
        Case "Delete"
            cut_Click
        
        Case "Bold"
            
          If Text1.SelBold Then
    Text1.SelBold = False
    mnuFontsBold.Checked = False
Else
    Text1.SelBold = True
    mnuFontsBold.Checked = True
End If
        
        Case "Italic"
            
          If Text1.SelItalic Then
   Text1.SelItalic = False
    mnuFontsItalic.Checked = False
Else
    Text1.SelItalic = True
    mnuFontsItalic.Checked = True
End If
                        
        Case "Underline"
            
         If Text1.SelUnderline Then
    Text1.SelUnderline = False
    mnuFontsUnderline.Checked = False
Else
    Text1.SelUnderline = True
    mnuFontsUnderline.Checked = True
End If
            
        Case "Find"
            
            find_Click
                    
        Case "Align Left"
            
         
            Text1.SelAlignment = rtfLeft
        
        Case "Center"
            
          
           Text1.SelAlignment = rtfCenter
        
        Case "Align Right"
            
        
            Text1.SelAlignment = rtfRight
End Select

End Sub


Private Sub Text1_Change()
DocChanged = True
On Error Resume Next
Label2 = 0
Label4 = 0
Label6 = 0
Label8 = 0
Label10 = 0
Label12 = 0
If Text1 <> "" Then
If Mid(Text1.Text, 1, 1) <> " " Then
For i = 1 To Len(Text1.Text)
If Mid(Text1.Text, i, 1) = " " Then
Label2.Caption = Val(Label2.Caption) + 1
End If
Next
For i = 1 To Len(Text1.Text)
If Mid(Text1.Text, i, 2) = "  " Then
Label2.Caption = Label2.Caption - 1
End If
Next
Label2 = Label2 + 1
For i = 1 To Len(Text1.Text)
If Mid(Text1.Text, i, 1) = "a" Or Mid(Text1.Text, i, 1) = "e" Or Mid(Text1.Text, i, 1) = "i" Or Mid(Text1.Text, i, 1) = "o" Or Mid(Text1.Text, i, 1) = "u" Then
Label4.Caption = Val(Label4.Caption) + 1
ElseIf Mid(Text1.Text, i, 1) = "b" Or Mid(Text1.Text, i, 1) = "c" Or Mid(Text1.Text, i, 1) = "d" Or Mid(Text1.Text, i, 1) = "f" Or Mid(Text1.Text, i, 1) = "g" Or Mid(Text1.Text, i, 1) = "h" Or Mid(Text1.Text, i, 1) = "j" Or Mid(Text1.Text, i, 1) = "k" Or Mid(Text1.Text, i, 1) = "l" Or Mid(Text1.Text, i, 1) = "m" Or Mid(Text1.Text, i, 1) = "n" Or Mid(Text1.Text, i, 1) = "p" Or Mid(Text1.Text, i, 1) = "q" Or Mid(Text1.Text, i, 1) = "r" Or Mid(Text1.Text, i, 1) = "s" Or Mid(Text1.Text, i, 1) = "t" Or Mid(Text1.Text, i, 1) = "v" Or Mid(Text1.Text, i, 1) = "w" Or Mid(Text1.Text, i, 1) = "x" Or Mid(Text1.Text, i, 1) = "y" Or Mid(Text1.Text, i, 1) = "z" Then
Label6.Caption = Val(Label6.Caption) + 1
ElseIf Mid(Text1.Text, i, 1) = " " Then
Label12 = Label12 + 1
Else
Label10 = Label10 + 1
End If
Next
If Mid(Text1.Text, Len(Text1.Text), 1) = " " Then
Label2 = Label2 - 1
End If
End If
End If
If Text1.Text = "" Then
Label2 = 0
End If
Label8 = Len(Text1.Text)
End Sub

Private Sub timeanddate_Click()
Dim Text As String
Dim SelStart As Long


If Text1.SelLength > 0 Then
End If

Text = Text1.Text
SelStart = Text1.SelStart
Text1.Text = left(Text, SelStart) & Now & _
        Right(Text, Len(Text) - SelStart)

   
Text1.SelStart = SelStart

   

End Sub



Private Sub totalnumberofwords_Click()
MsgBox "There are " & Label2.Caption & " words.", vbInformation, "Word Count"
End Sub

Private Sub underline_Click()
Text1.SelText = "<u>" & Text1.SelText & "</u>"

End Sub

Private Sub undo_Click()
UndoT.Enabled = True
End Sub

Private Sub UndoT_Timer()
TextTS.SetFocus
SendKeys "(^)z"
Text1.Text = TextTS.Text
Text1.SetFocus
UndoT.Enabled = False
End Sub
