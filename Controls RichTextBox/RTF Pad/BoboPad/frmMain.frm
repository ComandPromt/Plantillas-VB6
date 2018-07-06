VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "BoboPad"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8385
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8385
   Begin VB.Timer PasteTimer 
      Interval        =   100
      Left            =   6600
      Top             =   1440
   End
   Begin VB.PictureBox PicLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5085
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   390
      Width           =   1215
      Begin RichTextLib.RichTextBox RTF 
         Height          =   615
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         _Version        =   393217
         ScrollBars      =   3
         OLEDragMode     =   0
         OLEDropMode     =   1
         TextRTF         =   $"frmMain.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "TBImages"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save As"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Select All"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Select All"
                  Text            =   "Select All"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Select Above"
                  Text            =   "Select Above"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Select Below"
                  Text            =   "Select Below"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Find"
            ImageIndex      =   12
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList TBImages 
         Left            =   6000
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":04C5
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0A5F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0FF9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1593
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1B2D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":20C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2661
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2BFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3195
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":32EF
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3401
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3513
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   5475
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10689
            Text            =   "File not saved"
            TextSave        =   "File not saved"
            Object.ToolTipText     =   "File path"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Cursor position"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Selection length"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "2 bytes"
            TextSave        =   "2 bytes"
            Object.ToolTipText     =   "File size"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTFtemp 
      DragIcon        =   "frmMain.frx":3625
      Height          =   615
      Left            =   6600
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":3777
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAssociations 
         Caption         =   "File Associations"
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuFileSP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEditBase 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Index           =   6
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select All"
         Index           =   8
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select Above"
         Index           =   9
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select Below"
         Index           =   10
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Datestamp"
         Index           =   12
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Find and Replace"
         Index           =   13
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuFormatWordwrap 
         Caption         =   "Wordwrap"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFormatSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatFont 
         Caption         =   "Font"
      End
      Begin VB.Menu mnuFormatBackcolor 
         Caption         =   "Backcolor"
      End
      Begin VB.Menu mnuFormatSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatStats 
         Caption         =   "Document Statistics"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Toolbar"
      End
      Begin VB.Menu mnuViewStatusbar 
         Caption         =   "Statusbar"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'******************************************************************
'Even though we are using a Rich Text Box so we have no file size
'limitations, this app will only display plain text - it's a Notepad
'clone. It can save in Rich Text Format and Word Document format,
'but will save the text only(no formatting). This requires quelling
'the Rich Text Box's desire to absorb objects dragged onto it
'and strict control of fonts. I might submit a Rich Text Editor
'in the future for those requiring such features
'******************************************************************

'******************************************************************
' Whilst we never intended to collaborate on this project, I'd
' like to thank Sebastian Thomschke for his help in finding
' the bugs I over-looked in my enthusiasm to make use of, and
' modify his excellent Undo/Redo classes to better suit the
' demands of this application. It's an example of how the PSC
' community works together to improve the standard of coding.
'******************************************************************

Public WithEvents Undo As clsUndo 'heavily modified version of a class by Sebastian Thomschke
Attribute Undo.VB_VarHelpID = -1
Dim onlyLoading As Boolean 'indicates form load complete
Dim mTStop() As Boolean 'allows the use of 'tab' within the richtextbox
'rather than moving focus to the next control
Dim myCommand As String
Private Sub Form_Load()
    'get settings from registry
    onlyLoading = True
    Me.Left = GetSetting("Psst Software\" + App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting("Psst Software\" + App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting("Psst Software\" + App.Title, "Settings", "MainWidth", 8500)
    Me.Height = GetSetting("Psst Software\" + App.Title, "Settings", "MainHeight", 6500)
    myCommand = Command()
    InitCmnDlg Me.hwnd
    cmndlg.flags = 5
    RTF.BackColor = Val(GetSetting("Psst Software\" + App.Title, "Settings", "Backcolor", Str(vbWhite)))
    mnuViewStatusbar.Checked = GetSetting("Psst Software\" + App.Title, "Settings", "Statusbar", True)
    mnuViewToolbar.Checked = GetSetting("Psst Software\" + App.Title, "Settings", "Toolbar", False)
    mnuFormatWordwrap.Checked = GetSetting("Psst Software\" + App.Title, "Settings", "Wordwrap", True)
    SB.Visible = mnuViewStatusbar.Checked
    TB.Visible = mnuViewToolbar.Checked
    RTF.RightMargin = IIf(mnuFormatWordwrap.Checked, 0, 200000)
    RTF.Text = " "
    RTF.SelStart = 0
    RTF.SelLength = 1
    SelFont
    RTF.Text = ""
    Set Undo = New clsUndo
    Undo.RichBox = RTF
    Undo.Reset
    FileChanged = False
End Sub

Private Sub Form_Paint()
    If onlyLoading Then
        If myCommand <> "" Then
            'We've been shelled
            DoEvents
            NoStatusUpdate = True
            Screen.MousePointer = 11
            SB.Panels(1) = "Loading file...."
            LockWindowUpdate Me.hwnd
            myCommand = strUnQuoteString(myCommand) 'sometimes explorer uses quotes('send to' for example)
            myCommand = GetLongFilename(myCommand) 'looks better than a dos path
            Select Case LCase(ExtOnly(myCommand))
                Case "txt"
                    RTF.SelText = OneGulp(myCommand) 'binary read
                Case "rtf"
                    RTFtemp.LoadFile myCommand 'rtf load
                    RTF.SelText = RTFtemp.Text
                Case "doc"
                    OpenWordDoc myCommand 'see sub - returns plain text
                Case Else
                    RTF.SelText = OneGulp(myCommand) 'otherwise do binary read
            End Select
            Me.Caption = FileOnly(myCommand)
            SB.Panels(1) = myCommand
            SB.Panels(4) = GetFileSize(Len(RTF.Text)) 'show size of file
            RTF.Tag = myCommand
            If FileLen(myCommand) > 100000 Then
                'just using SelFont, RTF selection falls
                'over somewhere around 100k so do this
                'slightly less efficient but more reliable
                'method of font control
                RTF.SelStart = 0
                RTF.SelLength = Len(RTF.Text)
                SelFont
                RTF.SelLength = 0
            End If
            NoStatusUpdate = False
            EditEnable
            RTF.SelStart = 0
            Screen.MousePointer = 0
            LockWindowUpdate 0
        End If
        Undo.Reset
        FileChanged = False
        onlyLoading = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Response As VbMsgBoxResult
    If FileChanged Then 'do we save current doc ?
        Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbYesNoCancel)
        Select Case Response
            Case vbCancel
                Cancel = 1 'dont unload, user must want to do something else after all
            Case vbYes
                'SaveAFile returns false if user cancels during save process - dont unload
                If Not SaveAFile Then Cancel = 1
        End Select
    End If
    'save settings to registry
    SaveSetting "Psst Software\" + App.Title, "Settings", "Wordwrap", mnuFormatWordwrap.Checked
    SaveSetting "Psst Software\" + App.Title, "Settings", "Statusbar", mnuViewStatusbar.Checked
    SaveSetting "Psst Software\" + App.Title, "Settings", "Toolbar", mnuViewToolbar.Checked
    SaveSetting "Psst Software\" + App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting "Psst Software\" + App.Title, "Settings", "MainTop", Me.Top
    SaveSetting "Psst Software\" + App.Title, "Settings", "MainWidth", Me.Width
    SaveSetting "Psst Software\" + App.Title, "Settings", "MainHeight", Me.Height
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    'placing the RTF on a left aligned picturebox makes
    'resizing easier
    PicLeft.Width = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'unload correctly
    Dim frm As Form
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Dim curvl As Long, st As Long
    LockWindowUpdate Me.hwnd
    Select Case Index
        Case 0 'Undo
            Undo.Undo
        Case 1 'Redo
            Undo.Redo
        Case 3 'cut
            Undo.Cut
        Case 4 'copy
            Undo.Copy
        Case 5 'paste
            Undo.Paste
        Case 6 'delete
            Undo.Delete
        Case 8 'select all
            RTF.SelStart = 0
            RTF.SelLength = Len(RTF.Text)
            RTF.SetFocus
        Case 9 'select above - maintaining scroll position
            Screen.MousePointer = 11
            'this is where the top line and therefore scroll position is
            curvl = SendMessage(RTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
            'do the selection
            st = RTF.SelStart
            RTF.SelStart = 0
            RTF.SelLength = st
            'return the scroll postion back to what it was - see SetScrollPos sub
            SetScrollPos curvl, RTF
            RTF.SetFocus
        Case 10 'select below
            Screen.MousePointer = 11
            curvl = SendMessage(RTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
            st = RTF.SelStart + RTF.SelLength
            RTF.SelStart = st
            RTF.SelLength = Len(RTF.Text) - st
            SetScrollPos curvl, RTF
            RTF.SetFocus
        Case 12 'date stamp
            frmDate.Show vbModal, Me
        Case 13 'find
            frmFind.Show , Me
    End Select
    LockWindowUpdate 0
    Screen.MousePointer = 0
End Sub

Private Sub mnuFileAbout_Click()
    'blah blah
    MsgBox "BoboPad - a better Notepad. PSST Software 2002" + vbCrLf + "BoboPad is Freeware" + vbCrLf + "Please visit www.psst.com.au"
End Sub

Private Sub mnuFileAssociations_Click()
    frmAssoc.Show vbModal, Me
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuFileNew_Click()
    Dim Response As VbMsgBoxResult
    If FileChanged Then 'do we save current doc ?
        Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbYesNoCancel)
        Select Case Response
            Case vbCancel
                Exit Sub
            Case vbYes
                If Not SaveAFile Then Exit Sub
        End Select
    End If
    'clear everything
    RTF.Text = ""
    Me.Caption = "Untitled.txt"
    SB.Panels(1) = "File not saved"
    SB.Panels(4) = GetFileSize(2)
    RTF.Tag = ""
    SelFont 'maintain control of fonts
    Undo.Reset 'clear undo buffer
    FileChanged = False
End Sub
Private Sub mnuFileOpen_Click()
    Dim Response As VbMsgBoxResult
    If FileChanged Then 'do we save current doc ?
        Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbYesNoCancel)
        Select Case Response
            Case vbCancel
                Exit Sub
            Case vbYes
                If Not SaveAFile Then Exit Sub
        End Select
    End If
    With cmndlg
        .filefilter = "Plain text (*.txt)|*.txt|Rich text (*.rtf)|*.rtf|Word Document (*.doc)|*.doc|All files (*.*)|*.*"
        OpenFile
        If Len(.filename) = 0 Then Exit Sub
        SB.Panels(1) = "Loading file...."
        NoStatusUpdate = True
        Me.Refresh
        Screen.MousePointer = 11 'hourglass
        LockWindowUpdate Me.hwnd 'RTF.hWnd
        RTF.Text = ""
        SelFont
        Select Case .filefilterindex
            Case 1
                RTF.SelText = OneGulp(.filename) 'binary read
            Case 2
                RTFtemp.LoadFile .filename 'rtf load
                RTF.SelText = RTFtemp.Text
            Case 3
                OpenWordDoc .filename 'see sub - returns plain text
            Case 4
                RTF.SelText = OneGulp(.filename) 'otherwise do binary read
        End Select
        If FileLen(.filename) > 100000 Then
            'just using SelFont, RTF selection falls
            'over somewhere around 100k so do this
            'slightly less efficient but more reliable
            'method of font control
            RTF.SelStart = 0
            RTF.SelLength = Len(RTF.Text)
            SelFont
            RTF.SelLength = 0
        End If
        Me.Caption = .filetitle
        SB.Panels(1) = .filename
        SB.Panels(4) = GetFileSize(Len(RTF.Text))
        RTF.Tag = .filename
        Undo.Reset 'clear undo buffer
        FileChanged = False 'reset need to save flag
        RTF.SelStart = 0
        NoStatusUpdate = False
        EditEnable
        LockWindowUpdate 0
        Screen.MousePointer = 0 'hourglass
    End With
End Sub
Private Sub mnuFilePageSetup_Click()
    ShowPageSetupDlg
End Sub
Private Sub mnuFilePrint_Click()
    ShowPrinter
End Sub
Private Sub mnuFileProperties_Click()
    Dim temp As Variant, z As Long, count As Long, msg As String
    Dim charcnt As Long, linecnt As Long, Mcount As Long
    If RTF.Tag <> "" And FileExists(RTF.Tag) Then
        'do a windows Property Dialog if we can
        GetPropDlg Me, RTF.Tag
    Else
        'otherwise just give document statistics
        linecnt = SendMessage(RTF.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&) 'line count
        charcnt = Len(RTF.Text) 'character count
        temp = Split(RTF.Text, Chr(32)) 'word count
        For z = 0 To UBound(temp)
            Select Case Trim(temp(z))
                Case vbNullString
                Case vbCrLf
                Case vbCr
                Case Else
                    Mcount = Mcount + 1
            End Select
        Next z
        msg = "File not yet saved." + vbCrLf
        msg = msg + "Words :" + Format(Mcount, "#,###,###,##0") + vbCrLf
        msg = msg + "Characters :" + Format(charcnt, "#,###,###,##0") + vbCrLf
        msg = msg + "Lines :" + Format(linecnt, "#,###,###,##0")
        MsgBox msg, vbInformation, "PSST SOFTWARE"
    End If
End Sub
Private Sub mnuFileSave_Click()
    SaveAFile
End Sub
Private Sub mnuFileSaveAs_Click()
    Dim sfile As String
    With cmndlg
        .filefilter = "Plain text (*.txt)|*.txt|Rich text (*.rtf)|*.rtf|Word Document (*.doc)|*.doc|All files (*.*)|*.*"
        .flags = 5 Or 2
        SaveFile
        If Len(.filename) = 0 Then Exit Sub
        sfile = .filename
        'make sure we have the correct extension
        Select Case .filefilterindex
            Case 1
                If InStr(1, sfile, ".") = 0 Then
                    sfile = sfile + ".txt"
                Else
                    sfile = ChangeExt(sfile, "txt")
                End If
                FileSave RTF.Text, sfile 'plain text
            Case 2
                If InStr(1, sfile, ".") = 0 Then
                    sfile = sfile + ".rtf"
                Else
                    sfile = ChangeExt(sfile, "rtf")
                End If
                RTF.SaveFile sfile 'rich text format
            Case 3
                If InStr(1, sfile, ".") = 0 Then
                    sfile = sfile + ".doc"
                Else
                    sfile = ChangeExt(sfile, "doc")
                End If
                SaveAsWordDoc sfile 'word document
            Case 4
                If InStr(1, sfile, ".") = 0 Then sfile = sfile + ".txt"
                FileSave RTF.Text, .filename 'plain text
        End Select
        FileChanged = False 'reset flag
        Me.Caption = FileOnly(sfile)
        SB.Panels(1) = sfile
        SB.Panels(4) = GetFileSize(Len(RTF.Text))
        RTF.Tag = sfile
    End With
End Sub
Private Sub mnuFormatBackcolor_Click()
    Dim col As Long 'new backcolor
    col = ShowColor
    If col <> -1 Then
        If col < 1 Then col = -col
        RTF.BackColor = col
        SaveSetting "Psst Software\" + App.Title, "Settings", "Backcolor", Str(col)
    End If
End Sub
Private Sub mnuFormatFont_Click()
    On Error GoTo woops
    Dim st As Long, curvl As Long, FontChange As Boolean
    'FileChanged is set to true by RTF change event (in the class)
    'As we are only changing font - not content, we dont want
    'this to alter due to this sub, so remember current state
    'so we can reset it below
    ChangeState = FileChanged
    Undo.IgnoreChange True  'dont add this action to the Undo buffer
    'current position
    curvl = SendMessage(RTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    st = RTF.SelStart
    With SelectFont
        .mFontName = GetSetting("Psst Software\" + App.Title, "Settings", "Fontname", "Lucida Console")
        .mFontsize = Val(GetSetting("Psst Software\" + App.Title, "Settings", "FontSize", "9"))
        .mBold = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "Bold", False))
        .mItalic = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "Italic", False))
        .mStrikethru = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "StrikeThru", False))
        .mUnderline = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "Underline", False))
        .mFontColor = Val(GetSetting("Psst Software\" + App.Title, "Settings", "Color", Str(vbBlack)))
        ShowFont
        If .mFontName <> GetSetting("Psst Software\" + App.Title, "Settings", "Fontname", "Lucida Console") Then FontChange = True
        If .mFontsize <> .mFontsize = Val(GetSetting("Psst Software\" + App.Title, "Settings", "FontSize", "9")) Then FontChange = True
        If .mBold <> .mBold = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "Bold", False)) Then FontChange = True
        If .mItalic <> .mItalic = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "Italic", False)) Then FontChange = True
        If .mStrikethru <> .mStrikethru = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "StrikeThru", False)) Then FontChange = True
        If .mUnderline <> .mUnderline = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "Underline", False)) Then FontChange = True
        If .mFontColor <> .mFontColor = Val(GetSetting("Psst Software\" + App.Title, "Settings", "Color", Str(vbBlack))) Then FontChange = True
        If Not FontChange Then GoTo woops
        'save new font
        SaveSetting "Psst Software\" + App.Title, "Settings", "Fontname", .mFontName
        SaveSetting "Psst Software\" + App.Title, "Settings", "FontSize", Str(.mFontsize)
        SaveSetting "Psst Software\" + App.Title, "Settings", "Bold", .mBold
        SaveSetting "Psst Software\" + App.Title, "Settings", "Italic", .mItalic
        SaveSetting "Psst Software\" + App.Title, "Settings", "StrikeThru", .mStrikethru
        SaveSetting "Psst Software\" + App.Title, "Settings", "Underline", .mUnderline
        SaveSetting "Psst Software\" + App.Title, "Settings", "Color", Str(.mFontColor)
        'implement on our new font
        LockWindowUpdate Me.hwnd
        RTF.SelStart = 0
        RTF.SelLength = Len(RTF.Text)
        RTF.SelColor = .mFontColor
        RTF.SelFontName = .mFontName
        RTF.SelFontSize = .mFontsize
        RTF.SelBold = .mBold
        RTF.SelItalic = .mItalic
        RTF.SelStrikeThru = .mStrikethru
        RTF.SelUnderline = .mUnderline
        RTF.SelStart = st
        RTF.SelLength = 0
        SetScrollPos curvl, RTF 'reset to current scroll position
    End With
woops:
    Undo.IgnoreChange False  'start using Undo system again
    FileChanged = ChangeState 'reset to what it was
    RTF.SetFocus
    LockWindowUpdate 0
End Sub

Private Sub mnuFormatStats_Click()
    Dim temp As Variant, z As Long, count As Long, msg As String
    Dim charcnt As Long, linecnt As Long, Mcount As Long
    linecnt = SendMessage(RTF.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&) 'lines
    charcnt = Len(RTF.Text) 'characters
    temp = Split(RTF.Text, Chr(32)) 'words
    For z = 0 To UBound(temp)
        Select Case Trim(temp(z))
            Case vbNullString
            Case vbCrLf
            Case vbCr
            Case Else
                Mcount = Mcount + 1
        End Select
    Next z
    msg = IIf(RTF.Tag = "", "File not yet saved.", RTF.Tag) + vbCrLf
    msg = msg + "Words :" + Format(Mcount, "#,###,###,##0") + vbCrLf
    msg = msg + "Characters :" + Format(charcnt, "#,###,###,##0") + vbCrLf
    msg = msg + "Lines :" + Format(linecnt, "#,###,###,##0")
    MsgBox msg, vbInformation, "PSST SOFTWARE"
End Sub

Private Sub mnuFormatWordwrap_Click()
    mnuFormatWordwrap.Checked = Not mnuFormatWordwrap.Checked
    RTF.RightMargin = IIf(mnuFormatWordwrap.Checked, 0, 200000)
End Sub
Private Sub mnuViewStatusbar_Click()
    mnuViewStatusbar.Checked = Not mnuViewStatusbar.Checked
    SB.Visible = mnuViewStatusbar.Checked
End Sub
Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    TB.Visible = mnuViewToolbar.Checked
End Sub
Private Sub PasteTimer_Timer()
    'you could hook the clipboard, but this will do
    mnuEdit(5).Enabled = Clipboard.GetFormat(vbCFText)
    TB.Buttons(8).Enabled = mnuEdit(5).Enabled
End Sub
Private Sub PicLeft_Resize()
    On Error Resume Next
    RTF.Width = PicLeft.Width - 120
    RTF.Height = PicLeft.Height
End Sub

Private Sub RTF_Change()
    If NoStatusUpdate Then Exit Sub
    SB.Panels(2) = RTF.SelStart
    SB.Panels(3) = RTF.SelLength
End Sub

Private Sub RTF_GotFocus()
    Dim z As Long 'allow tabs WITHIN richtextbox
    ReDim mTStop(0 To Controls.count - 1) As Boolean
    On Local Error Resume Next
    For z = 0 To Controls.count - 1
        mTStop(z) = Controls(z).TabStop
        Controls(z).TabStop = False
    Next
    SelFont
End Sub

Private Sub RTF_KeyDown(KeyCode As Integer, Shift As Integer)
    SelFont
    SB.Panels(2) = RTF.SelStart
    SB.Panels(3) = RTF.SelLength
End Sub

Private Sub RTF_LostFocus()
    Dim z As Long 'reset tabstops to original state
    On Local Error Resume Next
    For z = 0 To Controls.count - 1
        Controls(z).TabStop = mTStop(z)
    Next
End Sub

Private Sub RTF_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SelFont
    SB.Panels(2) = RTF.SelStart
    SB.Panels(3) = RTF.SelLength
End Sub

Private Sub RTF_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Me.PopupMenu mnuEditBase
End Sub

Private Sub RTF_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Response As VbMsgBoxResult, temp As String
    If Data.GetFormat(vbCFFiles) Then
        If FileChanged Then 'do we save current doc ?
            Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbYesNoCancel)
            Select Case Response
                Case vbCancel
                    Effect = vbDropEffectNone
                    Exit Sub
                Case vbYes
                    If Not SaveAFile Then
                        Effect = vbDropEffectNone
                        Exit Sub
                    End If
            End Select
        End If
        'Data.Files is a collection of the filepaths of files
        'dropped onto a control. Multiple files may be dropped
        'but in this app, we can only open one at a time
        'so we are only interested in Data.Files(1)
        temp = Data.Files(1)
        temp = strUnQuoteString(temp) 'sometimes explorer uses quotes('send to' for example)
        temp = GetLongFilename(temp) 'looks better than a dos path
        RTF.Text = ""
        SelFont
        Select Case LCase(ExtOnly(temp))
            Case "txt"
                RTF.SelText = OneGulp(temp) 'binary read
            Case "rtf"
                RTFtemp.LoadFile temp 'rtf load
                RTF.SelText = RTFtemp.Text
            Case "doc"
                OpenWordDoc temp 'see sub - returns plain text
            Case Else
                RTF.SelText = OneGulp(temp) 'otherwise do binary read
        End Select
        Me.Caption = FileOnly(temp)
        SB.Panels(1) = temp
        SB.Panels(4) = GetFileSize(Len(RTF.Text)) 'show size of file
        RTF.Tag = temp
        Undo.Reset
        FileChanged = False
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub RTF_SelChange()
    If NoStatusUpdate Then Exit Sub
    EditEnable
End Sub
Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
    'see menu items for comments
    Select Case Button.Index
        Case 1
            mnuFileNew_Click
        Case 2
            mnuFileOpen_Click
        Case 3
            mnuFileSave_Click
        Case 4
            mnuFileSaveAs_Click
        Case 6
            mnuEdit_Click 3
        Case 7
            mnuEdit_Click 4
        Case 8
            mnuEdit_Click 5
        Case 9
            mnuEdit_Click 6
        Case 11
            mnuEdit_Click 8
        Case 13
            mnuEdit_Click 0
        Case 14
            mnuEdit_Click 1
        Case 16
            mnuEdit_Click 13
    End Select
End Sub

Private Sub TB_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    mnuEdit_Click ButtonMenu.Index + 7
End Sub

Private Sub Undo_StateChanged()
    'enable menus/toolbar buttons according to undo class
    mnuEdit(0).Enabled = Undo.canUndo
    mnuEdit(1).Enabled = Undo.canRedo
    TB.Buttons(13).Enabled = Undo.canUndo
    TB.Buttons(14).Enabled = Undo.canRedo
    'Set FileChanged flag so we know if we need to save
    FileChanged = (Undo.canUndo = True Or Undo.canRedo = True)
End Sub
Private Sub OpenWordDoc(mfile As String)
    'standard call to Word to open a file
    'and get just the text
    Dim WordApp As Object
    On Error GoTo woops
    Screen.MousePointer = 11
    Set WordApp = CreateObject("Word.Application")
    WordApp.Documents.Open mfile
    WordApp.ActiveDocument.Content.Copy
    SelFont
    RTF.SelText = Clipboard.GetText(vbCFText)
    WordApp.Application.Quit
    Set WordApp = Nothing
    Screen.MousePointer = 0
    Exit Sub
woops:
    Set WordApp = Nothing
    Screen.MousePointer = 0
    MsgBox "Error converting Word Document", vbCritical
End Sub
Private Sub SaveAsWordDoc(mfile As String)
    ' get Word to save as .doc file
    'Why bother ? Well firstly, this is a demo of functionality
    'and I thought some people might like to know how to do this,
    'and secondly, when the file gets opened by Word in the future
    'sometimes Word is not happy with the fact that a text file
    'has a .doc extension and prompts for a new 'filter'
    'to be installed from CD or some other complaint - but still
    'opens the file. So... save as Word Document, avoid problems.
    Dim WordApp As Object
    Dim Document As Object
    On Error GoTo woops
    Screen.MousePointer = 11
    Set WordApp = CreateObject("Word.Application")
    Set Document = WordApp.Documents.Add
    Clipboard.Clear
    Clipboard.SetText RTF.Text, vbCFText
    WordApp.ActiveDocument.Content.Paste
    Document.SaveAs mfile
    WordApp.Application.Quit
    Set WordApp = Nothing
    Set Document = Nothing
    Screen.MousePointer = 0
    Exit Sub
woops:
    Set WordApp = Nothing
    Set Document = Nothing
    Screen.MousePointer = 0
    MsgBox "Error converting Word Document", vbCritical
End Sub
Public Function SaveAFile() As Boolean
    Dim Response As VbMsgBoxResult, sfile As String
    If Not FileExists(RTF.Tag) Then
        GoTo DoSaveAs 'must be a new file
    Else
        Select Case LCase(ExtOnly(RTF.Tag))
            Case "rtf"
                'if it was an existing .rtf file it will lose
                'formatting because we're only plain text
                'even though we'll save in Rich text format
                Response = MsgBox("Any rich text formatting in this file will be lost." + vbCrLf + "Do you wish to save this file using a different name ?", vbYesNoCancel)
                Select Case Response
                    Case vbCancel
                        SaveAFile = False
                        Exit Function
                    Case vbYes
                        GoTo DoSaveAs
                End Select
                RTF.SaveFile RTF.Tag
            Case "doc"
                'same as .rtf
                Response = MsgBox("Any document formatting in this file will be lost." + vbCrLf + "Do you wish to save this file using a different name ?", vbYesNoCancel)
                Select Case Response
                    Case vbCancel
                        SaveAFile = False
                        Exit Function
                    Case vbYes
                        GoTo DoSaveAs
                End Select
                SaveAsWordDoc RTF.Tag
            Case Else 'just plain text
                Kill RTF.Tag
                FileSave RTF.Text, RTF.Tag
        End Select
        FileChanged = False
        SaveAFile = True
    End If
    Exit Function
DoSaveAs:
    With cmndlg
        .filefilter = "Plain text (*.txt)|*.txt|Rich text (*.rtf)|*.rtf|Word Document (*.doc)|*.doc|All files (*.*)|*.*"
        .flags = 5 Or 2
        SaveFile
        If Len(.filename) = 0 Then
            SaveAFile = False
            Exit Function
        End If
        'make sure we have the correct extension
        Select Case .filefilterindex
            Case 1
                If InStr(1, sfile, ".") = 0 Then
                    sfile = sfile + ".txt"
                Else
                    sfile = ChangeExt(sfile, "txt")
                End If
                FileSave RTF.Text, sfile 'plain text
            Case 2
                If InStr(1, sfile, ".") = 0 Then
                    sfile = sfile + ".rtf"
                Else
                    sfile = ChangeExt(sfile, "rtf")
                End If
                RTF.SaveFile sfile 'rich text format
            Case 3
                If InStr(1, sfile, ".") = 0 Then
                    sfile = sfile + ".doc"
                Else
                    sfile = ChangeExt(sfile, "doc")
                End If
                SaveAsWordDoc sfile 'word document
            Case 4
                If InStr(1, sfile, ".") = 0 Then sfile = sfile + ".txt"
                FileSave RTF.Text, .filename 'plain text
        End Select
        Me.Caption = .filetitle
        SB.Panels(1) = .filename
        SB.Panels(4) = GetFileSize(Len(RTF.Text))
        RTF.Tag = .filename
        FileChanged = False 'reset flag
    End With
    SaveAFile = True
    
End Function

Public Sub SelFont()
    'This a plain text editor - just this font thanks
    'It is far more efficient to call this fast routine often
    'than do a SelectAll, ChangeFont, Select none routine -
    'particularly with large files
    RTF.SelFontName = GetSetting("Psst Software\" + App.Title, "Settings", "Fontname", "Lucida Console")
    RTF.SelFontSize = Val(GetSetting("Psst Software\" + App.Title, "Settings", "FontSize", "9"))
    RTF.SelBold = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "Bold", False))
    RTF.SelItalic = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "Italic", False))
    RTF.SelStrikeThru = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "StrikeThru", False))
    RTF.SelUnderline = CBool(GetSetting("Psst Software\" + App.Title, "Settings", "Underline", False))
    RTF.SelColor = Val(GetSetting("Psst Software\" + App.Title, "Settings", "Color", Str(vbBlack)))
End Sub


Public Sub EditEnable()
    'enable menus/toolbar buttons according to selection length
    Dim Enabled As Boolean
    Enabled = (RTF.SelLength > 0)
    mnuEdit(3).Enabled = Enabled
    mnuEdit(4).Enabled = Enabled
    mnuEdit(6).Enabled = Enabled
    TB.Buttons(6).Enabled = Enabled
    TB.Buttons(7).Enabled = Enabled
    TB.Buttons(9).Enabled = Enabled
    SB.Panels(2) = RTF.SelStart
    SB.Panels(3) = RTF.SelLength
End Sub
