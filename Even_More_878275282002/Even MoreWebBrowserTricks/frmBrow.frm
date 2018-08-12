VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBrow 
   Caption         =   "Even More WebBrowser Tricks"
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8340
   Icon            =   "frmBrow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   556
      TabIndex        =   7
      Top             =   0
      Width           =   8340
      Begin VB.CommandButton cmdGo 
         Height          =   315
         Left            =   2760
         Picture         =   "frmBrow.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Navigate"
         Top             =   60
         Width           =   375
      End
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   60
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox PicRight 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5445
      Left            =   2925
      ScaleHeight     =   5445
      ScaleWidth      =   5415
      TabIndex        =   2
      Top             =   435
      Width           =   5415
      Begin SHDocVwCtl.WebBrowser Brow 
         Height          =   735
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   615
         ExtentX         =   1085
         ExtentY         =   1296
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.PictureBox PicLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5445
      Left            =   0
      ScaleHeight     =   5445
      ScaleWidth      =   2895
      TabIndex        =   1
      Top             =   435
      Visible         =   0   'False
      Width           =   2895
      Begin MSComDlg.CommonDialog cmnDlg 
         Left            =   2040
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Flags           =   5
      End
      Begin MSComctlLib.ImageList TVImages 
         Left            =   1440
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrow.frx":058C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrow.frx":0B26
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrow.frx":10C0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   688
         ButtonWidth     =   2275
         ButtonHeight    =   688
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageCold"
         HotImageList    =   "ImageHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add...     "
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Organize..."
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageCold 
         Left            =   840
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   20
         ImageHeight     =   20
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrow.frx":165A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrow.frx":1C94
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageHot 
         Left            =   240
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   20
         ImageHeight     =   20
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrow.frx":22CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrow.frx":2908
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TV 
         Height          =   2415
         Left            =   0
         TabIndex        =   5
         Top             =   450
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4260
         _Version        =   393217
         Indentation     =   441
         Style           =   1
         ImageList       =   "TVImages"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11642
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Window"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "&Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOffline 
         Caption         =   "Work Offline"
      End
      Begin VB.Menu mnuFileSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties"
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
         Caption         =   "Undo"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Redo"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cut"
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Paste"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Delete"
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
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Find (on This Page)..."
         Index           =   10
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Find on local machine"
         Index           =   11
      End
   End
   Begin VB.Menu mnuNavigateBase 
      Caption         =   "&Navigate"
      Begin VB.Menu mnuNavigate 
         Caption         =   "&Back"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuNavigate 
         Caption         =   "&Forward"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuNavigate 
         Caption         =   "&Stop"
         Index           =   2
      End
      Begin VB.Menu mnuNavigate 
         Caption         =   "&Refresh"
         Index           =   3
      End
      Begin VB.Menu mnuNavigate 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuNavigate 
         Caption         =   "&Home"
         Index           =   5
      End
      Begin VB.Menu mnuNavigate 
         Caption         =   "&Search"
         Index           =   6
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewFavorites 
         Caption         =   "&Favorites"
      End
      Begin VB.Menu mnuViewAddressbar 
         Caption         =   "&Address bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusbar 
         Caption         =   "&Statusbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTextSize 
         Caption         =   "&Text Size"
         Begin VB.Menu mnuText 
            Caption         =   "Smallest"
            Index           =   0
         End
         Begin VB.Menu mnuText 
            Caption         =   "Smaller"
            Index           =   1
         End
         Begin VB.Menu mnuText 
            Caption         =   "Medium"
            Index           =   2
         End
         Begin VB.Menu mnuText 
            Caption         =   "Larger"
            Index           =   3
         End
         Begin VB.Menu mnuText 
            Caption         =   "Largest"
            Index           =   4
         End
      End
      Begin VB.Menu mnuViewSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSource 
         Caption         =   "&Source"
      End
      Begin VB.Menu mnuViewInternetOptions 
         Caption         =   "&Internet Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuTV 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuTV 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuTV 
         Caption         =   "Move to..."
         Index           =   2
      End
      Begin VB.Menu mnuTV 
         Caption         =   "Rename"
         Index           =   3
      End
      Begin VB.Menu mnuTV 
         Caption         =   "Delete"
         Index           =   4
      End
      Begin VB.Menu mnuTV 
         Caption         =   "Properties"
         Index           =   5
      End
      Begin VB.Menu mnuTV 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuTV 
         Caption         =   "Open folder"
         Index           =   7
      End
      Begin VB.Menu mnuTV 
         Caption         =   "New folder"
         Index           =   8
      End
      Begin VB.Menu mnuTV 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuTV 
         Caption         =   "Import Bookmarks"
         Index           =   10
      End
      Begin VB.Menu mnuTV 
         Caption         =   "Export Favorites"
         Index           =   11
      End
      Begin VB.Menu mnuTV 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuTV 
         Caption         =   "Refresh"
         Index           =   13
      End
   End
End
Attribute VB_Name = "frmBrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2001**********************
'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au

'Here is the third part of my Web Browser Tricks
'This time as a Web Browser(nearly)

'You need to reference Microsoft Internet Controls shdocvw.dll
'and Microsoft HTML Object Library mshtml.tbl

'What's here ?
'   - Favorites - a clone of Internet Explorer's Favorites menu
'     and a treeview, with access to IE's dialogs. The treeview has
'     a popup menu as an improvement over IE's.
'   - Autocomplete Address combo
'   - Menu's enabled/disabled by the Webbrowser
'   - Text resizing including identification of current size
'   - Progress bar in the Statusbar
'   - plus all the standard WebBrowser stuff

'To do - I'll leave it to you to add icons to the menu,
'   this will be easier than normal as we're already using
'   API to create the menu
'   - remove any inherrant bugs - this is only a demo !!

'I'm pretty sure all the code is mine - though using cut and paste
'may mean you recognize some procedures or routines, though I doubt it,
'- if so, thanks to the authors
'
'*************************************************************

'This form contains only standard WebBrowser stuff
'All the Favorites code is in a single re-usable module


Dim forwardenable As Boolean 'variables for Back/Forward
Dim backenable As Boolean
Dim cboselecting As Boolean 'used in controlling the cboAddress click event
Private Sub Brow_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Dim vStrPos As Long
    PB.Value = 0
    PB.Visible = SB.Visible
    PB.Move SB.Panels(2).Left + 30, SB.Top + 45, SB.Panels(2).Width - 60, SB.Height - 60
    If Brow.Busy = False Then
        If LCase(Left(URL, 4)) = "file" Then
            temp = URL
            temp = Replace(temp, "file:///", "")
            temp = Replace(temp, "/", "\")
            temp = Replace(temp, "%20", " ")
            cboAddress.Text = temp 'make local files appear readable
        Else
            cboAddress.Text = URL
        End If
        vStrPos = SendMessageByString&(cboAddress.hwnd, CB_FINDSTRING, 0, cboAddress.Text)
        If vStrPos = -1 Then
            cboAddress.AddItem cboAddress.Text 'add address if not there already
        End If
    End If
End Sub

Private Sub Brow_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
    On Error Resume Next
    Select Case Command 'identify Back/Forward state
        Case CSC_NAVIGATEFORWARD
                forwardenable = Enable
        Case CSC_NAVIGATEBACK
                backenable = Enable
    End Select
    mnuNavigate(0).Enabled = backenable
    mnuNavigate(1).Enabled = forwardenable
    If Command = -1 Then Exit Sub

End Sub

Private Sub Brow_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim temp As String
    Dim vStrPos As Long
    If Brow.Document Is Nothing Then Exit Sub
    On Error Resume Next
    If Brow.Busy = False Then
        If LCase(Left(Brow.LocationURL, 4)) = "file" Then
            temp = Brow.LocationURL
            temp = Replace(temp, "file:///", "")
            temp = Replace(temp, "/", "\")
            temp = Replace(temp, "%20", " ")
            cboAddress.Text = temp 'make local files appear readable
        Else
            cboAddress.Text = Brow.LocationURL
        End If
        vStrPos = SendMessageByString&(cboAddress.hwnd, CB_FINDSTRINGEXACT, 0, cboAddress.Text)
        If vStrPos = -1 Then
            cboAddress.AddItem cboAddress.Text 'add address if not there already
        End If
        PB.Value = 0
        PB.Visible = False
        Me.Refresh
        If Brow.LocationName <> "about:blank" Then
            Me.Caption = Brow.LocationName + " - Bobo Browser"
        Else
            Me.Caption = "Bobo Browser" 'Update caption accordingly
        End If
    End If

End Sub

Private Sub Brow_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Dim f As New frmBrow
    Set ppDisp = f.Brow.Object
    f.Show
End Sub

Private Sub Brow_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    On Error Resume Next
    If Not SB.Visible Then Exit Sub
    DoEvents
    If Progress = -1 Then PB.Value = 100
    If Progress > 0 And ProgressMax > 0 Then
        PB.Value = Progress * 100 / ProgressMax
        PB.Move SB.Panels(2).Left + 30, SB.Top + 45, SB.Panels(2).Width - 60, SB.Height - 60
        PB.Visible = True
    End If
    If PB.Value = 100 Then
        PB.Visible = False
    Else
        PB.Visible = True
        PB.Move SB.Panels(2).Left + 30, SB.Top + 45, SB.Panels(2).Width - 60, SB.Height - 60

    End If

End Sub

Private Sub Brow_StatusTextChange(ByVal Text As String)
    SB.Panels(1).Text = Text
End Sub

Private Sub cboAddress_Click()
    'only navigate if selected from dropdown
    If cboselecting Then cmdGo_Click
    cboselecting = False
End Sub

Private Sub cboAddress_DropDown()
    cboselecting = True 'enable navigation by clicking
End Sub

Private Sub cboAddress_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim z As Long, ll As Long, vStrPos As Long
    If KeyCode <> vbKeyBack Then
        If cboAddress.ListCount > 0 Then 'autocomplete routine
            z = SendMessageByString&(cboAddress.hwnd, CB_FINDSTRING, 0, cboAddress.Text)
            If z <> -1 And z <> cboAddress.ListIndex Then
                ll = Len(cboAddress.Text)
                cboAddress.ListIndex = z
                cboAddress.SelStart = ll
                cboAddress.SelLength = Len(cboAddress.Text) - ll
            End If
        End If
    End If
    If KeyCode = vbKeyReturn Then 'navigate if Return pressed
        Brow.Navigate cboAddress.Text
        vStrPos = SendMessageByString&(cboAddress.hwnd, CB_FINDSTRING, 0, cboAddress.Text)
        If vStrPos = -1 Then
            cboAddress.AddItem cboAddress.Text 'add address if not already there
            cboAddress.SelLength = 0
        End If
    End If
End Sub

Private Sub cmdGo_Click()
    Dim vStrPos As Long
    PicTop.SetFocus
    If Trim(cboAddress.Text) = "" Then Exit Sub
    Brow.Navigate cboAddress.Text 'navigate to address
    vStrPos = SendMessageByString&(cboAddress.hwnd, CB_FINDSTRINGEXACT, 0, cboAddress.Text)
    If vStrPos = -1 Then
        cboAddress.AddItem cboAddress.Text 'add address if not already there
    End If
End Sub
Private Sub Form_Load()
    Dim mycommand As String, v As Long
    'load up users favorites - see module
    'the treeview parameter is optional - if you only want the menu
    GetFaves Me.hwnd, TV
    mnuFileOffline.Checked = Brow.Offline
    'Read settings from last session
    mnuViewStatusbar.Checked = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "StatusbarVisible", True)
    SB.Visible = mnuViewStatusbar.Checked
    mnuViewAddressbar.Checked = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "AddressbarVisible", True)
    PicTop.Visible = mnuViewAddressbar.Checked
    mnuViewFavorites.Checked = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "FavoritesVisible", False)
    PicLeft.Visible = mnuViewFavorites.Checked
    Me.Left = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "MainWidth", 8685)
    Me.Height = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "MainHeight", 6330)
    Me.WindowState = 0
    Me.WindowState = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "WindowState", 2)
    DoEvents
    For v = 0 To 9 'load up addresses from registry
        temp = GetSetting("PSST SOFTWARE\" + App.Title, "Addresses", "Text" + Str(v))
        If temp <> "" Then
            vStrPos = SendMessageByString&(cboAddress.hwnd, CB_FINDSTRING, 0, temp)
            If vStrPos = -1 Then
                cboAddress.AddItem temp
            End If
        End If
    Next v
    If mycommand <> "" Then 'were we shelled ?
        Brow.Navigate mycommand 'yep
    Else
        Brow.Navigate2 "about:blank", 2 'open a blank page but dont add to travel log (back/forward history)
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim v As Integer, cnt As Long 'save settings to registry
    For v = cboAddress.ListCount - 1 To 0 Step -1
        If Len(Trim(cboAddress.List(v))) <> 0 Then
            SaveSetting "PSST SOFTWARE\" + App.Title, "Addresses", "Text" + Str(cnt), cboAddress.List(v)
            cnt = cnt + 1
        End If
        If cnt = 200 Then Exit For
    Next v
    If Me.WindowState = 2 Then
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "WindowState", 2
    Else
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "WindowState", 0
    End If
    If Me.WindowState <> 1 And Me.WindowState <> 2 Then
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "MainTop", Me.Top
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "MainHeight", Me.Height
    End If

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'I find it good practice to off-load resizing to
    'individual controls and in-built align functions
    If PicLeft.Visible Then
        PicRight.Width = Me.Width - PicLeft.Width - 150
    Else
        PicRight.Width = Me.Width - 150
    End If
End Sub


Private Sub mnuEdit_Click(Index As Integer)
    On Error Resume Next
    'These menu items will be enabled/disabled by
    'mnuEditBase_Click event
    Select Case Index
        Case 0
            Brow.ExecWB OLECMDID_UNDO, OLECMDEXECOPT_DODEFAULT
        Case 1
            Brow.ExecWB OLECMDID_REDO, OLECMDEXECOPT_DODEFAULT
        Case 3
            Brow.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
        Case 4
            Brow.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
        Case 5
            Brow.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
        Case 6
            Brow.ExecWB OLECMDID_DELETE, OLECMDEXECOPT_DODEFAULT
        Case 8
            Brow.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
        Case 10
            Brow.SetFocus
            SendKeys "^f", True
        Case 11
            On Error Resume Next
            Brow.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DONTPROMPTUSER
    End Select
End Sub


Private Sub mnuEditBase_Click()
    Dim mDoc As HTMLDocument
    Set mDoc = Brow.Document
    'enable/disable menus
    'some real power here - experiment with other calls
    mnuEdit(0).Enabled = mDoc.queryCommandEnabled("undo")
    mnuEdit(1).Enabled = mDoc.queryCommandEnabled("redo")
    mnuEdit(3).Enabled = mDoc.queryCommandEnabled("cut")
    mnuEdit(4).Enabled = mDoc.queryCommandEnabled("copy")
    mnuEdit(5).Enabled = mDoc.queryCommandEnabled("paste")
    mnuEdit(6).Enabled = mDoc.queryCommandEnabled("delete")
    mnuEdit(8).Enabled = mDoc.queryCommandEnabled("selectall")
End Sub

Private Sub mnuFileNew_Click()
    On Error Resume Next
    'you need to have compiled an exe for this to work
    Shell App.Path + "\" + App.EXEName + ".exe " + Brow.LocationURL
End Sub
Private Sub mnuFileOffline_Click()
    mnuFileOffline.Checked = Not mnuFileOffline.Checked
    Brow.Offline = mnuFileOffline.Checked
End Sub

Private Sub mnuFileOpen_Click()
    On Error GoTo woops
    With cmnDlg
        .Filter = "Web page (*.htm;*.html)|*.htm;*.html|Supported image formats|*.gif;*.tif;*.pcd;*.jpg;*.wmf;*.tga;*.jpeg;*.ras;*.png;*.eps;*.bmp;*.pcx|Text formats (*.txt;*.doc)|*.txt;*.doc|All files (*.*)|*.*"
        .Flags = 5
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        Me.Refresh
        Brow.Navigate .FileName
     End With
woops:
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    Brow.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    Brow.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuFileProperties_Click()
    On Error Resume Next
    Brow.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuFileSaveAs_Click()
    On Error GoTo woops
    If Brow.LocationURL = "" Then Exit Sub
    Brow.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
woops:
End Sub

Private Sub mnuHelpAbout_Click()
    Dim temp As String
    temp = "This is the third of my 'Web Browser Tricks' submissions to PSC." + vbCrLf + _
    "It demonstrates some more advanced options than the first two." + vbCrLf + _
    "This is not a full blown Web Browser - simply a demo." + vbCrLf + _
    "As usual all illegal operations and bugs are provided free of charge." + vbCrLf + _
    "If you find this demo helpfull you might consider giving me a vote !" + vbCrLf + vbCrLf + _
    "By MrBobo - ©PSST Software 2001"
    MsgBox temp, vbInformation, "PSST Software"
End Sub

Private Sub mnuNavigate_Click(Index As Integer)
    Select Case Index
        Case 0
            Brow.GoBack
        Case 1
            Brow.GoForward
        Case 2
            Brow.Stop
        Case 3
            Brow.Refresh
        Case 4
            Brow.GoHome
        Case 5
            Brow.GoSearch
    End Select
End Sub

Private Sub mnuText_Click(Index As Integer)
    Dim z As Long 'set font size
    On Error Resume Next
    For z = 0 To 4
        mnuText(z).Checked = False
    Next
    mnuText(Index).Checked = True
    Brow.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DODEFAULT, CLng(Index)

End Sub

Private Sub mnuView_Click()
    'determine current font size
    'this function is usually missing in other examples I've seen
    Dim q
    Dim z As Long
    On Error Resume Next
    For z = 0 To 4
        mnuText(z).Checked = False
    Next
    Brow.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DODEFAULT, , q
    mnuText(q).Checked = True

End Sub

Private Sub mnuViewAddressbar_Click()
    mnuViewAddressbar.Checked = Not mnuViewAddressbar.Checked
    PicTop.Visible = mnuViewAddressbar.Checked
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "AddressbarVisible", mnuViewAddressbar.Checked
    Form_Resize
End Sub

Private Sub mnuViewFavorites_Click()
    mnuViewFavorites.Checked = Not mnuViewFavorites.Checked
    PicLeft.Visible = mnuViewFavorites.Checked
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "FavoritesVisible", mnuViewFavorites.Checked
    Form_Resize
End Sub

Private Sub mnuViewInternetOptions_Click()
    'control panel applet
    Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl"
End Sub

Private Sub mnuViewSource_Click()
    Dim mDoc As HTMLDocument
    Dim temp As String
    Dim f As Integer
    f = FreeFile
    Set mDoc = Brow.Document
    temp = mDoc.documentElement.outerHTML
    If FileExists(App.Path + "\tmp.tmp") Then Kill App.Path + "\tmp.tmp"
    Open App.Path + "\tmp.tmp" For Binary As #f
        Put #f, , temp
    Close #f
    Shell "notepad.exe " + App.Path + "\tmp.tmp", vbNormalFocus
    'Alternate method suggested by Juha Söderqvist
    'Brow.Navigate "view-source:" & Brow.LocationURL
    'neat eh ?
    'opens in notepad
    'Advantages: one line of code
    'Disadvantages:needs on error resume next because if the
    'URL is not valid will fail the HTMLDocument method will always return
    'the source of the current page
    'Also you have to use notepad - if you want to display
    'source code in your own textbox use the HTMLDocument method
End Sub

Private Sub mnuViewStatusbar_Click()
    mnuViewStatusbar.Checked = Not mnuViewStatusbar.Checked
    SB.Visible = mnuViewStatusbar.Checked
    PB.Move SB.Panels(2).Left + 30, SB.Top + 45, SB.Panels(2).Width - 60, SB.Height - 60
    PB.Visible = SB.Visible 'hide progressbar accordingly
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "StatusbarVisible", mnuViewStatusbar.Checked
End Sub

Private Sub PicLeft_Resize()
    On Error Resume Next
    Toolbar1.Width = PicTVBase.ScaleWidth
    TV.Width = PicLeft.ScaleWidth
    TV.Height = PicLeft.ScaleHeight - TV.Top
End Sub
Private Sub PicTop_Resize()
    On Error Resume Next
    cmdGo.Left = PicTop.ScaleWidth - cmdGo.Width - 8
    cboAddress.Width = cmdGo.Left - cboAddress.Left - 4
End Sub
Private Sub PicRight_Resize()
    On Error Resume Next
    Brow.Width = PicRight.Width
    Brow.Height = PicRight.Height
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            AddFaves Me.hwnd 'see module
        Case 2
            OrgFaves Me.hwnd 'see module
    End Select
End Sub
Private Sub mnuPopup_Click()
    If TV.SelectedItem Is Nothing Then TV.Nodes(1).Selected = True
    If TV.Nodes(1).Selected = True Then
        mnuTV(2).Enabled = False
        mnuTV(3).Enabled = False
        mnuTV(4).Enabled = False
    Else
        mnuTV(2).Enabled = True
        mnuTV(3).Enabled = True
        mnuTV(4).Enabled = True
    End If

End Sub
Private Sub mnuTV_Click(Index As Integer)
    Dim temp As String, temp1 As String, z As Long, IsDir As Boolean
    If Right(TV.SelectedItem.Key, 1) = "\" Then IsDir = True
    Select Case Index
        Case 0 'add favorite
            If mnuTV(0).Caption = "Add to Favorites..." Then
                AddFaves Me.hwnd
                Exit Sub
            End If
            If IsDir Then
                temp = TV.SelectedItem.Key + ChangeExt(Brow.LocationName, "url")
            Else
                temp = TV.SelectedItem.parent.Key + ChangeExt(Brow.LocationName, "url")
            End If
            temp = SafeSave(temp) 'make sure we have a unique file name
            WriteINI temp, "InternetShortcut", "URL", Brow.LocationURL 'create an internet shortcut
            LockWindowUpdate TV.hwnd
            RefreshFaves Me.hwnd 'reload tv and menu so we pick up changes
            LockWindowUpdate 0
        Case 2 'move favorite
            temp = BrowseForFolder(Me.hwnd)
            If temp = "" Then Exit Sub
            temp1 = TV.SelectedItem.Key
            If LCase(TV.SelectedItem.Key) = LCase(temp) Then Exit Sub
            If IsDir Then temp1 = Left(temp1, Len(temp1) - 1)
            temp = temp + "\" + mID$(temp1, InStrRev(temp1, "\") + 1)
            MoveFave temp1, temp ' see module
            LockWindowUpdate TV.hwnd
            RefreshFaves Me.hwnd 'reload tv and menu so we pick up changes
            If IsDir Then temp = temp + "\"
            For z = 1 To TV.Nodes.Count
                If TV.Nodes(z).Key = temp Then
                    TV.Nodes(z).Expanded = True
                    TV.Nodes(z).Selected = True
                    TV.Nodes(z).EnsureVisible
                    Exit For
                End If
            Next
            LockWindowUpdate 0
        Case 3 'rename
            TV.StartLabelEdit
        Case 4 'delete
            temp1 = TV.SelectedItem.Key
            temp = TV.SelectedItem.parent.Key
            If IsDir Then temp1 = Left(temp1, Len(temp1) - 1)
            DeleteFave temp1 'see module
            LockWindowUpdate TV.hwnd
            RefreshFaves Me.hwnd 'reload tv and menu so we pick up changes
            For z = 1 To TV.Nodes.Count
                If TV.Nodes(z).Key = temp Then
                    TV.Nodes(z).Expanded = True
                    TV.Nodes(z).Selected = True
                    TV.Nodes(z).EnsureVisible
                    Exit For
                End If
            Next
            LockWindowUpdate 0
        Case 5 'properties
            temp1 = TV.SelectedItem.Key
            If IsDir Then temp1 = Left(temp1, Len(temp1) - 1)
            GetPropDlg Me, temp1 'see module
        Case 7 'open folder in explorer
            If IsDir Then
                temp = TV.SelectedItem.Key
            Else
                temp = TV.SelectedItem.parent.Key
            End If
            Shell "explorer.exe " + temp, vbNormalFocus
        Case 8 'new folder
            If IsDir Then
                temp = TV.SelectedItem.Key
            Else
                temp = TV.SelectedItem.parent.Key
            End If
            LockWindowUpdate TV.hwnd
            MkDir SafeSave(temp + "New folder") 'get a unique name
            RefreshFaves Me.hwnd
            temp = temp + safesavename + "\"
            For z = 1 To TV.Nodes.Count
                If TV.Nodes(z).Key = temp Then
                    TV.Nodes(z).Expanded = True
                    TV.Nodes(z).Selected = True
                    TV.Nodes(z).EnsureVisible
                    TV.SetFocus
                    TV.StartLabelEdit
                    Exit For
                End If
            Next
            LockWindowUpdate 0
        Case 10
            BrowDlg.ImportExportFavorites True, "" 'show IE dialog
        Case 11
            BrowDlg.ImportExportFavorites False, "" 'show IE dialog
        Case 13 'refresh
            LockWindowUpdate TV.hwnd
            RefreshFaves Me.hwnd 'reload tv and menu so we pick up changes
            LockWindowUpdate 0
    End Select
End Sub
Private Sub TV_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim temp As String, temp1 As String, z As Long, IsDir As Boolean, ExtStr As String
    Dim isExpanded As Boolean
    temp = TV.SelectedItem.Key
    temp1 = TV.SelectedItem.Key
    If Right(temp, 1) = "\" Then
        IsDir = True
        isExpanded = TV.SelectedItem.Expanded
    Else
        ExtStr = ".url"
    End If
    If NewString = TV.SelectedItem.Text Then
        Cancel = 1
        Exit Sub
    End If
    LockWindowUpdate TV.hwnd
    If IsDir Then temp = Left(temp, Len(temp) - 1)
    RenameFave temp, PathOnly(temp) + "\" + NewString + ExtStr 'see module
    RefreshFaves Me.hwnd
    If IsDir Then
        temp1 = PathOnly(temp) + "\" + NewString + "\"
    Else
        temp1 = PathOnly(temp) + "\" + NewString + ExtStr
    End If
    For z = 1 To TV.Nodes.Count
        If TV.Nodes(z).Key = temp1 Then
            If IsDir Then
                If isExpanded Then TV.Nodes(z).Expanded = True
            Else
                TV.Nodes(z).Expanded = True
            End If
            TV.Nodes(z).Selected = True
            TV.Nodes(z).EnsureVisible
            Exit For
        End If
    Next
    LockWindowUpdate 0
End Sub
Private Sub TV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim fred As Node
    Dim temp As String
    Set fred = TV.HitTest(x, y)
    If fred Is Nothing Then Exit Sub
    If fred.Index = 1 Then Exit Sub
    Select Case Button
        Case 1
            If Right(fred.Key, 1) <> "\" Then
                If Not FileExists(fred.Key) Then Exit Sub
                temp = ReadINI(fred.Key, "InternetShortcut", "URL") 'get address from URL file
                If temp = "" Then Exit Sub
                If Brow.LocationURL <> temp Then Brow.Navigate temp 'navigate there
            End If
        Case 2
            If Right(fred.Key, 1) = "\" Then
                Me.PopupMenu mnuPopup 'show menu
            Else
                Me.PopupMenu mnuPopup
            End If
    End Select
End Sub

