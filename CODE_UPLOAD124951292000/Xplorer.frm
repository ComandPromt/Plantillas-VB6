VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{FC07EBD4-FE92-11D0-A199-A0077383D901}#5.5#0"; "CCRPPRG.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Xplorer"
   ClientHeight    =   4470
   ClientLeft      =   495
   ClientTop       =   2160
   ClientWidth     =   7740
   Icon            =   "Xplorer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   7740
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward"
      DownPicture     =   "Xplorer.frx":27A2
      Height          =   660
      Left            =   2640
      Picture         =   "Xplorer.frx":2F50
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      DownPicture     =   "Xplorer.frx":3202
      Height          =   660
      Index           =   1
      Left            =   3600
      Picture         =   "Xplorer.frx":3A14
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   735
   End
   Begin CCRProgressBar.ccrpProgressBar ccrpProgressBar1 
      Height          =   300
      Left            =   4440
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   529
      AutoCaption     =   3
      Caption         =   "0 "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IncrementSize   =   1
      Picture         =   "Xplorer.frx":3F06
      Smooth          =   -1  'True
      Style           =   1
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   300
      Left            =   6960
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Text            =   "Input your search phrase in here"
      ToolTipText     =   "Input the search phrase here"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Text            =   "Bookmarks"
      ToolTipText     =   "Browse you bookmarks here"
      Top             =   1200
      Width           =   5175
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      DownPicture     =   "Xplorer.frx":4F05C
      Height          =   660
      Left            =   840
      Picture         =   "Xplorer.frx":4F86E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Type in the URL's here"
      Top             =   720
      Width           =   7335
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      DownPicture     =   "Xplorer.frx":4FC8C
      Height          =   660
      Left            =   0
      Picture         =   "Xplorer.frx":5049E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      DownPicture     =   "Xplorer.frx":50954
      Height          =   660
      Left            =   1800
      Picture         =   "Xplorer.frx":50FC6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   2990
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
   Begin VB.CommandButton cmdGo 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -120
      Picture         =   "Xplorer.frx":51278
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Go!"
      Top             =   1560
      Width           =   375
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   0
      X2              =   4320
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   0
      Width           =   3855
      WordWrap        =   -1  'True
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Home 
         Caption         =   "Homepage"
      End
      Begin VB.Menu NewWindow 
         Caption         =   "New Window"
         Shortcut        =   ^N
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Properties 
         Caption         =   "Properties"
      End
      Begin VB.Menu PageSetp 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu Print 
         Caption         =   "Print"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu SelectAll 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu Bookmarks 
      Caption         =   "Bookmarks"
      Begin VB.Menu ClearBkmarks 
         Caption         =   "Clear Bookmarks"
      End
      Begin VB.Menu Add 
         Caption         =   "Add a Bookmark"
         Shortcut        =   +{INSERT}
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Popups 
         Caption         =   "Disable Popups"
         Shortcut        =   ^P
      End
      Begin VB.Menu Clear 
         Caption         =   "Clear History"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu Bugs 
         Caption         =   "Bugs"
      End
      Begin VB.Menu About 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AllowPopups As Boolean
Public NumberOfTimesClicked As Integer
Public Index As Integer
Public State As Integer
Dim i As Integer
Dim mbDontNavigateNow As Boolean
    
Private Sub About_Click()
    Form2.Visible = True
End Sub

Private Sub Add_Click()
    AddBookmarks (App.Path & "\bookmarks.txt")
End Sub

Private Sub Bugs_Click()
    MsgBox ("There are a couple of bugs still in the program. The Edit functions don't work properly for now, but that'll get fixed soon. Also, the New Window function and the windows it opens don work quite right. :)")
End Sub

Private Sub Clear_Click()
    On Error Resume Next
    Combo1.Clear
    Combo1.Text = WebBrowser1.LocationURL
    Dim i As Integer
    Dim a As String
    Open App.Path & "\history.txt" For Output As #1
    For i = 0 To Combo1.ListCount - 1
    Write #1, ""
    Next i
    Close #1
End Sub

Private Sub ClearBkmarks_Click()
    On Error Resume Next
    Combo2.Clear
    Dim i As Integer
    Dim a As String
    Open App.Path & "\bookmarks.txt" For Output As #1
    For i = 0 To Combo2.ListCount - 1
    Write #1, ""
    Next i
    Close #1
    Combo2.Text = "Bookmarks"
End Sub

Private Sub cmbSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdSearch.Default = True
End Sub

Private Sub cmdForward_Click()
On Error GoTo disable2
WebBrowser1.GoForward
cmdBack.Enabled = True
disable2:
cmdForward.Enabled = False
cmdBack.Enabled = True
End Sub

Private Sub cmdGo_Click()
    WebBrowser1.Navigate (Combo1.Text)
    cmdBack.Enabled = True
End Sub

Private Sub cmdHome_Click(Index As Integer)
    WebBrowser1.GoHome
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    WebBrowser1.Refresh
End Sub

Private Sub cmdStop_Click()
    WebBrowser1.Stop
    ccrpProgressBar1.Value = 0
End Sub

Private Sub Combo2_Change()
    WebBrowser1.Navigate Combo2.SelText
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    WebBrowser1.Navigate ("http://search.dogpile.com/texis/search?q=" & cmbSearch.Text & "&geo=no&refer=dp-search&fs=web")
    cmbSearch.AddItem (cmbSearch.Text)
    cmbSearch.SetFocus
End Sub

Private Sub Copy_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Cut_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Form_GotFocus()
    Combo1.Clear
    Dim a As String
    On Error Resume Next
    Open App.Path & "\history.txt" For Input As #1
    Do
        Input #1, a
        If a <> "" Then
        Combo1.AddItem a
    End If
    Loop Until EOF(1)
    Close #1
End Sub

Private Sub Form_Load()
    i = 0
    cmdBack.Enabled = False
    cmdForward.Enabled = False
    State = 1
    WebBrowser1.GoHome
    NumberOfTimesClicked = 0
    AllowPopups = True
    LoadBookmarks
    Dim a As String
    On Error Resume Next
    Open App.Path & "\history.txt" For Input As #1
    Do
        Input #1, a
        If a <> "" Then
        Combo1.AddItem a
    End If
    Loop Until EOF(1)
    Close #1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Form1.WindowState <> 1 Then
        WebBrowser1.Width = Form1.ScaleWidth
        WebBrowser1.Height = Form1.ScaleHeight - 1525
        ccrpProgressBar1.Width = Form1.ScaleWidth - Line1.X2 - Line1.X1
        Combo1.Width = Form1.ScaleWidth - 200
        lblStatus.Width = Form1.ScaleWidth - Combo2.Width
        cmbSearch.Width = Form1.Width - Combo2.Width - cmdSearch.Width - 300
        cmdSearch.Left = Form1.ScaleWidth - cmdSearch.Width
    End If
End Sub

Private Sub cmdBack_Click()
On Error GoTo disable
WebBrowser1.GoBack
cmdForward.Enabled = True
disable:
cmdBack.Enabled = False
cmdForward.Enabled = True
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub LoadBookmarks()
    Dim a As String
    On Error Resume Next
    Open App.Path & "\bookmarks.txt" For Input As #1
    Do
    Input #1, a
    If a <> "" Then
    Combo2.AddItem a
    End If
    Loop Until EOF(1)
    Close #1
End Sub

Private Sub Home_Click()
    WebBrowser1.GoHome
End Sub

Private Sub NewWindow_Click()
    On Error Resume Next
    Static lDocumentCount As Long
    Dim frmD As Form
    lDocumentCount = lDocumentCount + 1
    Set frmD = New Form1
    frmD.Show
    frmD.SetFocus
End Sub

Private Sub AddBookmarks(filename As String)
    Combo2.AddItem WebBrowser1.LocationURL
    Dim i As Integer
    Dim a As String
    Dim URL As String
    Open App.Path & "\bookmarks.txt" For Output As #1
    For i = 0 To Combo1.ListCount + 1
    Write #1, Combo2.List(i)
    Next i
    Close #1
End Sub
Private Sub PageSetp_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Paste_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Popups_Click()
    If Popups.Checked = False Then
        Popups.Checked = True
        AllowPopups = False
    ElseIf Popups.Checked = True Then
        Popups.Checked = False
        AllowPopups = True
    End If
End Sub

Private Sub Print_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Properties_Click()
    WebBrowser1.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Save_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub SelectAll_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub WebBrowser1_DownloadBegin()
    ccrpProgressBar1.Max = 100
    cmdStop.Enabled = True
End Sub

Private Sub WebBrowser1_DownloadComplete()
    ccrpProgressBar1.Value = 0
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Combo1.Text = WebBrowser1.LocationURL
    Form1.Caption = "Xplorer - " + WebBrowser1.LocationName
    ccrpProgressBar1.Value = 0
    WebBrowser1.Height = Form1.ScaleHeight - 1520
    Combo1.AddItem Combo1.Text
    Dim i As Integer
    Dim a As String
    Open App.Path & "\history.txt" For Output As #1
    For i = 0 To Combo1.ListCount - 1
    Write #1, Combo1.List(i)
    Next i
    Close #1
    cmdGo.Default = True
    cmdStop.Enabled = False
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
    If AllowPopups = True Then
        Dim frmD As New Form1
        Cancel = False
        frmD.Show
    ElseIf AllowPopups = False Then
        Cancel = True
    End If
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    If ProgressMax >= 0 And Progress > 0 And Progress <= ProgressMax Then
        ccrpProgressBar1.Value = Progress / ProgressMax * 100
    End If
    If ccrpProgressBar1.Value = 100 Then ccrpProgressBar1.Value = 0
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    lblStatus.Caption = Text
End Sub

