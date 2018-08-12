VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Mainform 
   BackColor       =   &H008D679B&
   Caption         =   "Job Tracker"
   ClientHeight    =   8175
   ClientLeft      =   1065
   ClientTop       =   2235
   ClientWidth     =   12390
   Icon            =   "Mainform.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrReminder 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2250
      Top             =   6825
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   7890
      Width           =   12390
      _ExtentX        =   21855
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "1/18/01"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13732
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:10 PM"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00997367&
      Height          =   7890
      Left            =   0
      ScaleHeight     =   7830
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   0
      Width           =   1245
      Begin VB.Image imgShowAll 
         Height          =   840
         Left            =   105
         Picture         =   "Mainform.frx":0442
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   900
      End
      Begin VB.Label lblShowAll 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Show All Records"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   525
         Left            =   60
         TabIndex        =   5
         Top             =   6480
         Width           =   1095
      End
      Begin VB.Label lblPreviousQuery 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Saved Queries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   525
         Left            =   30
         TabIndex        =   4
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Image imgPrevious 
         Height          =   840
         Left            =   120
         Picture         =   "Mainform.frx":0884
         Stretch         =   -1  'True
         Top             =   4005
         Width           =   900
      End
      Begin VB.Label lblQuery 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Query Builder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   525
         Left            =   0
         TabIndex        =   2
         Top             =   2985
         Width           =   1095
      End
      Begin VB.Label lblRecruiters 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Tracker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   150
         TabIndex        =   1
         Top             =   1275
         Width           =   825
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgQuery 
         Height          =   585
         Left            =   240
         Picture         =   "Mainform.frx":0B8E
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   705
      End
      Begin VB.Image ImgRecuiter 
         Height          =   840
         Left            =   240
         Picture         =   "Mainform.frx":0FD0
         Stretch         =   -1  'True
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuRecuiters 
         Caption         =   "&Job Tracker"
      End
      Begin VB.Menu mnuQuery 
         Caption         =   "&Query Builder"
      End
      Begin VB.Menu mnuPreviousQueries 
         Caption         =   "&Saved Queries"
      End
      Begin VB.Menu mnuShortcut 
         Caption         =   "&Shortcut Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuStillOpenReport 
         Caption         =   "&Still Open Report"
      End
      Begin VB.Menu mnuAllJobsReport 
         Caption         =   "&All Jobs Report"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuReminder 
         Caption         =   "&Reminder"
      End
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuHorizontally 
         Caption         =   "Tile &Horizontally"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVertically 
         Caption         =   "Tile &Vertically"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************
'RecruiterApp is sort of an address book type app to keep track of
'**Jobs that you apply for that lets you add and delete contacts.
'**Keeps track of jobs that are StillPending (Highlights comments in red)
'**Go to website or send email to your contacts
'Using Access Database
'
'Author: Rick Bales copyright 2000
'rb.sb@gte.net
'Date: 12/21/2000
'**************************************************************************
'New features added on version 2
'query builder enhanced
'saved queries form added
'printing of reports added
'searching with wildcards added
'follow up reminder added
'show all records icon added
'1/18/2001
'***************************************************************************

Option Explicit
Dim WithEvents dc As Adodc
Attribute dc.VB_VarHelpID = -1

Private Sub Image1_Click()

End Sub

Private Sub imgPrevious_Click()
    PreviousQueryForm.Show
End Sub

Private Sub imgPrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblPreviousQuery.ForeColor = RGB(255, 155, 255)
End Sub

Private Sub imgQuery_Click()
    QueryForm.Show
End Sub

Private Sub imgQuery_dblClick()
    QueryForm.Show
End Sub

Private Sub imgQuery_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblQuery.ForeColor = RGB(255, 155, 255)
End Sub

Private Sub ImgRecuiter_Click()
    RecruiterForm.Show
End Sub

Private Sub ImgRecuiter_DblClick()
    RecruiterForm.Show
End Sub


Private Sub ImgRecuiter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     lblRecruiters.ForeColor = RGB(255, 155, 255)
End Sub

Private Sub imgShowAll_Click()

     Dim sqlstring As String
    
    sqlstring = "SELECT * FROM info ORDER by CompanyName"
    dc.RecordSource = sqlstring
    dc.Refresh
    
End Sub

Private Sub imgShowAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblShowAll.ForeColor = RGB(255, 155, 255)
End Sub

Private Sub lblPreviousQuery_Click()
    PreviousQueryForm.Show
End Sub

Private Sub lblPreviousQuery_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblPreviousQuery.ForeColor = RGB(255, 155, 255)
End Sub

Private Sub lblQuery_Click()
    QueryForm.Show
End Sub

Private Sub lblRecuiters_Click()
    RecruiterForm.Show
End Sub

Private Sub lblRecuiters_DblClick()
    RecruiterForm.Show
End Sub

Private Sub lblQuery_DblClick()
    QueryForm.Show
End Sub

Private Sub lblRecruiters_Click()
    RecruiterForm.Show
End Sub

Private Sub lblRecruiters_DblClick()
    RecruiterForm.Show
End Sub

Private Sub lblRecruiters_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblRecruiters.ForeColor = RGB(255, 155, 255)
End Sub

Private Sub lblQuery_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblQuery.ForeColor = RGB(255, 155, 255)
End Sub

Private Sub lblShowAll_Click()
    Call imgShowAll_Click
End Sub

Private Sub lblShowAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblShowAll.ForeColor = RGB(255, 155, 255)
End Sub

Private Sub MDIForm_Load()
    Dim PQToolTip As String 'previous query tool tip
    
    PQToolTip = "Every time that you run a query with the query builder, " & _
        "it is saved so that you can run it again quickly."
    
    lblPreviousQuery.ToolTipText = PQToolTip
    imgPrevious.ToolTipText = PQToolTip
    
    Set dc = RecruiterForm.info
    StatusBar.Panels(2).Text = "Reminder Disabled"
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblQuery.ForeColor = vbWhite
    lblRecruiters.ForeColor = vbWhite
    lblPreviousQuery.ForeColor = vbWhite
    lblShowAll.ForeColor = vbWhite
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
    Unload Me
End Sub

Private Sub MDIForm_Resize()

    'keep forms center if user resizes the mainform
    Dim frm As Form
    For Each frm In Forms
        If frm.Name = "RecruiterForm" Or frm.Name = "QueryForm" Then
            CenterForm frm
            CenterForm frm
        End If
    Next
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    End
End Sub

Private Sub mnuAbout_Click()
    frmSplash.OK.Visible = True
    frmSplash.Show vbModal, Me
End Sub

Private Sub mnuAllJobsReport_Click()
    PrintAllReport.Show vbModal
End Sub

Private Sub mnuArrange_Click()
    Mainform.Arrange vbArrangeIcons
End Sub

Private Sub mnuCascade_Click()
    Mainform.Arrange vbCascade
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuHorizontally_Click()
    Mainform.Arrange vbTileHorizontal
End Sub

Private Sub mnuPreviousQueries_Click()
    PreviousQueryForm.Show
End Sub

Private Sub mnuQuery_Click()
    QueryForm.Show
End Sub

Private Sub mnuRecuiters_Click()
    RecruiterForm.Show
End Sub

Private Sub mnuReminder_Click()
    If mnuReminder.Checked = False Then
        mnuReminder.Checked = True
        tmrReminder.Enabled = True
        StatusBar.Panels(2).Text = "Reminder Enabled"
    Else
        mnuReminder.Checked = False
        tmrReminder.Enabled = False
        StatusBar.Panels(2).Text = "Reminder Disabled"
    End If
End Sub

Private Sub mnuShortcut_Click()
    If mnuShortcut.Checked = True Then
        Picture1.Visible = False
        mnuShortcut.Checked = False
    Else
        Picture1.Visible = True
        mnuShortcut.Checked = True
    End If
End Sub



Private Sub mnuStillOpenReport_Click()
    StillAvailableReport.WindowState = vbMaximized
    StillAvailableReport.Show vbModal
End Sub

Private Sub mnuVertically_Click()
    Mainform.Arrange vbTileVertical
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblRecruiters.ForeColor = vbWhite
    lblQuery.ForeColor = vbWhite
    lblPreviousQuery.ForeColor = vbWhite
    lblShowAll.ForeColor = vbWhite

End Sub

Private Sub tmrReminder_Timer()
    
    Static CheckDate As Date
    Dim sqlstring  As String
    
    
    'The timer interval is initially set to 2 seconds so that if you enable it
    ' it will immediatly check for you, after we don't need to check as often
    tmrReminder.Interval = 65000
    
    'once the reminder is made don't execute this procedure again until tomorrow
    If CheckDate = Date Then Exit Sub
    'select all jobs that need to be contacted today
    sqlstring = "SELECT * From info WHERE NextDate = #" & Date & "#"
    
    'get a recordset of only those jobs to display on Job Tracker
    dc.RecordSource = sqlstring
    dc.Refresh
    
    CheckDate = Date
    If dc.Recordset.RecordCount = 0 Then
        sqlstring = "SELECT * FROM info ORDER by CompanyName"
        dc.RecordSource = sqlstring
        dc.Refresh
        Exit Sub
    End If
    
    MsgBox "You have " & dc.Recordset.RecordCount & " jobs to contact today. Please check the Job Tracker form." _
            & vbNewLine & " The database has been filtered to show only those jobs.", vbInformation
    
    
End Sub
