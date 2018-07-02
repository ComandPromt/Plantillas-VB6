VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Port Spy"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   4785
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Object.ToolTipText     =   "Hostname of this computer."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Object.ToolTipText     =   "IP of this computer."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   10292
            Object.ToolTipText     =   "Ping requests and replies."
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbr1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   1111
      ButtonWidth     =   1376
      ButtonHeight    =   1058
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgCold"
      DisabledImageList=   "imgCold"
      HotImageList    =   "imgHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "stop"
            Object.ToolTipText     =   "Stop Monitoring"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Object.ToolTipText     =   "Refresh Connection List"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Active"
            Object.ToolTipText     =   "Showing only active connections."
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "IP Addr"
            Object.ToolTipText     =   "Showing IP addresses."
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filter (Off)"
            Object.ToolTipText     =   "Filtering is disabled."
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Log (Off)"
            Object.ToolTipText     =   "Logging is disabled."
            ImageIndex      =   6
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Active Ports"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvwCon"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tmr1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgHot"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgCold"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Filter"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraFilter(0)"
      Tab(1).Control(1)=   "fraFilter(1)"
      Tab(1).Control(2)=   "fraFilter(2)"
      Tab(1).Control(3)=   "chkAct(0)"
      Tab(1).Control(4)=   "chkAct(1)"
      Tab(1).Control(5)=   "chkAct(2)"
      Tab(1).Control(6)=   "cmdSave"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Log"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraLog"
      Tab(2).Control(1)=   "lvwLog"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   -68400
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame fraLog 
         Caption         =   ".:: Logging Options ::."
         Height          =   3375
         Left            =   -69360
         TabIndex        =   18
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton cmdViewLog 
            Caption         =   "View Log"
            Height          =   375
            Left            =   1200
            TabIndex        =   26
            Top             =   2880
            Width           =   975
         End
         Begin VB.CheckBox chkLog 
            Caption         =   "ICMP Statistics"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "Ping requests and replies."
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CheckBox chkLog 
            Caption         =   "Blocked connections"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox chkLog 
            Caption         =   "Connected Ports"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkLog 
            Caption         =   "Listening Ports"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkAct 
         Caption         =   "Actived"
         Height          =   255
         Index           =   2
         Left            =   -70080
         TabIndex        =   12
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkAct 
         Caption         =   "Actived"
         Height          =   255
         Index           =   1
         Left            =   -71760
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkAct 
         Caption         =   "Actived"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   10
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.Frame fraFilter 
         Caption         =   ".:: Local Port ::."
         Height          =   3135
         Index           =   2
         Left            =   -70080
         TabIndex        =   5
         Top             =   600
         Width           =   1575
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   14
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtAdd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   120
            MaxLength       =   5
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
         Begin MSComctlLib.ListView lvwFilter 
            Height          =   2415
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   4260
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame fraFilter 
         Caption         =   ".:: Remote Port ::."
         Height          =   3135
         Index           =   1
         Left            =   -71760
         TabIndex        =   4
         Top             =   600
         Width           =   1575
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   9
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtAdd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   5
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin MSComctlLib.ListView lvwFilter 
            Height          =   2415
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   4260
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame fraFilter 
         Caption         =   ".:: Remote Address ::."
         Height          =   3135
         Index           =   0
         Left            =   -74880
         TabIndex        =   3
         Top             =   600
         Width           =   3015
         Begin MSComctlLib.ListView lvwFilter 
            Height          =   2415
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   4260
            View            =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtAdd 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2055
         End
      End
      Begin MSComctlLib.ImageList imgCold 
         Left            =   840
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   20
         ImageHeight     =   20
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0054
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":05B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0B0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0F60
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":14BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1A18
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgHot 
         Left            =   240
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   20
         ImageHeight     =   20
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1F74
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":24D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2A2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2E80
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":33DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3938
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Timer tmr1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1440
         Top             =   600
      End
      Begin MSComctlLib.ListView lvwCon 
         Height          =   3360
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5927
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Host Address"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Local Port"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Remote Port"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Status"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Time Stamp"
            Object.Width           =   2999
         EndProperty
      End
      Begin MSComctlLib.ListView lvwLog 
         Height          =   3360
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5927
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Host Address"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Local Port"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Remote Port"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Status"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Time Stamp"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.Menu mnuKport 
      Caption         =   "Actions"
      Begin VB.Menu mnuBlock 
         Caption         =   "Block this!"
      End
      Begin VB.Menu mnuHello 
         Caption         =   "Say Hello"
      End
      Begin VB.Menu mS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKill 
         Caption         =   "Terminate Connection"
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "mnuFilter"
      End
      Begin VB.Menu mnuRem 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Copyright © 2001 Gregory Kirk. All rights reserved.
'
' You have a royalty-free right to use, modify, reproduce and distribute the
' Application Files (and/or any modified version) in any way you find useful,
' provided that you agree that Gregory Kirk has no warranty, obligations or
' liability for any Application Files.
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' Main form:
' Note on Blocking web addresses:
'    Blocking addresses does not always work. Some web sites have multiple
'    servers baking up the primary site. If connection to the first server
'    does not work then you will be routed to the second server and so on.
'    To successfully block an address you need to know all the ip addresses
'    that are registered to the primary web address.
'    For a test run I had to add 8 ip addresses to fully block CNN.com from loading in IE 6.0
'
' Note on the connection timer:
'    The default setting is 500 miliseconds.
'    If data from the address or port you are blocking still comes through it is because
'    that data is being received/sent at a rate faster than 500 miliseconds. Experiment with
'    the time interval till you are satisfied with the blocking ability.
'    A time interval of 500 will usually block all data, but, if you have a DSL connection
'    then I suggest setting the interval to 400 or below.
'-------------------------------------------------------------------------------

Option Explicit
Dim StartTmr As String
Public ShowAll As Boolean
Public ResolveAddr As Boolean
Public Filtering As Boolean
Public Logging As Boolean
Dim SaveSuccess As Boolean

Private OnTop As New clsOnTop



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : Form_Load
' Purpose    : Loads settings, activates the tcp table enumeration.
' Parameters : NA
' Return val : NA
' Algorithm  : Calls InitState; LoadSettings; RefreshNS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()


Me.Show
DoEvents
Dim P As Integer
InitStates

'//By setting these to false it will enable the program to load faster.
ShowAll = False
ResolveAddr = False
'\\

LoadSettings
RefreshNS
StartTmr = Now
DoEvents 'Give the cpu a chance to catch its breath
tmr1.Enabled = True 'Turn connection timer on
SaveSuccess = True
SBar.Panels(1).Text = frmCon.WS1.LocalHostName 'Name of this computer
SBar.Panels(2).Text = frmCon.WS1.LocalIP 'Address of this computer
Last_ICMP_Cnt = 0 'Set last ICMP count to 0 to ensure the real count gets logged
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : Form_Unload
' Purpose    : Checks for changes and prompts to save them, writes to the log(PortSpy.log).
' Parameters : Cancel as integer
' Return val : NA
' Algorithm  : Calls FileExist
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
Dim vMsg, i As Integer, tmpStr As String
If SaveSuccess = False Then
    vMsg = MsgBox("Do you want to save setting changes?", vbInformation + vbYesNo, ".:: Save Settings ::.")
    If vMsg = vbYes Then cmdSave_Click
End If
If (FileExist(App.Path & "\PortSpy.Log")) Then 'If log file not found then write new one
    Open App.Path & "\PortSpy.Log" For Append As #1 'Log found, append(add) to it
        For i = 1 To lvwLog.ListItems.Count
            With lvwLog.ListItems(i)
                tmpStr = tmpStr & .Text & " " & .ListSubItems(1).Text & " " & .ListSubItems(2).Text & " " & .ListSubItems(3).Text & " " & .ListSubItems(4).Text & vbCrLf
            End With
        Next
        Print #1, vbCrLf & "<Session :: Start: " & StartTmr & " End: " & Now & ">" & vbCrLf & tmpStr
    Close #1
Else
    Open App.Path & "\PortSpy.Log" For Output As #1 'Log not found, write new log
        For i = 1 To lvwLog.ListItems.Count
            With lvwLog.ListItems(i)
                tmpStr = tmpStr & .Text & " " & .ListSubItems(1).Text & " " & .ListSubItems(2).Text & " " & .ListSubItems(3).Text & " " & .ListSubItems(4).Text & vbCrLf
            End With
        Next
        Print #1, "<<<New Log Started: " & Now & ">>>" & vbCrLf & "[Host address/IP address], [Local Port], [Remote Port], [Connection Status]" & vbCrLf & vbCrLf & "<Session :: Start: " & StartTmr & " End: " & Now & ">" & vbCrLf & tmpStr
    Close #1
End If
'Because SBar is linked to the socket object WS1 on frmCon, in order to close the program
'properly you must force the form, frmCon, to unload even if it is already unloaded.
'I do not know why this happens. Perhaps something to do with the socket being stuck in memory.
Unload frmCon
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : LoadSettings
' Purpose    : Loads settings from configuration file(config.ini).
' Parameters : NA
' Return val : NA
' Algorithm  : Calls ReadINI
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadSettings()
Dim AddrCnt As Integer, RemPcnt As Integer, LocPcnt As Integer, i As Integer

'//Set blocking options
Filtering = ReadINI("Main", "Filter", "Config.ini")
If Filtering = True Then Filter (0) Else Filter (1)
If ReadINI("Address", "Activate", "Config.ini") = "1" Then
    chkAct(0).Value = 1
    chkAct_Click (0)
Else
    chkAct(0).Value = 0
    chkAct_Click (0)
End If
If ReadINI("RemP", "Activate", "Config.ini") = "1" Then
    chkAct(1).Value = 1
    chkAct_Click (1)
Else
    chkAct(1).Value = 0
    chkAct_Click (1)
End If
If ReadINI("LocP", "Activate", "Config.ini") = "1" Then
    chkAct(2).Value = 1
    chkAct_Click (2)
Else
    chkAct(2).Value = 0
    chkAct_Click (2)
End If

'//Add config.ini blocking info to lists
AddrCnt = Val(ReadINI("Address", "AddrCnt", "Config.ini")) '# of address
RemPcnt = Val(ReadINI("RemP", "RemPcnt", "Config.ini")) '# of remote ports
LocPcnt = Val(ReadINI("LocP", "LocPcnt", "Config.ini")) '# of Local ports
For i = 1 To AddrCnt
On Error Resume Next
    lvwFilter(0).ListItems.Add , ReadINI("Address", "ip_" & i, "Config.ini"), ReadINI("Address", "Addr_" & i, "Config.ini")
    If ReadINI("Address", "chk_" & i, "Config.ini") = "True" Then lvwFilter(0).ListItems.Item(i).Checked = True
Next
For i = 1 To RemPcnt
    lvwFilter(1).ListItems.Add , , ReadINI("RemP", "RemP_" & i, "Config.ini")
    If ReadINI("RemP", "chk_" & i, "Config.ini") = "True" Then lvwFilter(1).ListItems.Item(i).Checked = True
Next
For i = 1 To LocPcnt
    lvwFilter(2).ListItems.Add , , ReadINI("LocP", "LocP_" & i, "Config.ini")
    If ReadINI("LocP", "chk_" & i, "Config.ini") = "True" Then lvwFilter(2).ListItems.Item(i).Checked = True
Next

'//Set logging options
Logging = ReadINI("Main", "Log", "Config.ini")
If Logging = True Then Log (0) Else Log (1)
chkLog(0).Value = ReadINI("Log", "Listening", "Config.ini")
chkLog(1).Value = ReadINI("Log", "Connected", "Config.ini")
chkLog(2).Value = ReadINI("Log", "Blocked", "Config.ini")
chkLog(3).Value = ReadINI("Log", "ICMP", "Config.ini")
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : tbr1_ButtonClick
' Purpose    : Handles the buttons that are pressed.
' Parameters : ByVal Button As MSComctlLib.Button
' Return val : NA
' Algorithm  : Calls RefreshNS; Filter(); Log()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbr1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button
    Case "Stop"
        tmr1.Enabled = False
        tbr1.Buttons(1).Caption = "Stopped"
        tbr1.Buttons(1).ToolTipText = "Press to resume monitoring."
        tbr1.Buttons(2).Enabled = False
    Case "Stopped"
        tmr1.Enabled = True
        tbr1.Buttons(1).Caption = "Stop"
        tbr1.Buttons(1).ToolTipText = "Press to stop monitoring."
        tbr1.Buttons(2).Enabled = True
    Case "Refresh": RefreshNS
    Case "Active"
        tbr1.Buttons(4).Caption = "All Ports"
        tbr1.Buttons(4).ToolTipText = "Showing all connections."
        ShowAll = True
        RefreshNS
    Case "All Ports"
        tbr1.Buttons(4).Caption = "Active"
        tbr1.Buttons(4).ToolTipText = "Showing only active connections."
        ShowAll = False
        RefreshNS
    Case "IP Addr"
        tbr1.Buttons(5).Caption = "Resolve"
        tbr1.Buttons(5).ToolTipText = "Showing resolved addresses."
        ResolveAddr = True
        RefreshNS
    Case "Resolve"
        tbr1.Buttons(5).Caption = "IP Addr"
        tbr1.Buttons(5).ToolTipText = "Showing IP addresses."
        ResolveAddr = False
        RefreshNS
    Case "Filter (On)": Filter (1)
    Case "Filter (Off)": Filter (0)
    Case "Log (On)": Log (1)
    Case "Log (Off)": Log (0)
End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : Filter
' Purpose    : To alter settings based on Index.
' Parameters : Index As Integer
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Filter(Index As Integer)
Dim i As Integer
Select Case Index
    Case 0
        tbr1.Buttons(6).Caption = "Filter (On)"
        tbr1.Buttons(6).ToolTipText = "Filtering is enabled."
        tbr1.Buttons(6).Value = tbrPressed
        Filtering = True
        For i = 0 To 2
            chkAct(i).Enabled = True
            chkAct(i).Value = 1
            fraFilter(i).Enabled = True
            txtAdd(i).BackColor = vbWhite
            cmdAdd(i).Enabled = True
            lvwFilter(i).BackColor = vbWhite
            lvwFilter(i).ForeColor = vbBlack
        Next
    Case 1
        tbr1.Buttons(6).Caption = "Filter (Off)"
        tbr1.Buttons(6).ToolTipText = "Filtering is disabled."
        tbr1.Buttons(6).Value = tbrUnpressed
        Filtering = False
        For i = 0 To 2
            chkAct(i).Enabled = False
            chkAct(i).Value = 0
            fraFilter(i).Enabled = False
            txtAdd(i).BackColor = &H8000000F
            cmdAdd(i).Enabled = False
            lvwFilter(i).BackColor = &H8000000F
            lvwFilter(i).ForeColor = &H8000000C
        Next
End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : Log
' Purpose    : To alter settings based on Index.
' Parameters : Index As Integer
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Log(Index As Integer)
Dim i As Integer
Select Case Index
    Case 0
        Logging = True
        tbr1.Buttons(7).Caption = "Log (On)"
        tbr1.Buttons(7).ToolTipText = "Logging is enabled."
        tbr1.Buttons(7).Value = tbrPressed
        lvwLog.BackColor = vbWhite
        lvwLog.Enabled = True
        lvwLog.BackColor = vbWhite
        fraLog.Enabled = True
        For i = 0 To 3
            chkLog(i).Enabled = True
        Next
    Case 1
        Logging = False
        tbr1.Buttons(7).Caption = "Log (Off)"
        tbr1.Buttons(7).ToolTipText = "Logging is disabled."
        tbr1.Buttons(7).Value = tbrUnpressed
        lvwLog.Enabled = False
        lvwLog.BackColor = &H8000000F
        fraLog.Enabled = False
        For i = 0 To 3
            chkLog(i).Enabled = False
        Next
End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : tmr1_Timer
' Purpose    : Connection list timer. Set at 500 miliseconds.
' Parameters : NA
' Return val : NA
' Algorithm  : Calls CheckTcp
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmr1_Timer() '
    CheckTcp
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : RefreshNS
' Purpose    : Refreshes the connection list.
' Parameters : NA
' Return val : NA
' Algorithm  : Calls GetTcpTable; GetAscIP; ntohs; GetHostNameFromIP; rLog; Time
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshNS()
On Error Resume Next
Dim Item As ListItem
Dim LTmp As Long, State As Long, Val As Long
Dim X As Integer, i As Integer, n As Integer
Dim rHost As String, LocP As String
Dim tcpt As MIB_TCPTABLE

LTmp = Len(MIB_TCPTABLE)
GetTcpTable tcpt, LTmp, 1
lvwCon.ListItems.Clear

For i = 0 To tcpt.dwNumEntries - 1
    State = tcpt.table(i).dwState
    If ((State <> 2) And (ShowAll = False)) Or (ShowAll = True) Then
        rHost = GetAscIP(tcpt.table(i).dwRemoteAddr)
        LocP = ntohs(tcpt.table(i).dwLocalPort) 'Retrieve the actual IP
        Set Item = lvwCon.ListItems.Add()
        If (State <> 2) Then 'If not listening then...
            If ResolveAddr = True Then
                rHost = GetHostNameFromIP(rHost) 'Retrieve host name
            End If
            Item.Text = rHost 'Host Address (IP or Name alias)
            Item.Tag = i 'TCP table number
            Item.SubItems(1) = LocP 'Local Port
            Item.SubItems(2) = ntohs(tcpt.table(i).dwRemotePort) 'Remote Port, Retrieve the actual IP
            Item.SubItems(3) = IP_States(State) 'Connection Status
            Item.SubItems(4) = Time
            If Logging = True And chkLog(1).Value = 1 Then rLog Item.Text, LocP, ntohs(tcpt.table(i).dwRemotePort), IP_States(State), Time
        Else
            Item.Text = "[ Listening For Connection ]"
            Item.Tag = i 'TCP table number
            Item.SubItems(1) = LocP 'Local Port
            Item.SubItems(2) = "n\a" 'no remote port detected.
            Item.SubItems(3) = IP_States(State) 'Connection Status
            Item.SubItems(4) = Time
            If Logging = True And chkLog(0).Value = 1 Then rLog Item.Text, LocP, "n\a", IP_States(State), Time
        End If
    End If
Next i
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : lvwCon_MouseUp
' Purpose    : Displays the list view(lvwCon) menu.
' Parameters : Button As Integer, Shift As Integer, x As Single, Y As Single
' Return val : NA
' Algorithm  : Calls mnuKport
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvwCon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And lvwCon.ListItems.Count > 0 Then
    frmMain.PopupMenu mnuKport
End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : mnuKill_Click
' Purpose    : Terminates the selected established connection.
' Parameters : NA
' Return val : NA
' Algorithm  : Calls GetTcpTable; SetTcpEntry
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuKill_Click() '
Dim LTmp As Long, tg As String
Dim tcpt As MIB_TCPTABLE
LTmp = Len(MIB_TCPTABLE)
GetTcpTable tcpt, LTmp, 1
tg = lvwCon.SelectedItem.Tag
If tcpt.table(tg).dwState <> 2 Then
    tcpt.table(tg).dwState = 12
    SetTcpEntry tcpt.table(tg)
Else
    MsgBox "Currently unable to terminate listening ports..." 'I would love to know how to do this!
End If
DoEvents
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : mnuHello_click
' Purpose    : Connecting attempt to selected address via an HTTP header or user text.
' Parameters : NA
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuHello_click()
Dim adr As String, prt As Integer
frmCon.txtAddr = lvwCon.SelectedItem.Text  'Host Address
frmCon.txtPort = lvwCon.SelectedItem.ListSubItems(2) 'Remote Port
frmCon.Show
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : cmdViewLog_Click
' Purpose    : Displays the Port Spy Log.
' Parameters : NA
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdViewLog_Click()
frmView.txtConView.LoadFile App.Path & "\PortSpy.log"
frmView.Show
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : mnuBlock_Click
' Purpose    : Adds an entry to be blocked.
' Parameters : NA
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuBlock_Click()
With frmBlock
    .txtAddr = lvwCon.SelectedItem.Text
    .txtRemP = lvwCon.SelectedItem.ListSubItems(2)
    .txtLocP = lvwCon.SelectedItem.ListSubItems(1)
    .Show
End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : chkAct_Click
' Purpose    : Activate/de-activate selected blocking feature based upon Index.
' Parameters : Index As Integer
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkAct_Click(Index As Integer)
SaveSuccess = False
If chkAct(Index).Value = 1 Then
    fraFilter(Index).Enabled = True
    txtAdd(Index).BackColor = vbWhite
    cmdAdd(Index).Enabled = True
    lvwFilter(Index).BackColor = vbWhite
    lvwFilter(Index).ForeColor = vbBlack
Else
    fraFilter(Index).Enabled = False
    txtAdd(Index).BackColor = &H8000000F
    cmdAdd(Index).Enabled = False
    lvwFilter(Index).BackColor = &H8000000F
    lvwFilter(Index).ForeColor = &H8000000C
End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : cmdAdd_Click
' Purpose    : Add new entry to Filter list based on Index.
' Parameters : Index As Integer
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub cmdAdd_Click(Index As Integer)
If txtAdd(Index).Text = "" Then Exit Sub

Dim itmNum As Integer, i As Integer, tmp As Integer
Dim rHost As String
SaveSuccess = False

itmNum = lvwFilter(Index).ListItems.Count
For i = 1 To itmNum
    If txtAdd(Index).Text = lvwFilter(Index).ListItems.Item(i).Text Then
        MsgBox "Item allready exists..."
        txtAdd(Index).Text = ""
        Exit Sub
    End If
Next i
Select Case Index
    Case 0 'Validate address
        rHost = GetIPFromHostName(txtAdd(Index).Text) 'Return IP of address
        For i = 1 To itmNum
            If GetAscIP(rHost) = lvwFilter(Index).ListItems.Item(i).Key Then
                MsgBox "Ip address allready exists..."
                txtAdd(Index).Text = ""
            Exit Sub
            End If
        Next i
        lvwFilter(Index).ListItems.Add , GetAscIP(rHost), txtAdd(Index).Text 'Set item Key to IP, set text as host name
    Case 1, 2 'Validate port number
            For i = 1 To Len(txtAdd(Index).Text)
                tmp = Asc(Mid(txtAdd(Index).Text, i, 1))
                If tmp < 48 Or tmp > 57 Then
                    MsgBox "Not a valid number..."
                    txtAdd(Index).Text = ""
                    Exit Sub
                End If
            Next i
        lvwFilter(Index).ListItems.Add , , txtAdd(Index).Text
End Select
txtAdd(Index).Text = ""
lvwFilter(Index).ListItems.Item(itmNum + 1).Checked = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : cmdAdd_Click
' Purpose    : Add new entry to Filter list based on Index.
' Parameters : Index As Integer
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvwFilter_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And lvwFilter(Index).ListItems.Count > 0 Then
    mnuRem.Tag = Index
    On Error Resume Next
    frmMain.PopupMenu mnuFilter
End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : mnuRem_Click
' Purpose    : Removes selected item from list
' Parameters : NA
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRem_Click() '
SaveSuccess = False
lvwFilter(mnuRem.Tag).ListItems.Remove (lvwFilter(mnuRem.Tag).SelectedItem.Index)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : cmdSave_Click
' Purpose    : Saves settings to configuration file(config.ini).
' Parameters : NA
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSave_Click()
Dim str As String, i As Integer

'//Main Settings
str = "[Main]" & vbCrLf
str = str & "Filter=" & Filtering & vbCrLf
str = str & "Log=" & Logging & vbCrLf

'//Address Settings
str = str & vbCrLf
str = str & "[Address]" & vbCrLf
str = str & "Activate=" & chkAct(0).Value & vbCrLf
str = str & "AddrCnt=" & lvwFilter(0).ListItems.Count & vbCrLf
For i = 1 To lvwFilter(0).ListItems.Count
    str = str & "Addr_" & i & "=" & lvwFilter(0).ListItems(i).Text & vbCrLf
    str = str & "ip_" & i & "=" & lvwFilter(0).ListItems(i).Key & vbCrLf
    str = str & "chk_" & i & "=" & lvwFilter(0).ListItems(i).Checked & vbCrLf
Next

'//Remote Port Settings
str = str & vbCrLf
str = str & "[RemP]" & vbCrLf
str = str & "Activate=" & chkAct(1).Value & vbCrLf
str = str & "RemPcnt=" & lvwFilter(1).ListItems.Count & vbCrLf
For i = 1 To lvwFilter(1).ListItems.Count
    str = str & "RemP_" & i & "=" & lvwFilter(1).ListItems(i).Text & vbCrLf
    str = str & "chk_" & i & "=" & lvwFilter(1).ListItems(i).Checked & vbCrLf
Next

'//Local Port Settings
str = str & vbCrLf
str = str & "[LocP]" & vbCrLf
str = str & "Activate=" & chkAct(2).Value & vbCrLf
str = str & "LocPcnt=" & lvwFilter(2).ListItems.Count & vbCrLf
For i = 1 To lvwFilter(2).ListItems.Count
    str = str & "LocP_" & i & "=" & lvwFilter(2).ListItems(i).Text & vbCrLf
    str = str & "chk_" & i & "=" & lvwFilter(2).ListItems(i).Checked & vbCrLf
Next

'//Log Settings
str = str & vbCrLf
str = str & "[Log]" & vbCrLf
str = str & "Listening=" & chkLog(0).Value & vbCrLf
str = str & "Connected=" & chkLog(1).Value & vbCrLf
str = str & "Blocked=" & chkLog(2).Value & vbCrLf
str = str & "ICMP=" & chkLog(3).Value & vbCrLf

'It is a lot faster to over-write the config.ini file then calling ini functions multiple times.
Open App.Path & "\Config.ini" For Output As #1
    Print #1, str
Close #1
SaveSuccess = True
End Sub
