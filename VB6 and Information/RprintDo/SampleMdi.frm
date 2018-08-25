VERSION 5.00
Object = "{91D3C678-32ED-11D4-8ECF-B7CD5A3EC84A}#1.0#0"; "RoboprintS5R.ocx"
Begin VB.MDIForm SampleMdi 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "SampleMdi"
   ClientHeight    =   5715
   ClientLeft      =   2625
   ClientTop       =   1770
   ClientWidth     =   6795
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox SampleMdi 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   6765
      TabIndex        =   0
      Top             =   0
      Width           =   6795
      Begin VB.CommandButton Command1 
         Height          =   322
         Index           =   0
         Left            =   930
         Picture         =   "SampleMdi.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "DBGRid"
         Top             =   30
         Width           =   345
      End
      Begin VB.CommandButton Command1 
         Height          =   330
         Index           =   6
         Left            =   2805
         Picture         =   "SampleMdi.frx":0386
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Grid Input Edit"
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton Command1 
         Height          =   322
         Index           =   4
         Left            =   3270
         Picture         =   "SampleMdi.frx":07C8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit"
         Top             =   45
         Width           =   345
      End
      Begin VB.CommandButton Command1 
         Height          =   322
         Index           =   3
         Left            =   2320
         Picture         =   "SampleMdi.frx":0C0A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "ListBox"
         Top             =   30
         Width           =   345
      End
      Begin VB.CommandButton Command1 
         Height          =   322
         Index           =   2
         Left            =   1860
         Picture         =   "SampleMdi.frx":0FF4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "MsGraph"
         Top             =   30
         Width           =   345
      End
      Begin VB.CommandButton Command1 
         Height          =   322
         Index           =   1
         Left            =   1395
         Picture         =   "SampleMdi.frx":138F
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "MsFlexGrid"
         Top             =   30
         Width           =   345
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   365
         Left            =   120
         ScaleHeight     =   360
         ScaleMode       =   0  'User
         ScaleWidth      =   633.673
         TabIndex        =   7
         Top             =   0
         Width           =   690
         Begin RoboprintS5R.Roboprint Roboprint1 
            Height          =   375
            Left            =   15
            TabIndex        =   8
            Top             =   0
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   661
            BeginProperty TitlesFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "SampleMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 1
MsFlexSample.Show
MsFlexSample.ZOrder 0
Case 6
Formiga.Show
Formiga.ZOrder 0
Case 0
DBGridSample.Show
DBGridSample.ZOrder 0
Case 2
MsChartFrm.Show
MsChartFrm.ZOrder 0
Case 3
ListFrm.Show
ListFrm.ZOrder 0
Case 5
ListViewSmp.Show
ListViewSmp.ZOrder 0
Case 4
Unload Me
End Select
End Sub

Private Sub Roboprint1_AfterPrint()
Formiga.inputpic.ZOrder 0
End Sub

