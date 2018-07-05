VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{70395669-C294-4A5A-A7F3-EDA4A19841E1}#2.1#0"; "KRScroll.ocx"
Begin VB.Form frmMain 
   Caption         =   "Kath-Rock - Image Viewer"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7005
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   361
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   Begin MSComctlLib.ImageList imlTabs 
      Left            =   1860
      Top             =   3885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   14
      ImageHeight     =   15
      MaskColor       =   -2147483633
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1042
            Key             =   "On"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":136E
            Key             =   "Off"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolsDisabled 
      Left            =   1140
      Top             =   3885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":169A
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2376
            Key             =   "Views"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2692
            Key             =   "AddTabs"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":29AE
            Key             =   "DelTab"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2CD2
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":39AE
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolsHover 
      Left            =   570
      Top             =   3885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3CCA
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":49A6
            Key             =   "Views"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4CC2
            Key             =   "AddTabs"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4FDE
            Key             =   "DelTab"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5302
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5FDE
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrPos 
      Interval        =   200
      Left            =   2655
      Top             =   4020
   End
   Begin VB.PictureBox picFrame 
      Height          =   1965
      Left            =   2925
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1110
      Width           =   1425
      Begin vbpKRScroll.KRScroll krsHorz 
         Height          =   195
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1710
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   344
         AutoSystemResize=   -1  'True
         Orientation     =   1
         OverlapPosition =   2
         OverlapTwips    =   195
         BorderStyle     =   0
         FocusRect       =   0   'False
         MouseWheel      =   0
      End
      Begin vbpKRScroll.KRScroll krsVert 
         Height          =   1725
         Left            =   1170
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   -15
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   3043
         AutoSystemResize=   -1  'True
         BorderStyle     =   0
         FocusRect       =   0   'False
         MouseWheel      =   0
      End
      Begin VB.PictureBox picView 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   6
         Top             =   0
         Width           =   780
      End
   End
   Begin MSComctlLib.TabStrip tbsMain 
      Height          =   3150
      Left            =   2835
      TabIndex        =   4
      Top             =   720
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   5556
      MultiRow        =   -1  'True
      TabFixedHeight  =   503
      TabMinWidth     =   1005
      ImageList       =   "imlTabs"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   3195
      Left            =   0
      TabIndex        =   1
      Top             =   690
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   5636
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modified"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Attributes"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.FileListBox filFiles 
      Height          =   285
      Left            =   1950
      Pattern         =   "*.bmp;*.dib;*.rle;*.gif;*.jpg;*.jpe;*.wmf;*.emf;*.ico;*.cur"
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4650
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   0
      Top             =   3885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":62FA
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":6FD6
            Key             =   "Views"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":72F2
            Key             =   "AddTabs"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":760E
            Key             =   "DelTab"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":7932
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":860E
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlTools"
      DisabledImageList=   "imlToolsDisabled"
      HotImageList    =   "imlToolsHover"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Folder"
            Object.ToolTipText     =   "Select Folder"
            ImageKey        =   "Folder"
            Style           =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Views"
            Object.ToolTipText     =   "Change View"
            ImageKey        =   "Views"
            Style           =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AddTabs"
            Object.ToolTipText     =   "Add Tab(s) from selected Image(s)"
            ImageKey        =   "AddTabs"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DelTab"
            Object.ToolTipText     =   "Delete Selected Tab"
            ImageKey        =   "DelTab"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Edit"
            Object.ToolTipText     =   "Edit Current Image"
            ImageKey        =   "Edit"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit Program"
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolsSm 
      Left            =   0
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":892A
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":8CC4
            Key             =   "Views"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":8E1E
            Key             =   "AddTabs"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":8F78
            Key             =   "DelTab"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":90D2
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":946C
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolsHoverSm 
      Left            =   570
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":95C6
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":9960
            Key             =   "View"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":9ABA
            Key             =   "AddTabs"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":9C14
            Key             =   "DelTab"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":9D6E
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A108
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolsDisabledSm 
      Left            =   1140
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A262
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A5FC
            Key             =   "View"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A756
            Key             =   "AddTabs"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A8B0
            Key             =   "DelTab"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":AA0A
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":ADA4
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplit 
      Height          =   3165
      Left            =   2670
      MousePointer    =   9  'Size W E
      Top             =   705
      Width           =   75
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_Select 
         Caption         =   "&Select Folder"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "&Recent Folders"
         Begin VB.Menu mnuRecent_Logging 
            Caption         =   "&Log Recent Folders"
         End
         Begin VB.Menu mnuRecent_Clear 
            Caption         =   "&Clear Recent Folders Log"
         End
         Begin VB.Menu mnuRecent_Line10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecent_Folders 
            Caption         =   "(Empty)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuFile_Line10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Edit 
         Caption         =   "&Edit Current Image"
      End
      Begin VB.Menu mnuFile_Line20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuViews 
      Caption         =   "&View"
      Begin VB.Menu mnuViews_FileList 
         Caption         =   "&File List"
         Begin VB.Menu mnuViews_View 
            Caption         =   "Lar&ge Icons"
            Index           =   0
         End
         Begin VB.Menu mnuViews_View 
            Caption         =   "S&mall Icons"
            Index           =   1
         End
         Begin VB.Menu mnuViews_View 
            Caption         =   "&List"
            Index           =   2
         End
         Begin VB.Menu mnuViews_View 
            Caption         =   "&Details"
            Index           =   3
         End
      End
      Begin VB.Menu mnuViews_ToolView 
         Caption         =   "&Toolbar"
         Begin VB.Menu mnuToolView_Sub 
            Caption         =   "&Hide"
            Index           =   0
         End
         Begin VB.Menu mnuToolView_Sub 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuToolView_Sub 
            Caption         =   "Show &Text"
            Index           =   2
         End
         Begin VB.Menu mnuToolView_Sub 
            Caption         =   "Align &Bottom"
            Index           =   3
         End
         Begin VB.Menu mnuToolView_Sub 
            Caption         =   "Align &Right"
            Index           =   4
         End
         Begin VB.Menu mnuToolView_Sub 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuToolView_Sub 
            Caption         =   "&Large Icons"
            Index           =   6
         End
         Begin VB.Menu mnuToolView_Sub 
            Caption         =   "&Small Icons"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuTabs 
      Caption         =   "&Tabs"
      Begin VB.Menu mnuTabs_Sub 
         Caption         =   "&Add Tab(s)"
         Index           =   0
      End
      Begin VB.Menu mnuTabs_Sub 
         Caption         =   "&Remove Tab"
         Index           =   1
      End
      Begin VB.Menu mnuTabs_Sub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTabs_Sub 
         Caption         =   "&Multiple Rows"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ImageSizeConstants
    Large
    Small
End Enum

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type ImageSpecs
    Filename    As String
    HMax        As Long
    HPos        As Long
    VMax        As Long
    VPos        As Long
End Type

Private Type ToolbarCapsKeys
    Caption As String
    Key     As Variant
End Type

Private Type ToolbarSpecs
    ShowText    As Boolean
    TextAlign   As ToolbarTextAlignConstants
    ImageSize   As ImageSizeConstants
End Type

Private mbNoScroll      As Boolean
Private miIconSize      As Integer
Private mlMinWidth      As Long
Private mlMinHeight     As Long
Private mlMinSplit      As Long
Private mlMaxSplit      As Long
Private mlDragOffset    As Long
Private mlItemIndex     As Long
Private mImages()       As ImageSpecs
Private mToolCapKey()   As ToolbarCapsKeys
Private mToolSpecs      As ToolbarSpecs

Private Const MAX_TABS      As Integer = 200
Private Const SW_RESTORE    As Long = &H9&
Private Const DI_NORMAL     As Long = &H3&

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long

Private Sub Browse()

Dim bLogIt  As Integer
Dim iIdx    As Integer
Dim sFolder As String

    'Browse for a new folder
    sFolder = BrowseForFolder(Me.hWnd, "Select an Image Folder", filFiles.Path, CSIDL_DESKTOP, True)
    If Len(sFolder) > 0 Then
        mlItemIndex = -1
        filFiles.Path = sFolder
        'Add the folder to the recent list, if needed
        sFolder = sFolder & IIf(Right$(sFolder, 1) <> "\", "\", "")
        For iIdx = 0 To mnuRecent_Folders.UBound
            mnuRecent_Folders(iIdx).Checked = False
            If LCase$(mnuRecent_Folders(iIdx).Tag) = LCase$(sFolder) Then
                bLogIt = True
            End If
        Next
        If bLogIt Or mnuRecent_Logging.Checked Then
            Call AddRecentFile(sFolder, mnuRecent_Folders)
        End If
        'Fill the file list and force an update for the preview
        Call FillFileList(filFiles, lvwFiles)
        Call lvwFiles_ItemClick(lvwFiles.SelectedItem)
    End If
    
    Call EnableTools
    
End Sub

Private Sub DelTab()

Dim iIdx As Integer

    'Remove the current tab
    If tbsMain.SelectedItem.Index > 1 Then
        mbNoScroll = True
        For iIdx = tbsMain.SelectedItem.Index To tbsMain.Tabs.Count - 1
            mImages(iIdx) = mImages(iIdx + 1)
        Next
        iIdx = tbsMain.SelectedItem.Index
        tbsMain.Tabs.Remove iIdx
        ReDim Preserve mImages(tbsMain.Tabs.Count)
        
        If iIdx <= tbsMain.Tabs.Count Then
            Set tbsMain.SelectedItem = tbsMain.Tabs(iIdx)
        Else
            Set tbsMain.SelectedItem = tbsMain.Tabs(iIdx - 1)
        End If
        mbNoScroll = False
        Call EnableTools
    
        picFrame.Move tbsMain.ClientLeft + 2, tbsMain.ClientTop + 4, tbsMain.ClientWidth - 4, tbsMain.ClientHeight - 6
        
        'Force an update
        picFrame.ZOrder
        Call picFrame_Resize
    End If
    
End Sub

Private Function DragDrop(Data As DataObject) As Long

Dim lIdx    As Long
Dim lPos    As Long
Dim sFile   As String

    'Add a new tab if an image was dropped on the frame or view
    If Data.GetFormat(vbCFFiles) Then
        For lIdx = 1 To Data.Files.Count
            If InStr(1, filFiles.Pattern, "*" & Right$(Data.Files(lIdx), 4)) > 0 Then
                Call AddTab(Data.Files(lIdx))
            End If
        Next
    End If

    DragDrop = vbDropEffectNone

End Function

Private Function DragOver(Data As DataObject) As Long

Dim lIdx    As Long
Dim lEffect As Long

    'Show drop effect image is being dragged over the frame or view
    lEffect = vbDropEffectNone
    If Data.GetFormat(vbCFFiles) Then
        For lIdx = 1 To Data.Files.Count
            If InStr(1, filFiles.Pattern, "*" & Right$(Data.Files(lIdx), 4)) > 0 Then
                lEffect = vbDropEffectCopy
                Exit For
            End If
        Next
    End If

    DragOver = lEffect
    
End Function


Private Sub EnableTools()

    'Enable/disable the tools as neccessary
    tbrMain.Buttons("Edit").Enabled = (Len(mImages(tbsMain.SelectedItem.Index).Filename) > 0)
    mnuFile_Edit.Enabled = tbrMain.Buttons("Edit").Enabled
    mnuTabs_Sub(0).Enabled = (tbsMain.Tabs.Count < MAX_TABS) And (lvwFiles.ListItems.Count > 0)
    tbrMain.Buttons("AddTabs").Enabled = mnuTabs_Sub(0).Enabled
    mnuTabs_Sub(1).Enabled = (tbsMain.SelectedItem.Index > 1)
    tbrMain.Buttons("DelTab").Enabled = mnuTabs_Sub(1).Enabled
    mnuRecent_Clear.Enabled = mnuRecent_Folders(0).Enabled
    
End Sub

Private Sub ExtractCursor(ByVal sPath As String, picDraw As PictureBox)

'This sub is used to draw a cursor in color, since VB's
'LoadPicture always converts cursors to monochrome.

Dim lRet    As Long
Dim hIcon   As Long

    hIcon = ExtractAssociatedIcon(App.hInstance, sPath, &H0&)
    If hIcon <> 0 Then
        picDraw.Width = 34
        picDraw.Height = 34
        Set picDraw.Picture = LoadPicture()
        picDraw.Cls
        lRet = DrawIconEx(picDraw.hDC, &H0&, &H0&, hIcon, &H20&, &H20&, &H0&, &H0&, DI_NORMAL)
        lRet = DestroyIcon(hIcon)
        picDraw.Refresh
    End If
    
End Sub

Private Sub AddTab(ByVal sFilename As String)

Dim iIdx    As Integer
Dim lPos    As Long
Dim sTitle  As String
Dim oTab    As Object

    'Find out if that tab already exists
    For iIdx = 2 To tbsMain.Tabs.Count
        If LCase$(mImages(iIdx).Filename) = LCase$(sFilename) Then
            'Tab already exists; Set it and get out
            Set tbsMain.SelectedItem = tbsMain.Tabs(iIdx)
            Exit Sub
        End If
    Next
    
    'Add a new tab and set it's caption
    iIdx = tbsMain.Tabs.Count + 1
    lPos = InStrRev(sFilename, "\")
    If lPos > 0 And lPos < Len(sFilename) Then
        sTitle = Mid$(sFilename, lPos + 1)
    End If
    Set oTab = tbsMain.Tabs.Add(, , sTitle)
    
    'Load the image for the new tab
    ReDim Preserve mImages(iIdx)
    With mImages(iIdx)
        .Filename = sFilename
        .HPos = 0
        .VPos = 0
    End With
    
    'Make the new tab the selected one
    Set tbsMain.SelectedItem = oTab
    Call EnableTools
    
    'Move the frame and view into position
    picView.Move 0, 0
    picFrame.Move tbsMain.ClientLeft + 2, tbsMain.ClientTop + 4, tbsMain.ClientWidth - 4, tbsMain.ClientHeight - 6
    
    'Force a resize to setup the scroll bars, etc.
    picFrame.ZOrder
    Call picFrame_Resize
    
    DoEvents
    
End Sub

Private Sub LoadImage()

Dim sPath As String

    Screen.MousePointer = vbHourglass
    sPath = mImages(tbsMain.SelectedItem.Index).Filename
    
    If InStr(1, LCase$(sPath), ".ico") > 0 Then
        'Load the large icon
        Set picView.Picture = LoadPicture(sPath, vbLPLargeShell, vbLPColor)
    ElseIf InStr(1, LCase$(sPath), ".cur") > 0 Then
        'Load the cursor in color
        Call ExtractCursor(sPath, picView)
    Else
        'Load the image normally
        Set picView.Picture = LoadPicture(sPath)
    End If
    picView.Visible = True
    sbrMain.Panels(2).Text = "Size: " & CStr(picView.ScaleWidth) & " x " & CStr(picView.ScaleHeight) & " pixels"
    
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub SetupToolbar()

Dim iIdx    As Integer
    
    'Setup toolbar captions, images, text and text alignment
    Set tbrMain.ImageList = Nothing
    Set tbrMain.HotImageList = Nothing
    Set tbrMain.DisabledImageList = Nothing
    If mnuToolView_Sub(7).Checked Then
        Set tbrMain.ImageList = imlToolsSm
        Set tbrMain.HotImageList = imlToolsHoverSm
        Set tbrMain.DisabledImageList = imlToolsDisabledSm
    Else
        Set tbrMain.ImageList = imlTools
        Set tbrMain.HotImageList = imlToolsHover
        Set tbrMain.DisabledImageList = imlToolsDisabled
    End If
    
    If Not mnuToolView_Sub(2).Checked Then
        mnuToolView_Sub(3).Checked = False
        mnuToolView_Sub(4).Checked = False
        mnuToolView_Sub(3).Enabled = False
        mnuToolView_Sub(4).Enabled = False
    Else
        mnuToolView_Sub(3).Checked = CBool(GetInitEntry("Toolbar", "Text Align", "0") = 0)
        mnuToolView_Sub(4).Checked = Not mnuToolView_Sub(3).Checked
        mnuToolView_Sub(3).Enabled = True
        mnuToolView_Sub(4).Enabled = True
    End If
    
    tbrMain.TextAlignment = IIf(mnuToolView_Sub(4).Checked, tbrTextAlignRight, tbrTextAlignBottom)
    For iIdx = 1 To 6
        tbrMain.Buttons(mToolCapKey(iIdx).Key).Caption = IIf(mnuToolView_Sub(2).Checked, mToolCapKey(iIdx).Caption, "")
        tbrMain.Buttons(mToolCapKey(iIdx).Key).Image = mToolCapKey(iIdx).Key
    Next

    Call Form_Resize
    
End Sub

Private Sub filFiles_PathChange()

Dim lRet As Long

    'Show the current path in the status bar
    sbrMain.Panels(1).Text = filFiles.Path & IIf(Right$(filFiles.Path, 1) <> "\", "\", "")
    
End Sub


Private Sub Form_Load()

Dim bMaxed  As Boolean
Dim iIdx    As Integer
Dim lLeft   As Long
Dim lTop    As Long
Dim lWidth  As Long
Dim lHeight As Long
Dim sPath   As String
Dim aCapKey As Variant

    'Setup the mins, maxes, etc.
    mlMinWidth = 3000
    mlMinHeight = 3000
    mlMinSplit = 100
    mlItemIndex = -1
    ReDim mImages(1)
    
    'Retrieve the form's saved positions
    lLeft = Val(GetInitEntry("Positions", "Left", CStr(Me.Left)))
    lTop = Val(GetInitEntry("Positions", "Top", CStr(Me.Top)))
    lWidth = Val(GetInitEntry("Positions", "Width", CStr(Me.Width)))
    lHeight = Val(GetInitEntry("Positions", "Height", CStr(Me.Height)))
    
    'Test positions against screen resolution
    If lWidth < 0 Then
        lWidth = 0
    ElseIf lWidth > Screen.Width Then
        lWidth = Screen.Width
    End If
    If lLeft < 0 Then
        lLeft = 0
    ElseIf lLeft > Screen.Width - lWidth Then
        lLeft = Screen.Width - lWidth
    End If
    If lHeight < 0 Then
        lHeight = 0
    ElseIf lHeight > Screen.Height Then
        lHeight = Screen.Height
    End If
    If lTop < 0 Then
        lTop = 0
    ElseIf lTop > Screen.Height - lHeight Then
        lTop = Screen.Height - lHeight
    End If
    
    'Position the form
    Me.Move lLeft, lTop, lWidth, lHeight
    
    'Retrieve the tool view menu settings
    If CBool(GetInitEntry("Toolbar", "Visible", "True")) Then
        mnuToolView_Sub(0).Caption = "&Hide"
        tbrMain.Visible = True
        lTop = tbrMain.Height
    Else
        mnuToolView_Sub(0).Caption = "&Show"
        tbrMain.Visible = False
        lTop = 0
    End If
    mnuToolView_Sub(2).Checked = CBool(GetInitEntry("Toolbar", "Show Text", "False"))
    mnuToolView_Sub(3).Checked = CBool(GetInitEntry("Toolbar", "Text Align", "0") = 0)
    mnuToolView_Sub(4).Checked = Not mnuToolView_Sub(3).Checked
    mnuToolView_Sub(6).Checked = CBool(GetInitEntry("Toolbar", "Icon Size", "0") = 0)
    mnuToolView_Sub(7).Checked = Not mnuToolView_Sub(6).Checked
    
    'Retrieve the saved split position
    lvwFiles.Width = Val(GetInitEntry("Positions", "Split", CStr(lvwFiles.Width)))
    lvwFiles.Move 0, lTop + 4, lvwFiles.Width, Me.ScaleHeight - (lTop + 4) - sbrMain.Height
    imgSplit.Move lvwFiles.Width - 1, lvwFiles.Top, imgSplit.Width, lvwFiles.Height
    
    'Maximize the form if that's how it was previously
    If GetInitEntry("Positions", "Maximized", CStr(False)) Then
        Me.WindowState = vbMaximized
    End If
    
    tbsMain.MultiRow = CBool(GetInitEntry("Main", "MultiRow Tabs", "False"))
    mnuTabs_Sub(3).Checked = tbsMain.MultiRow
    
    'Setup toolbar captions, images, text and text alignment
    aCapKey = Array(Array("Folder", "View", "Add Tab", "Del Tab", "Edit", "Exit"), _
            Array("Folder", "Views", "AddTabs", "DelTab", "Edit", "Exit"))
    ReDim mToolCapKey(1 To 6)
    For iIdx = 1 To 6
        mToolCapKey(iIdx).Caption = aCapKey(0)(iIdx - 1)
        mToolCapKey(iIdx).Key = aCapKey(1)(iIdx - 1)
    Next
    mnuToolView_Sub(2).Checked = Not mnuToolView_Sub(2).Checked
    Call mnuToolView_Sub_Click(2)
    
    'Show the form before filling the ListView
    Me.Show
    Me.Refresh
    DoEvents
    
    'Retrieve the recent files list
    Call GetRecentFiles(mnuRecent_Folders)
    mnuRecent_Folders(0).Checked = mnuRecent_Folders(0).Enabled
    mnuRecent_Logging.Checked = CBool(GetInitEntry("Main", "Logging", "On") = "On")
    
    'Initialize the ListView
    Set lvwFiles.SmallIcons = imlTabs
    lvwFiles.ListItems.Add , , "initializing..."
    DoEvents
    lvwFiles.ListItems.Remove 1
    Set lvwFiles.SmallIcons = Nothing
    
    'Fill the ListView
    sPath = GetInitEntry("Recent Files", "File 1", App.Path)
    If Not Exists(sPath) Then
        sPath = App.Path
    End If
    filFiles.Path = sPath
    lvwFiles.View = lvwIcon
    Call mnuViews_View_Click(Val(GetInitEntry("Main", "Last View", CStr(3))))
    lvwFiles.Visible = False
    lvwFiles.Visible = True
    lvwFiles.SetFocus
    
    'Show the selected image
    tbsMain.Tabs(1).Caption = "Preview"
    Call lvwFiles_ItemClick(lvwFiles.SelectedItem)
    
    krsVert.AttachWindowToWheel picView.hWnd
    
End Sub

Private Sub Form_Paint()

Dim lTop    As Long
Static bBusy As Boolean

    If Not bBusy Then
        bBusy = True
        If tbrMain.Visible Then
            lTop = tbrMain.Height
        Else
            lTop = 0
        End If
        Me.Line (0, lTop)-Step(Me.ScaleWidth, 0), vb3DShadow
        Me.Line (0, lTop + 1)-Step(Me.ScaleWidth, 0), vb3DHighlight
        bBusy = False
    End If
    
End Sub

Private Sub Form_Resize()

Dim iIdx    As Integer
Dim lTop    As Long

    If Me.WindowState <> vbMinimized Then
        'Check the minimum width and height
        If Me.Width < mlMinWidth Then
            Me.Width = mlMinWidth
        ElseIf Me.Height < mlMinHeight Then
            Me.Height = mlMinHeight
        Else
            DoEvents
            If tbrMain.Visible Then
                lTop = tbrMain.Height
            Else
                lTop = 0
            End If
            
            'Setup the max split position and change
            'the width of the ListView, if needed.
            mlMaxSplit = Me.ScaleWidth - 100
            If lvwFiles.Width > mlMaxSplit Then
                lvwFiles.Width = mlMaxSplit + 1
            End If
            
            'Draw embossed line to separate the toolbar from the rest of the form
            Me.Line (0, lTop)-Step(Me.ScaleWidth, 0), vb3DShadow
            Me.Line (0, lTop + 1)-Step(Me.ScaleWidth, 0), vb3DHighlight
            
            'Set the height of the ListView
            lvwFiles.Top = lTop + 4
            lvwFiles.Height = Me.ScaleHeight - (lTop + 4) - sbrMain.Height
            
            'Set the height of the splitter
            imgSplit.Move lvwFiles.Width - 1, lvwFiles.Top, imgSplit.Width, lvwFiles.Height
            
            'Resize the Tab control and picture frames
            tbsMain.Move lvwFiles.Width + 5, lvwFiles.Top, Me.ScaleWidth - lvwFiles.Width - imgSplit.Width, lvwFiles.Height
            If (tbsMain.ClientHeight - krsHorz.Height - 12) > 0 Then
                picFrame.Visible = True
                picFrame.Move tbsMain.ClientLeft + 2, tbsMain.ClientTop + 4, tbsMain.ClientWidth - 4, tbsMain.ClientHeight - 6
            Else
                picFrame.Visible = False
            End If
        End If
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    'Save the current splitter position
    Call SetInitEntry("Positions", "Split", CStr(lvwFiles.Width))
    
    'Save the current form position
    If Me.WindowState = vbNormal Then
        Call SetInitEntry("Positions", "Left", CStr(Me.Left))
        Call SetInitEntry("Positions", "Top", CStr(Me.Top))
        Call SetInitEntry("Positions", "Width", CStr(Me.Width))
        Call SetInitEntry("Positions", "Height", CStr(Me.Height))
    End If
    Call SetInitEntry("Positions", "Maximized", CStr(Me.WindowState = vbMaximized))
    
End Sub


Private Sub krsHorz_Change()

    If Not mbNoScroll Then
        'Move the image
        picView.Left = -krsHorz.Value
        mImages(tbsMain.SelectedItem.Index).HPos = krsHorz.Value
    End If
    
End Sub

Private Sub krsHorz_KeyDown(KeyCode As Integer, Shift As Integer)

    picView.SetFocus
    Call picView_KeyDown(KeyCode, Shift)
    KeyCode = 0

End Sub


Private Sub krsHorz_MouseWheelScroll(ByVal hWndMouseOver As Long, ByVal Delta As Long, ByVal AutoHandled As Boolean)

    picView.SetFocus
    Call krsHorz.IncrementValue(Delta * (krsHorz.SmallChange * 2))

End Sub


Private Sub krsHorz_Scroll()

    krsHorz_Change

End Sub


Private Sub imgSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Set the drag offset
    mlDragOffset = X

End Sub


Private Sub imgSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lNewX As Long

    'Move the split
    If mlDragOffset > 0 Then
        lNewX = imgSplit.Left + ((X - mlDragOffset) / Screen.TwipsPerPixelX)
        If lNewX < mlMinSplit Then
            lNewX = mlMinSplit
        ElseIf lNewX > mlMaxSplit Then
            lNewX = mlMaxSplit
        End If
        If imgSplit.Left <> lNewX - 1 Then
            imgSplit.Left = lNewX - 1
            lvwFiles.Width = lNewX
            'Resize the Tab control and picture frames
            tbsMain.Move lvwFiles.Width + imgSplit.Width, lvwFiles.Top, _
                Me.ScaleWidth - lvwFiles.Width - imgSplit.Width
            If (tbsMain.ClientHeight - krsHorz.Height - 12) > 0 Then
                picFrame.Visible = True
                picFrame.Move tbsMain.ClientLeft + 2, tbsMain.ClientTop + 4, tbsMain.ClientWidth - 4, tbsMain.ClientHeight - 6
            Else
                picFrame.Visible = False
            End If
        End If
    End If

End Sub


Private Sub imgSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mlDragOffset = 0
    imgSplit.Left = lvwFiles.Width - 1

End Sub


Private Sub krsVert_KeyDown(KeyCode As Integer, Shift As Integer)

    picView.SetFocus
    Call picView_KeyDown(KeyCode, Shift)
    KeyCode = 0
    
End Sub

Private Sub krsVert_MouseWheelScroll(ByVal hWndMouseOver As Long, ByVal Delta As Long, ByVal AutoHandled As Boolean)

    picView.SetFocus
    Call krsVert.IncrementValue(Delta * (krsVert.SmallChange * 2))
    
End Sub

Private Sub krsVert_Resize(ByVal SystemForced As Boolean)

    If SystemForced Then
        krsHorz.OverlapTwips = krsVert.Width * Screen.TwipsPerPixelX
        DoEvents
        Call picFrame_Resize
    End If
    
End Sub

Private Sub lvwFiles_DblClick()

Dim lRet    As Long
Dim sFile   As String
Dim uPt     As PointAPI
Dim oItem   As ListItem

    'Test for a hit in the ListView
    lRet = GetCursorPos(uPt)
    lRet = ScreenToClient(lvwFiles.hWnd, uPt)
    Set oItem = lvwFiles.HitTest(uPt.X * Screen.TwipsPerPixelX, uPt.Y * Screen.TwipsPerPixelY)
    'If the double-click is on an item
    If Not oItem Is Nothing Then
        'Set the currently selected item to the hit item
        Set lvwFiles.SelectedItem = oItem
        'Add a New Tab
        sFile = filFiles.Path & IIf(Right$(filFiles.Path, 1) <> "\", "\", "") & oItem.Text
        Call AddTab(sFile)
    End If
    
End Sub

Private Sub lvwFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)

Dim bFound  As Boolean
Dim iIdx    As Integer
Dim sPath   As String

    If Not Item Is Nothing Then
        If Not lvwFiles.SelectedItem Is Nothing Then
            If lvwFiles.SelectedItem.Index = Item.Index Then
                Call EnableTools
                sPath = filFiles.Path & IIf(Right$(filFiles.Path, 1) <> "\", "\", "")
                If (Item.Index <> mlItemIndex) Or _
                  (mImages(1).Filename <> sPath & Item.Text) Then
                    'Get the path to the image file
                    sPath = sPath & Item.Text
                    mImages(1).Filename = sPath
                    'Select the preview tab which forces a LoadImage
                    Set tbsMain.SelectedItem = tbsMain.Tabs(1)
                    DoEvents
                    'Force an update
                    picView.Move 0, 0
                    Call picFrame_Resize
                    'Store the index
                    DoEvents
                    mlItemIndex = Item.Index
                End If
            End If
        End If
    
    Else
        'Select nothing
        mImages(1).Filename = ""
        If tbsMain.SelectedItem.Index <> 1 Then
            Set tbsMain.SelectedItem = tbsMain.Tabs(1)
        Else
            Call tbsMain_Click
        End If
    End If
    
End Sub

Private Sub mnuFile_Edit_Click()

    Call Edit
    
End Sub

Private Sub mnuFile_Exit_Click()

    Unload Me
    End

End Sub

Private Sub mnuFile_Select_Click()

    Call Browse

End Sub

Private Sub mnuRecent_Clear_Click()

    If MsgBox("Are you sure you want to completely remove the Recent Files Log?", _
        vbYesNo, "Confirm Delete") = vbYes Then
        Call ClearRecentFiles(mnuRecent_Folders)
        Call EnableTools
    End If
    
End Sub

Private Sub mnuRecent_Folders_Click(Index As Integer)

Dim sFolder As String

    On Error Resume Next
    
    sFolder = mnuRecent_Folders(Index).Tag
    If Len(sFolder) > 0 Then
        mlItemIndex = -1
        filFiles.Path = sFolder
        Call AddRecentFile(filFiles.Path & IIf(Right$(filFiles.Path, 1) <> "\", "\", ""), mnuRecent_Folders)
        Call FillFileList(filFiles, lvwFiles)
        Call lvwFiles_ItemClick(lvwFiles.SelectedItem)
    End If
    
    Call EnableTools
    
End Sub


Private Sub mnuRecent_Logging_Click()

Dim lRet As Long

    mnuRecent_Logging.Checked = Not mnuRecent_Logging.Checked
    lRet = SetInitEntry("Main", "Logging", IIf(mnuRecent_Logging.Checked, "On", "Off"))
    
End Sub

Private Sub mnuTabs_Click()

    If tbsMain.SelectedItem Is Nothing Then
        mnuTabs_Sub(1).Caption = "&Delete"
        mnuTabs_Sub(1).Enabled = False
    Else
        If tbsMain.SelectedItem.Index = 1 Then
            mnuTabs_Sub(1).Caption = "&Remove"
            mnuTabs_Sub(1).Enabled = False
        Else
            mnuTabs_Sub(1).Caption = "&Remove '" & _
                tbsMain.SelectedItem.Caption & "' Tab"
        End If
    End If
    
End Sub

Private Sub mnuTabs_Sub_Click(Index As Integer)

Dim lRet    As Long
Dim sFile   As String
Dim oTab    As MSComctlLib.Tab
Dim oItem   As MSComctlLib.ListItem

    Select Case Index
        Case 0  'Add Tab(s)
            If Not lvwFiles.SelectedItem Is Nothing Then
                For Each oItem In lvwFiles.ListItems
                    If oItem.Selected Then
                        If tbsMain.Tabs.Count <= MAX_TABS Then
                            sFile = filFiles.Path & IIf(Right$(filFiles.Path, 1) <> "\", "\", "") & oItem.Text
                            Call AddTab(sFile)
                        Else
                            Exit For
                        End If
                    End If
                Next
            End If
            Call EnableTools
            
        Case 1  'Remove Tab
            Call DelTab
        
        Case 3  'Multi-Row
            Set oTab = tbsMain.SelectedItem
            tbsMain.MultiRow = Not tbsMain.MultiRow
            mnuTabs_Sub(3).Checked = tbsMain.MultiRow
            lRet = SetInitEntry("Main", "MultiRow Tabs", CStr(tbsMain.MultiRow))
            Set tbsMain.SelectedItem = oTab
            Call Form_Resize
            
    End Select
    
End Sub

Private Sub mnuToolView_Sub_Click(Index As Integer)

    If Index > 2 Then
        If Not mnuToolView_Sub(Index).Checked Then
        
            mnuToolView_Sub(Index).Checked = True
            
            Select Case Index
                Case 3  'Align Bottom
                    mnuToolView_Sub(4).Checked = False
                    Call SetInitEntry("Toolbar", "Text Align", "0")
                Case 4  'Align Right
                    mnuToolView_Sub(3).Checked = False
                    Call SetInitEntry("Toolbar", "Text Align", "1")
                Case 6  'Large Icons
                    mnuToolView_Sub(7).Checked = False
                    Call SetInitEntry("Toolbar", "Icon Size", "0")
                Case 7  'Small Icons
                    mnuToolView_Sub(6).Checked = False
                    Call SetInitEntry("Toolbar", "Icon Size", "1")
            End Select
            
            Call SetupToolbar
            
        End If
        
    ElseIf Index = 2 Then
        'Show Text
        mnuToolView_Sub(Index).Checked = Not mnuToolView_Sub(Index).Checked
        Call SetInitEntry("Toolbar", "Show Text", CStr(mnuToolView_Sub(Index).Checked))
        Call SetupToolbar
        
    ElseIf Index = 0 Then
        'Show/hide the toolbar
        If mnuToolView_Sub(0).Caption = "&Hide" Then
            mnuToolView_Sub(0).Caption = "&Show"
            Call SetInitEntry("Toolbar", "Visible", "False")
            tbrMain.Visible = False
        Else
            mnuToolView_Sub(0).Caption = "&Hide"
            Call SetInitEntry("Toolbar", "Visible", "True")
            tbrMain.Visible = True
        End If
        Call Form_Resize
        
    End If
    
End Sub

Private Sub mnuViews_View_Click(Index As Integer)

'This menu is used instead of the MSCommCtrl.ButtonMenu,
'because there is no way to checkmark the ButtonMenu.

Dim iIdx As Integer
Dim lRet As Long

    If Not mnuViews_View(Index).Checked Then
        For iIdx = 0 To mnuViews_View.UBound
            mnuViews_View(iIdx).Checked = False
        Next
        mnuViews_View(Index).Checked = True
        
        Select Case Index
            Case 0  'Large Icons
                lvwFiles.View = lvwIcon
            Case 1  'Small Icons
                lvwFiles.View = lvwSmallIcon
            Case 2  'List
                lvwFiles.View = lvwList
            Case 3  'Details
                lvwFiles.View = lvwReport
        End Select
    
        DoEvents
        If lvwFiles.ListItems.Count = 0 Then
            Call FillFileList(filFiles, lvwFiles)
        End If
        'Must reassign ImageLists on View change, since
        'VB doesn't know about the System ImageLists.
        Call AssignSystemImageLists(filFiles.Path, lvwFiles)
        lvwFiles.Visible = False
        lvwFiles.Visible = True
        lRet = SetInitEntry("Main", "Last View", CStr(Index))
        
    End If
    
End Sub

Private Sub picFrame_OLECompleteDrag(Effect As Long)

    Effect = vbDropEffectNone

End Sub

Private Sub picFrame_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Effect = DragDrop(Data)

End Sub

Private Sub picFrame_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    Effect = DragOver(Data)

End Sub

Private Sub picFrame_Resize()

Dim iMax    As Integer

    'Resize the view picture and move the scrollbars and corner picture
    krsVert.Move picFrame.ScaleWidth - krsVert.Width, 0, krsVert.Width, picFrame.ScaleHeight - krsHorz.Height
    krsHorz.Move 0, picFrame.ScaleHeight - krsHorz.Height, picFrame.ScaleWidth
    
    'Reset picView's position, if needed
    If picView.Width <= picFrame.ScaleWidth - krsVert.Width Then
        picView.Left = 0
    ElseIf picView.Left + picView.Width < picFrame.ScaleWidth - krsVert.Width Then
        picView.Left = picFrame.ScaleWidth - krsVert.Width - picView.Width
    End If
    If picView.Height <= picFrame.ScaleHeight - krsHorz.Height Then
        picView.Top = 0
    ElseIf picView.Top + picView.Height < picFrame.ScaleHeight - krsHorz.Height Then
        picView.Top = picFrame.ScaleHeight - krsHorz.Height - picView.Height
    End If
    
    'Reset the scrollbars (Max, Value and LargeChange)
    mbNoScroll = True
    iMax = picView.Width - (picFrame.ScaleWidth - krsVert.Width)
    If iMax <= 0 Then
        iMax = 0
    End If
    krsHorz.Enabled = (iMax > 0)
    If krsHorz.Enabled Then
        krsHorz.Max = iMax
        krsHorz.Value = -picView.Left
        krsHorz.SmallChange = 5
        krsHorz.LargeChange = (picFrame.ScaleWidth - krsVert.Width) / 2
    End If
    mImages(tbsMain.SelectedItem.Index).HPos = -picView.Left
    mImages(tbsMain.SelectedItem.Index).HMax = iMax
    
    iMax = picView.Height - (picFrame.ScaleHeight - krsHorz.Height)
    If iMax <= 0 Then
        iMax = 0
    End If
    krsVert.Enabled = (iMax > 0)
    If krsVert.Enabled Then
        krsVert.Max = iMax
        krsVert.Value = -picView.Top
        krsVert.SmallChange = 5
        krsVert.LargeChange = (picFrame.ScaleHeight - krsHorz.Height) / 2
    End If
    mImages(tbsMain.SelectedItem.Index).VPos = -picView.Top
    mImages(tbsMain.SelectedItem.Index).VMax = iMax
    mbNoScroll = False
    sbrMain.Panels(1).MinWidth = lvwFiles.Width + 4
    
End Sub

Private Sub picView_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyLeft
            Call krsHorz.IncrementValue(-krsHorz.SmallChange)
        Case vbKeyUp
            Call krsVert.IncrementValue(-krsVert.SmallChange)
        Case vbKeyRight
            Call krsHorz.IncrementValue(krsHorz.SmallChange)
        Case vbKeyDown
            Call krsVert.IncrementValue(krsVert.SmallChange)
        Case vbKeyPageUp
            Call krsVert.IncrementValue(-krsVert.LargeChange)
        Case vbKeyPageDown
            Call krsVert.IncrementValue(krsVert.LargeChange)
    End Select
    
End Sub

Private Sub picView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    sbrMain.Panels(3).Text = "Mouse: " & CStr(X) & ", " & CStr(Y)
    If Not tmrPos.Enabled Then
        tmrPos.Enabled = True
    End If
    
End Sub

Private Sub picView_OLECompleteDrag(Effect As Long)

    Effect = vbDropEffectNone

End Sub


Private Sub picView_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Effect = DragDrop(Data)

End Sub


Private Sub picView_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    Effect = DragOver(Data)

End Sub


Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    
        Case "Folder"   'Change Folders
            Call mnuFile_Select_Click
            
        Case "AddTabs"
            Call mnuTabs_Sub_Click(0)
            
        Case "DelTab"
            Call mnuTabs_Sub_Click(1)
        
        Case "Edit"     'Shell the associated editor for the selected image
            Call mnuFile_Edit_Click
            
        Case "Exit"     'End the program
            Call mnuFile_Exit_Click
            
    End Select
    
End Sub

Private Sub Edit()

'Shell the associated editor for the selected image

Dim lRet    As Long
Dim sFile   As String

    sFile = mImages(tbsMain.SelectedItem.Index).Filename
    If Len(sFile) > 0 Then
        lRet = ShellExecute(Me.hWnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    End If

End Sub

Private Sub tbrMain_ButtonDropDown(ByVal Button As MSComctlLib.Button)

Dim lFlags  As Long
Dim fLeft   As Single
Dim fTop    As Single

    lFlags = vbPopupMenuLeftAlign Or vbPopupMenuRightButton
    fLeft = Button.Left
    fTop = tbrMain.Height
    
    Select Case Button.Key
        Case "Folder"
            PopupMenu mnuRecent, lFlags, fLeft, fTop
        
        Case "Views"
            PopupMenu mnuViews, lFlags, fLeft, fTop
        
    End Select
    
End Sub


Private Sub tbrMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        Call PopupMenu(mnuViews_ToolView, vbPopupMenuCenterAlign Or _
            vbPopupMenuRightButton, X / Screen.TwipsPerPixelX, _
            Y / Screen.TwipsPerPixelY + 2)
    End If
    
End Sub

Private Sub tbsMain_BeforeClick(Cancel As Integer)

    If Not tbsMain.SelectedItem Is Nothing Then
        tbsMain.SelectedItem.Image = "Off"
    End If
    
End Sub

Private Sub tbsMain_Click()

    If Len(mImages(tbsMain.SelectedItem.Index).Filename) > 0 Then
        Call LoadImage
        picView.Move -mImages(tbsMain.SelectedItem.Index).HPos, -mImages(tbsMain.SelectedItem.Index).VPos
        Call picFrame_Resize
    Else
        Set picView.Picture = LoadPicture()
        picView.Move 0, 0, 50, 50
        Call picFrame_Resize
        picView.Visible = (picView.Picture > 0)
        sbrMain.Panels(2).Text = ""
    End If
    
    tbsMain.SelectedItem.Image = "On"
    Call EnableTools
    
End Sub

Private Sub tbsMain_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iIdx As Integer

    If (Shift And vbCtrlMask) = vbCtrlMask Then
        If KeyCode = vbKeyTab Then
            If (Shift And vbShiftMask) = vbShiftMask Then
                iIdx = tbsMain.SelectedItem.Index - 1
                If iIdx < 1 Then
                    iIdx = tbsMain.Tabs.Count
                End If
            Else
                iIdx = tbsMain.SelectedItem.Index + 1
                If iIdx > tbsMain.Tabs.Count Then
                    iIdx = 1
                End If
            End If
            If iIdx <> tbsMain.SelectedItem.Index Then
                Set tbsMain.SelectedItem = tbsMain.Tabs(iIdx)
            End If
        End If
    End If
    
End Sub

Private Sub tbsMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim iIdx    As Integer
Dim fX      As Single
Dim fY      As Single

    If Button = vbRightButton Then
        fX = tbsMain.Left + (X / Screen.TwipsPerPixelX)
        fY = tbsMain.Top + (Y / Screen.TwipsPerPixelY)
        For iIdx = 1 To tbsMain.Tabs.Count
            If fX >= tbsMain.Tabs(iIdx).Left And _
              fX < tbsMain.Tabs(iIdx).Left + _
              tbsMain.Tabs(iIdx).Width _
              And fY >= tbsMain.Tabs(iIdx).Top And _
              fY < tbsMain.Tabs(iIdx).Top + _
              tbsMain.Tabs(iIdx).Height Then
                If tbsMain.SelectedItem.Index <> iIdx Then
                    Set tbsMain.SelectedItem = tbsMain.Tabs(iIdx)
                    DoEvents
                End If
                Exit For
            End If
        Next
        For iIdx = 1 To 10
            DoEvents
        Next
    End If
    
End Sub

Private Sub tbsMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim iIdx As Integer

    If Button = vbRightButton Then
        For iIdx = 1 To 10
            DoEvents
        Next
        Call PopupMenu(mnuTabs, vbPopupMenuCenterAlign Or _
            vbPopupMenuRightButton, tbsMain.Left + _
            (X / Screen.TwipsPerPixelX), tbsMain.ClientTop)
    End If
    
End Sub

Private Sub tbsMain_OLECompleteDrag(Effect As Long)

    Effect = vbDropEffectNone
    
End Sub

Private Sub tmrPos_Timer()

Dim lRet    As Long
Dim ptCurs  As PointAPI

    lRet = GetCursorPos(ptCurs)
    lRet = ScreenToClient(picView.hWnd, ptCurs)
    If ptCurs.X < 0 Or ptCurs.X > picView.ScaleWidth Or ptCurs.Y < 0 Or ptCurs.Y > picView.ScaleHeight Then
        sbrMain.Panels(3).Text = ""
        tmrPos.Enabled = False
    End If
    
End Sub

Private Sub krsVert_Change()

    If Not mbNoScroll Then
        'Move the image
        picView.Top = -krsVert.Value
        mImages(tbsMain.SelectedItem.Index).VPos = krsVert.Value
    End If
    
End Sub


Private Sub krsVert_Scroll()

    krsVert_Change
    
End Sub


