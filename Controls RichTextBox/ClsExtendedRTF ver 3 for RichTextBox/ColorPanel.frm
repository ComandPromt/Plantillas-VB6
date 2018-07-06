VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form ColourPanel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Style Selector"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrmMaterialButtons 
      Caption         =   "Materials"
      Height          =   2295
      Left            =   4560
      TabIndex        =   54
      Top             =   3480
      Width           =   1095
      Begin VB.CommandButton btnEditMaterialsOpen 
         Caption         =   "New"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton btnEditMaterialsOpen 
         Caption         =   "Edit"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   735
         Width           =   855
      End
      Begin VB.CommandButton btnEditMaterialsOpen 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   56
         Top             =   1230
         Width           =   855
      End
      Begin VB.CommandButton btnEditMaterialsOpen 
         Caption         =   "Restore"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   55
         Top             =   1725
         Width           =   855
      End
   End
   Begin VB.Frame FrmStyleButtons 
      Caption         =   "Style"
      Height          =   2295
      Left            =   3360
      TabIndex        =   49
      Top             =   3480
      Width           =   1095
      Begin VB.CommandButton CmdCreateStyle 
         Caption         =   "Restore"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   53
         Top             =   1725
         Width           =   855
      End
      Begin VB.CommandButton CmdCreateStyle 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   52
         Top             =   1230
         Width           =   855
      End
      Begin VB.CommandButton CmdCreateStyle 
         Caption         =   "Edit"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   51
         ToolTipText     =   "To Edit, Clcik this then Release the part you wish to change."
         Top             =   735
         Width           =   855
      End
      Begin VB.CommandButton CmdCreateStyle 
         Caption         =   "New"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FrmStyle 
      Caption         =   "Style"
      Height          =   1575
      Left            =   120
      TabIndex        =   39
      Top             =   8040
      Width           =   8895
      Begin VB.CommandButton cmdStyleAction 
         Caption         =   "Show"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   3720
         TabIndex        =   48
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdStyleAction 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   375
         Index           =   2
         Left            =   7560
         TabIndex        =   45
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdStyleAction 
         Caption         =   "Do It"
         Height          =   375
         Index           =   1
         Left            =   6240
         TabIndex        =   44
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdStyleAction 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   43
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtStyleName 
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Text            =   "Style"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.CommandButton CmdStyleer 
         Caption         =   "Store Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   41
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton CmdStyleer 
         Caption         =   "Store Text"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblStyle 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   47
         Top             =   600
         Width           =   6495
      End
      Begin VB.Label lblStyle 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   46
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame FrmNewMaterial 
      Caption         =   "Material"
      Height          =   1575
      Left            =   100
      TabIndex        =   16
      Top             =   6120
      Width           =   8895
      Begin VB.PictureBox PctMaterialChkOpt 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   4695
         TabIndex        =   34
         Top             =   960
         Width           =   4695
         Begin VB.CheckBox chkNewMaterial 
            Caption         =   "LightDark"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   1095
         End
         Begin VB.CheckBox chkNewMaterial 
            Caption         =   "InOut"
            Height          =   255
            Index           =   1
            Left            =   1335
            TabIndex        =   37
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton optNewMaterialTxtBck 
            Caption         =   "Back"
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   36
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton optNewMaterialTxtBck 
            Caption         =   "Text"
            Height          =   255
            Index           =   0
            Left            =   2550
            TabIndex        =   35
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.CommandButton btnEditMaterialsPanel 
         Caption         =   "&Close"
         Height          =   375
         Index           =   2
         Left            =   7560
         TabIndex        =   33
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton btnEditMaterialsPanel 
         Caption         =   "Save"
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   32
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtMaterialsName 
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Text            =   "User"
         Top             =   480
         Width           =   2895
      End
      Begin VB.CommandButton btnEditMaterialsPanel 
         Caption         =   "Do It"
         Height          =   375
         Index           =   0
         Left            =   6255
         TabIndex        =   30
         Top             =   1080
         Width           =   1215
      End
      Begin MSComCtl2.UpDown UDMaterialValue 
         Height          =   285
         Index           =   0
         Left            =   3585
         TabIndex        =   24
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtMaterialValue(0)"
         BuddyDispid     =   196624
         BuddyIndex      =   0
         OrigLeft        =   3720
         OrigTop         =   480
         OrigRight       =   3975
         OrigBottom      =   765
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox TxtMaterialValue 
         Height          =   285
         Index           =   5
         Left            =   7845
         TabIndex        =   23
         Text            =   "0"
         Top             =   480
         Width           =   510
      End
      Begin VB.TextBox TxtMaterialValue 
         Height          =   285
         Index           =   4
         Left            =   6900
         TabIndex        =   22
         Text            =   "0"
         Top             =   480
         Width           =   465
      End
      Begin VB.TextBox TxtMaterialValue 
         Height          =   285
         Index           =   3
         Left            =   5955
         TabIndex        =   21
         Text            =   "0"
         Top             =   480
         Width           =   465
      End
      Begin VB.TextBox TxtMaterialValue 
         Height          =   285
         Index           =   2
         Left            =   5010
         TabIndex        =   20
         Text            =   "0"
         Top             =   480
         Width           =   465
      End
      Begin VB.TextBox TxtMaterialValue 
         Height          =   285
         Index           =   1
         Left            =   4065
         TabIndex        =   19
         Text            =   "255"
         Top             =   480
         Width           =   465
      End
      Begin VB.TextBox TxtMaterialValue 
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   18
         Text            =   "0"
         Top             =   480
         Width           =   465
      End
      Begin MSComCtl2.UpDown UDMaterialValue 
         Height          =   285
         Index           =   1
         Left            =   4530
         TabIndex        =   25
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   255
         BuddyControl    =   "TxtMaterialValue(1)"
         BuddyDispid     =   196624
         BuddyIndex      =   1
         OrigLeft        =   4680
         OrigTop         =   480
         OrigRight       =   4935
         OrigBottom      =   765
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDMaterialValue 
         Height          =   285
         Index           =   2
         Left            =   5475
         TabIndex        =   26
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtMaterialValue(2)"
         BuddyDispid     =   196624
         BuddyIndex      =   2
         OrigLeft        =   5640
         OrigTop         =   480
         OrigRight       =   5895
         OrigBottom      =   765
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDMaterialValue 
         Height          =   285
         Index           =   3
         Left            =   6420
         TabIndex        =   27
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtMaterialValue(3)"
         BuddyDispid     =   196624
         BuddyIndex      =   3
         OrigLeft        =   6600
         OrigTop         =   480
         OrigRight       =   6855
         OrigBottom      =   765
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDMaterialValue 
         Height          =   285
         Index           =   4
         Left            =   7365
         TabIndex        =   28
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtMaterialValue(4)"
         BuddyDispid     =   196624
         BuddyIndex      =   4
         OrigLeft        =   7560
         OrigTop         =   480
         OrigRight       =   7815
         OrigBottom      =   765
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDMaterialValue 
         Height          =   285
         Index           =   5
         Left            =   8355
         TabIndex        =   29
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "TxtMaterialValue(5)"
         BuddyDispid     =   196624
         BuddyIndex      =   5
         OrigLeft        =   8520
         OrigTop         =   480
         OrigRight       =   8775
         OrigBottom      =   765
         Max             =   255
         Min             =   -255
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label LblMaterialValues 
         Caption         =   " Min  0-255 | Max  0-255 | Red  0-255 | Green  0-255 | Blue 0-255 | Multiplier/10 "
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   17
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.CheckBox ChkPreserveColour 
      Caption         =   "Preserve Colour Selection"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Frame FrmMode 
      Caption         =   "Mode"
      Height          =   735
      Left            =   7920
      TabIndex        =   12
      Top             =   120
      Width           =   975
      Begin MSComCtl2.UpDown UDmode 
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "TxtMode"
         BuddyDispid     =   196628
         OrigLeft        =   600
         OrigTop         =   240
         OrigRight       =   855
         OrigBottom      =   615
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox TxtMode 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Text            =   "0"
         Top             =   240
         Width           =   480
      End
   End
   Begin RichTextLib.RichTextBox RTBClrDemo 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"ColorPanel.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frmDescription 
      Caption         =   "Description"
      Height          =   2595
      Left            =   1920
      TabIndex        =   8
      Top             =   120
      Width           =   2655
      Begin VB.Label LblDescription 
         Caption         =   "LblDescription"
         Height          =   2295
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.ListBox LstSubStyle 
      Height          =   2595
      Left            =   6120
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton CmdClrPanel 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton CmdClrPanel 
      Caption         =   "Do It"
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.OptionButton OptTextBack 
      Caption         =   "Back"
      Height          =   195
      Index           =   1
      Left            =   7920
      TabIndex        =   4
      Top             =   2280
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton OptTextBack 
      Caption         =   "Text"
      Height          =   195
      Index           =   0
      Left            =   7920
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox Chk_LR0_IO1 
      Caption         =   "InOut"
      Height          =   195
      Index           =   1
      Left            =   7920
      TabIndex        =   2
      Top             =   1275
      Width           =   975
   End
   Begin VB.CheckBox Chk_LR0_IO1 
      Caption         =   "LeftRight"
      Height          =   195
      Index           =   0
      Left            =   7920
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.ListBox LstStyle 
      Height          =   2595
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame FrmBuildstyle 
      Caption         =   "Style Builder"
      Height          =   735
      Left            =   4680
      TabIndex        =   59
      Top             =   120
      Width           =   1095
      Begin VB.CommandButton CmdOpenStyle 
         Caption         =   "Open"
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label LblAsteriskMessage 
      Caption         =   $"ColorPanel.frx":00CC
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   5295
   End
End
Attribute VB_Name = "ColourPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Copyright 2002 Roger Gilchrist
'rojgilkrist@hotmail.com
'very new; not much comment
'you'll have to work it out
Private Description(13) As String
Private DefDoItActive As Boolean
Private Enum List2Fillers
    Blank
    Spectrum
    rainbows
    Material
    random
    styler
End Enum
Private StyleValues(1) As String
Private BackStored As Boolean
Private ForeStored As Boolean
Private Const NL As String = vbNewLine
Private Demo As New ClsRTFFontPainter
Private PreserveColour As Boolean
Private Const SmallDisp As Long = 5300
Private Const HideEditor As Long = 5000
Private Const ShowEditor As Long = 3360
Private Const WM_SETREDRAW As Long = &HB
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub ActivateTools(DescNumber As Integer, Optional LRCaption As Boolean = True)

  Dim LeftRight As Boolean, InOut As Boolean, TxtBck As Boolean, Mde As Boolean

    'Default settings
    List2Filler Blank
    LeftRight = False
    InOut = False
    TxtBck = False
    Mde = False
    FrmMaterialButtons.Visible = DescNumber = 5
    FrmStyleButtons.Visible = DescNumber = 11
    Select Case DescNumber
      Case 1 '"candy"
        List2Filler Spectrum
        LeftRight = True
        InOut = True
        TxtBck = True
        Mde = True
        UDmode.Max = 2
      Case 2, 3  '2"blenderauto *","2blender **" '3"rainbow"
        LeftRight = True
        InOut = True
        TxtBck = True
      Case 4 '4"spectrum"'
        List2Filler Spectrum
        LeftRight = True
        InOut = True
        TxtBck = True
      Case 5  '5"materials"
        List2Filler Material
        LeftRight = True
        InOut = True
        TxtBck = True
        FrmMaterialButtons.Top = 120
        FrmMaterialButtons.Left = 4680
        btnEditMaterialsOpen(1).Enabled = LstSubStyle.ListIndex <> -1
        btnEditMaterialsOpen(2).Enabled = LstSubStyle.ListIndex <> -1
        If btnEditMaterialsPanel(0).Enabled = False Then
            btnEditMaterialsPanel(1).Left = btnEditMaterialsPanel(0).Left
        End If

      Case 6 '6"random"
        List2Filler random
        TxtBck = True
      Case 7 '7"dither *",7"dither2 *"
        LeftRight = True
        InOut = True
        TxtBck = True
      Case 8, 9, 10 '8"highlightuser *"'9"highlightuserauto *"'10"highlightuseruser **"
      Case 11 '11"'styles"
        List2Filler styler
        FrmStyleButtons.Top = 120
        FrmStyleButtons.Left = 4680
        CmdCreateStyle(1).Enabled = LstSubStyle.ListIndex <> -1
        CmdCreateStyle(2).Enabled = LstSubStyle.ListIndex <> -1
        If cmdStyleAction(1).Enabled = False Then
            'cmdStyleAction(3).Left = cmdStyleAction(0).Left
            'cmdStyleAction(0).Left = cmdStyleAction(1).Left
        End If
      Case 12 ' "Text colour *"
        TxtBck = True
        OptTextBack(0).Value = True
      Case 13 '"fuzzy *"
        LeftRight = True
        InOut = True
        TxtBck = True
      Case Else
    End Select
    Chk_LR0_IO1(0).Visible = LeftRight
    Chk_LR0_IO1(0).Caption = IIf(LRCaption, "LeftRight", "LightDark")
    Chk_LR0_IO1(1).Visible = InOut
    LstSubStyle.Visible = LstSubStyle.ListCount > 0
    LblDescription.Caption = Description(DescNumber)
    OptTextBack(0).Visible = TxtBck
    OptTextBack(1).Visible = TxtBck
    FrmMode.Visible = Mde
    ChkPreserveColour.Value = vbUnchecked 'turn off Preserve Colour Selection
    ChkPreserveColour.Visible = (InStr(LstStyle.List(LstStyle.ListIndex), "*") > 0) 'Hide if not needed
    '    DemoDefault

End Sub

Private Sub ApplyStyle(Index As Integer, DemoDoc As Boolean)

  Dim SA As Variant
  Dim PreserveColourLocal As Boolean

    PreserveColourLocal = PreserveColour
    PreserveColour = True
    SA = Split(StyleValues(Index), "|")
    Demo.StylesPainter Demo.Descriptor(CStr(SA(0)), CStr(SA(1)), CInt(SA(2)), CInt(SA(3)), CLng(SA(4)), CLng(SA(5)), SA(6) = "T", SA(7) = "T", SA(8) = "T"), PreserveColour
    PreserveColour = PreserveColourLocal

End Sub

Private Sub btnEditMaterialsOpen_Click(Index As Integer)

  Dim Min As Integer, Max As Integer, R As Integer, G As Integer, B As Integer, Multiplier As Double

    FrmNewMaterial.Top = ShowEditor
    chkNewMaterial(0).Value = Chk_LR0_IO1(0).Value
    chkNewMaterial(1).Value = Chk_LR0_IO1(1).Value
    optNewMaterialTxtBck(0).Value = OptTextBack(0).Value
    optNewMaterialTxtBck(1).Value = OptTextBack(1).Value

    Select Case Index
      Case 0 'do nothing
        txtMaterialsName.Text = Demo.MaterialsDefaultName
      Case 1
        Demo.MaterialsReader LCase$(LstSubStyle.List(LstSubStyle.ListIndex)), Min, Max, R, G, B, Multiplier
        txtMaterialsName.Text = LstSubStyle.List(LstSubStyle.ListIndex)
        TxtMaterialValue(0).Text = Min
        TxtMaterialValue(1).Text = Max
        TxtMaterialValue(2).Text = R
        TxtMaterialValue(3).Text = G
        TxtMaterialValue(4).Text = B
        TxtMaterialValue(5).Text = Multiplier * 10
      Case 2
        Demo.MaterialsDelete LCase$(LstSubStyle.List(LstSubStyle.ListIndex))
        List2Filler Material
        FrmNewMaterial.Top = HideEditor

      Case 3
        Demo.MaterialsRestore
        List2Filler Material
        FrmNewMaterial.Top = HideEditor

    End Select

End Sub

Private Sub btnEditMaterialsPanel_Click(Index As Integer)

    Select Case Index
      Case 0

        MaterialManual False
        Unload ColourPanel

      Case 1
        MaterialManual True, True
        List2Filler Material
        FrmNewMaterial.Top = HideEditor

      Case 2
        FrmNewMaterial.Top = HideEditor

    End Select

End Sub

Private Sub Check3_Click(Index As Integer)

    MaterialManual

End Sub

Private Sub Chk_LR0_IO1_Click(Index As Integer)

    DemoShow

End Sub

Private Sub chkNewMaterial_Click(Index As Integer)

    MaterialManual

End Sub

Private Sub ChkPreservEColour_Click()

    PreserveColour = (ChkPreserveColour.Value = vbChecked)

End Sub

Private Sub CmdClrPanel_Click(Index As Integer)

    If Index = 0 Then
        TakeAction False
    End If
    Unload ColourPanel

End Sub

Private Sub CmdCreateStyle_Click(Index As Integer)

  Dim txt As String, Bck As String

    FrmStyle.Top = ShowEditor

    Select Case Index
      Case 0 'new
        txtStyleName.Text = Demo.StylesDefaultName
      Case 1 'edit
        txtStyleName.Text = LstSubStyle.List(LstSubStyle.ListIndex)
        Demo.StylesReader LstSubStyle.List(LstSubStyle.ListIndex), txt$, Bck$
        lblStyle(0).Caption = txt
        lblStyle(1).Caption = Bck
        CmdStyleer(0).Caption = "Release Text"
        CmdStyleer(1).Caption = "Release Back"
        cmdStyleAction(1).Enabled = True
      Case 2 'delete
        Demo.StylesDelete LCase$(LstSubStyle.List(LstSubStyle.ListIndex))
        List2Filler styler
      Case 3 'restore
        Demo.StylesRestore
        List2Filler styler
    End Select

End Sub

Private Sub CmdOpenStyle_Click()

    Select Case CmdOpenStyle.Caption
      Case "Open"
        FrmStyle.Top = ShowEditor
        CmdOpenStyle.Caption = "Close"
      Case "Close"
        FrmStyle.Top = 8040
        CmdOpenStyle.Caption = "Open"
    End Select

End Sub

Private Sub cmdStyleAction_Click(Index As Integer)

  Dim SA As Variant

    Select Case Index
      Case 0 'save
        Demo.StylesCreator txtStyleName.Text & "^" & lblStyle(0).Caption & "^" & lblStyle(1).Caption, PreserveColour
        List2Filler styler
      Case 1 'do it
        RTBLooks.StylesCreator txtStyleName.Text & "^" & lblStyle(0).Caption & "^" & lblStyle(1).Caption, PreserveColour
        RTBLooks.StylesEngine txtStyleName.Text, PreserveColour
        Unload ColourPanel
      Case 2 'cancel
        FrmStyle.Top = 8040
        CmdOpenStyle.Caption = "Open"
      Case 3 'show
        DemoDefault
        Demo.StylesPainter lblStyle(0).Caption, PreserveColour
        Demo.StylesPainter lblStyle(1).Caption, PreserveColour
    End Select

End Sub

Private Sub CmdStyleer_Click(Index As Integer)

  Dim NoHit As Boolean

    Select Case Index
      Case 0
        Select Case CmdStyleer(0).Caption
          Case "Store Text"
            If Len(StyleValues(0)) Then
                ForeStored = True
                CmdStyleer(0).Caption = "Release Text"
                lblStyle(0).Caption = StyleValues(0)
              Else 'LEN(STYLEVALUES(0)) = FALSE
                NoHit = True
            End If

          Case "Release Text"
            ForeStored = False
            CmdStyleer(0).Caption = "Store Text"
            lblStyle(0).Caption = ""
        End Select

      Case 1

        Select Case CmdStyleer(1).Caption
          Case "Store Back"
            If Len(StyleValues(1)) Then
                BackStored = True
                CmdStyleer(1).Caption = "Release Back"
                lblStyle(1).Caption = StyleValues(1)
              Else 'LEN(STYLEVALUES(1)) = FALSE
                NoHit = True
            End If
          Case "Release Back"
            BackStored = False
            CmdStyleer(1).Caption = "Store Back"
            lblStyle(1).Caption = ""

        End Select

    End Select
    If NoHit Then
        MsgBox "No style settings available.", , "Style Builder"
    End If

    cmdStyleAction(3).Enabled = Len(lblStyle(1).Caption) > 0 And Len(lblStyle(0).Caption) > 0
    cmdStyleAction(0).Enabled = Len(lblStyle(1).Caption) > 0 And Len(lblStyle(0).Caption) > 0

End Sub

Private Sub DemoDefault()

  'select a length of text to use

    RTBClrDemo.SelStart = 0
    RTBClrDemo.SelLength = Len(RTBClrDemo.Text)
    Demo.ColourRemoveAll
    If Left$(LCase$(LstStyle.List(LstStyle.ListIndex)), 6) = "highli" Then
        RTBClrDemo.Find "not be easily read", 1
    End If

End Sub

Private Sub DemoShow(Optional FromMaterialManual As Boolean = False)

  'allows many dirrerent controls to call TakeAction

    DemoDefault
    If Not FromMaterialManual Then
        TakeAction True
    End If

End Sub

Private Sub Form_Load()

    Me.Height = SmallDisp
    FrmNewMaterial.Top = HideEditor

    Demo.AssignControls RTBClrDemo, ExtendedRTFDemo.CommonDialog1
    'The name of this   V_________V needs to match that being used by you RichTextBox
    CmdClrPanel(0).Enabled = RTBLooks.IsSelection
    btnEditMaterialsPanel(0).Enabled = CmdClrPanel(0).Enabled
    cmdStyleAction(1).Enabled = CmdClrPanel(0).Enabled

    CmdClrPanel(0).Caption = IIf(CmdClrPanel(0).Enabled, "Do It", "No Selection")
    btnEditMaterialsPanel(0).Caption = CmdClrPanel(0).Caption
    cmdStyleAction(1).Caption = CmdClrPanel(0).Caption

    DefDoItActive = CmdClrPanel(0).Enabled

    'LstStyle is Sorted=True as the rest of the form actually reads the lcase string value of this list
    'You can add either Lcase,Ucase or ProperCase strings here but make sure you use lcase everywhere else
    'Dont forget to add a description
    'Add a 1 * or 2 ** if the option includes 1 or 2 colour selection(s)
    With LstStyle
        .Clear
        .AddItem "Blender **"
        .AddItem "BlenderAuto *"
        .AddItem "Candy"
        .AddItem "Dither *"
        .AddItem "Fuzzy *"
        .AddItem "Dither2 *"
        .AddItem "Materials"
        .AddItem "Rainbow"
        .AddItem "Random"
        .AddItem "Spectrum"
        .AddItem "Styles"
        .AddItem "HighlightUser *"
        .AddItem "HighlightUserAuto *"
        .AddItem "HighlightUserUser **"
        .AddItem "Text colour *"

    End With 'LstStyle
    Description(0) = "Select a colour style" & NL & _
                "If there is 1 * you will be asked to select one colour" & NL & _
                "If there are 2 ** you will be asked to select two colours"
    Description(1) = "Candy" & NL & "" & NL & _
                "1. Select a Spectrum Setion." & NL & _
                "2. Select Text or Back."
    Description(2) = "Blend" & NL & _
                "Two colours are blended into each other." & NL & _
                "1. Check/Uncheck Inout." & NL & _
                "2. Select Text or Back."
    Description(3) = "Rainbow" & NL & _
                "A full width rainbow spectrum is created." & NL & _
                "1. Select a Direction." & NL & _
                "2. Select Text or Back."
    Description(4) = "Spectrum" & NL & _
                "One of 6 rainbow colour ranges is created." & NL & _
                "1. Select a Spectrum Setion." & NL & _
                "2. Check/UnCheck InOut" & NL & _
                "3. Check/UnCheck LeftRight" & NL & _
                "4. Select Text or Back."
    Description(5) = "Materials" & NL & _
                "Smoothly changing material colour spread is created." & NL & _
                "1. Select a Material" & NL & _
                "2. Check/UnCheck InOut" & NL & _
                "3. Check/UnCheck LeftRight" & NL & _
                "4. Select Text or Back."
    Description(6) = "Random" & NL & _
                "Each Character gets its own colour" & NL & _
                "1. Select a Colour range" & NL & _
                "2. Check/UnCheck InOut" & NL & _
                "3. Select Text or Back." & NL & NL & _
                "Remember this is really random. The sample is NOT exactly what you will get in the main document."
    Description(7) = "Dither" & NL & _
                "A selected colour is dithered from dark to light" & NL & _
                "1. Check/UnCheck LightDark" & NL & _
                "2. Select Text or Back."
    Description(8) = "RTFHighlightUser" & NL & _
                "Back Colour is selected by user." & NL & _
                "1. Select Back colour from ColorDialog" & NL & NL & _
                "Note there are also HighlghtHard versions of this routine if you want to hard code a colour." & NL & _
                "The Preserve Colour Selection Checkbox uses the Hard version to save time"
    Description(9) = "RTFHighlightUserAuto" & NL & _
                "Back Colour is selected by user, Text colour by Program." & NL & _
                "1. Select Back Colour from ColorDialog" & NL _
                & "2. Class creates a Contrasting Text Colour" & NL & NL & _
                "Note there are also HighlghtHard versions of this routine if you want to hard code a colour." & NL & _
                "The Preserve Colour Selection Checkbox uses the Hard version to save time"
    Description(10) = "RTFHighlightUserUser" & NL & _
                "Text and Back colour are selected by user." & NL & _
                "1. Select Back Colour from ColorDialog" & NL & _
                "2. Select Text Colour from ColorDialog" & NL & NL & _
                "Note there are also HighlghtHard versions of this routine if you want to hard code a colour." & NL & _
                "The Preserve Colour Selection Checkbox uses the Hard version to save time"
    Description(11) = "Styles" & NL & _
                "Combos of other styles"
    Description(12) = "Text Colour" & NL & _
                "Set text colour"
    Description(13) = "Fuzzy Colour" & NL & _
                "create colours around selected colour."

    ActivateTools 0  'turn everything off at first

End Sub

Private Sub List2Filler(Mode As List2Fillers)

  Dim TmpArray As Variant, TmpMember As Variant

    'LstSubStyle is deliberately left Sorted=False so that the panel can use list postion to read selections
    'For neatness try to set them in Alpha order
    'Mode = Blank just falls past the Select case leaving a blank list
    With LstSubStyle
        SendMessage LstSubStyle.hwnd, WM_SETREDRAW, False, ByVal 0& 'speed up list filling by disabling redraw while you add items
        .Clear
        Select Case Mode
          Case Spectrum
            .AddItem "s1RedYellow"
            .ItemData(.NewIndex) = 0
            .AddItem "s2YellowGreen"
            .ItemData(.NewIndex) = 1
            .AddItem "s3GreenCyan"
            .ItemData(.NewIndex) = 2
            .AddItem "s4CyanBlue"
            .ItemData(.NewIndex) = 3
            .AddItem "s5BlueMagenta"
            .ItemData(.NewIndex) = 4
            .AddItem "s6MagentaRed"
            .ItemData(.NewIndex) = 5
            .ListIndex = 0
          Case Blank 'do nothing

          Case Material
            TmpArray = Demo.MaterialsKnownColourNames
            For Each TmpMember In TmpArray
                .AddItem StrConv(TmpMember, vbProperCase)
            Next TmpMember
            .ListIndex = 0

          Case styler
            TmpArray = Demo.StylesKnownNames
            For Each TmpMember In TmpArray
                .AddItem StrConv(TmpMember, vbProperCase)
            Next TmpMember
            '.ListIndex = 0

          Case random
            .AddItem "All colours"
            .ItemData(.NewIndex) = 0
            .AddItem "Rainbow"
            .ItemData(.NewIndex) = 1
            .AddItem "s1RedYellow"
            .ItemData(.NewIndex) = 2
            .AddItem "s2YellowGreen"
            .ItemData(.NewIndex) = 3
            .AddItem "s3GreenCyan"
            .ItemData(.NewIndex) = 4
            .AddItem "s4CyanBlue"
            .ItemData(.NewIndex) = 5
            .AddItem "s5BlueMagenta"
            .ItemData(.NewIndex) = 6
            .AddItem "s6MagentaRed"
            .ItemData(.NewIndex) = 7
            .AddItem "grey"
            .ItemData(.NewIndex) = 8
            .ListIndex = 0
        End Select
    End With 'LstSubStyle
    SendMessage LstSubStyle.hwnd, WM_SETREDRAW, True, ByVal 0& 'restart listbox updates

End Sub

Private Sub LstStyle_Click()

  'note use lcase name in Case "whatever" when adding new ones
  'work out which tools should be active
  'and whether LstSubStyle needs to show any thing

    FrmNewMaterial.Top = HideEditor
    PreserveColour = False
    Select Case LCase$(LstStyle.List(LstStyle.ListIndex))
      Case "candy"
        ActivateTools 1
      Case "blender **", "blenderauto *"
        ActivateTools 2
      Case "rainbow"
        ActivateTools 3
      Case "spectrum"
        ActivateTools 4
      Case "materials"
        ActivateTools 5
      Case "random"
        ActivateTools 6
      Case "dither *", "dither2 *"
        ActivateTools 7, False
      Case "highlightuser *"
        ActivateTools 8
      Case "highlightuserauto *"
        ActivateTools 9
      Case "highlightuseruser **"
        ActivateTools 10
      Case "styles"
        ActivateTools 11
      Case "Text colour *"
        ActivateTools 12
      Case "fuzzy *"
        ActivateTools 13
      Case Else
        ActivateTools 0
    End Select

    DemoShow
    'turn on Preserve Colour Selection if necessary
    'value has to be set after DemoShow or you don't get initial colour choice option
    If InStr(LCase$(LstStyle.List(LstStyle.ListIndex)), "*") > 0 Then
        PreserveColour = True
        ChkPreserveColour.Value = IIf(PreserveColour, vbChecked, vbUnchecked)
    End If

End Sub

Private Sub LstSubStyle_Click()

    If FrmMaterialButtons.Visible Then
        btnEditMaterialsOpen(1).Enabled = LstSubStyle.ListIndex <> -1
        btnEditMaterialsOpen(2).Enabled = LstSubStyle.ListIndex <> -1
    End If
    If FrmStyleButtons.Visible Then
        CmdCreateStyle(1).Enabled = LstSubStyle.ListIndex <> -1
        CmdCreateStyle(2).Enabled = LstSubStyle.ListIndex <> -1
    End If

    'LstSubStyle is deliberately left Sorted=False so that the panel can use list postion to read selections

    DemoShow

End Sub

Private Sub MaterialManual(Optional DemoDoc As Boolean = True, Optional forceSave As Boolean = False)

  Dim Target As Variant
  Dim MatName As String

    Demo.ColourRemoveAll
    DemoShow True
    If DemoDoc Then
        Set Target = Demo
      Else 'DEMODOC = FALSE
        Set Target = RTBLooks
        forceSave = True 'always save anything applied to the main text
        'you can always delete it later if you don't want it
    End If
    If forceSave Then
        MatName$ = txtMaterialsName.Text
      Else 'FORCESAVE = FALSE
        MatName$ = "@@@@@@@@"
    End If
    Target.MaterialsCreator MatName, chkNewMaterial(0) = Checked, _
                            chkNewMaterial(1) = Checked, (optNewMaterialTxtBck(0).Value = True), _
                            Val(TxtMaterialValue(0).Text), _
                            Val(TxtMaterialValue(1).Text), _
                            Val(TxtMaterialValue(2).Text), _
                            Val(TxtMaterialValue(3).Text), _
                            Val(TxtMaterialValue(4).Text), _
                            Val(TxtMaterialValue(5).Text) / 10      '/10 because the value is coming from an UpDown controlled TextBox

End Sub

Private Sub optNewMaterialTxtBck_Click(Index As Integer)

    MaterialManual

End Sub

Private Sub OptTextBack_Click(Index As Integer)

    DemoShow

End Sub

Private Sub TakeAction(DemoDoc As Boolean)

  'PreserveColour prevents the colour dialog from firing every time you change things
  'for those tools which need user to select colours

  Dim TextBack As Boolean
  Dim InOut As Boolean
  Dim LeftRight As Boolean
  Dim CandyMode As Integer
  Dim style As String
  Dim CurSubStyleVal As Integer
  Dim CurStyle As String, ClrBack As Long, ClrText As Long, CurSubStyle As String
  Dim Target As Variant

    Demo.ColourRemoveAll
    If DemoDoc Then
        Set Target = Demo
        Target.ColourRemoveAll ' only on demo do you need to reset to basic
      Else 'DEMODOC = FALSE
        Set Target = RTBLooks
    End If
    TextBack = (OptTextBack(0).Value = True)
    LeftRight = (Chk_LR0_IO1(0).Value = vbChecked)
    InOut = (Chk_LR0_IO1(1).Value = vbChecked)
    CandyMode = Val(TxtMode.Text)
    style$ = LCase$(LstStyle.List(LstStyle.ListIndex))
    CurSubStyleVal = LstSubStyle.ListIndex
    If CurSubStyleVal > -1 Then
        CurSubStyle = LCase$(LstSubStyle.List(CurSubStyleVal))
    End If

    If Not PreserveColour Then
        Demo.StylesPainterPreserveColours style, ClrBack, ClrText, TextBack
    End If
    If style = "styles" Then
        Target.StylesEngine CurSubStyle, PreserveColour

      Else 'NOT STYLE...
        Target.StylesPainter Demo.Descriptor(style, CurSubStyle, CurSubStyleVal, CandyMode, ClrBack, ClrText, LeftRight, InOut, TextBack), PreserveColour
    End If
    CurStyle = Demo.Descriptor(style, CurSubStyle, CurSubStyleVal, CandyMode, ClrBack, ClrText, LeftRight, InOut, TextBack, True)
    If TextBack Then
        StyleValues(0) = CurStyle
        CmdStyleer(0).Enabled = Len(CurStyle) > 0
      Else 'TEXTBACK = FALSE
        StyleValues(1) = CurStyle
        CmdStyleer(1).Enabled = Len(CurStyle) > 0
    End If

End Sub

Private Sub TxtMaterialValue_Change(Index As Integer)

    MaterialManual

End Sub

Private Sub txtMode_Change()

    DemoShow

End Sub

':) Ulli's VB Code Formatter V2.13.6 (28/08/2002 2:39:34 PM) 27 + 665 = 692 Lines
