VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.UserControl ctxFormats 
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6576
   ScaleHeight     =   4560
   ScaleWidth      =   6576
   Begin VB.CommandButton cmdOpen 
      Height          =   285
      Left            =   1476
      Picture         =   "ctxFormats.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   84
      Width           =   300
   End
   Begin VB.CommandButton cmdSave 
      Height          =   285
      Left            =   1806
      Picture         =   "ctxFormats.ctx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   84
      Width           =   300
   End
   Begin MSComctlLib.TreeView trvFormats 
      Height          =   4044
      Left            =   84
      TabIndex        =   1
      Top             =   420
      Width           =   2028
      _ExtentX        =   3577
      _ExtentY        =   7154
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   0
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   2520
      Top             =   0
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   259
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontStrikeThru  =   -1  'True
      FontUnderLine   =   -1  'True
   End
   Begin VB.PictureBox picFormat 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4464
      Left            =   2100
      ScaleHeight     =   4464
      ScaleWidth      =   4464
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   4464
      Begin MSComDlg.CommonDialog comDlgFile 
         Left            =   924
         Top             =   0
         _ExtentX        =   699
         _ExtentY        =   699
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "OBF"
         Filter          =   "Format files (*.obf)|*.obf|All files (*.*)|*.*"
         Flags           =   4
      End
      Begin VB.CheckBox chkSunken 
         Caption         =   "Sunken"
         Height          =   264
         Left            =   2940
         TabIndex        =   9
         Top             =   924
         Width           =   1356
      End
      Begin VB.ComboBox cobVertAlignment 
         Height          =   288
         Left            =   2856
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1764
         Width           =   1440
      End
      Begin VB.Frame Frame1 
         Caption         =   "BackGradient"
         Height          =   1692
         Left            =   84
         TabIndex        =   27
         Top             =   2604
         Width           =   4350
         Begin VB.PictureBox picSelectBitmap 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   852
            Left            =   168
            ScaleHeight     =   852
            ScaleWidth      =   4128
            TabIndex        =   42
            Top             =   756
            Visible         =   0   'False
            Width           =   4128
            Begin VB.PictureBox picBitmap 
               AutoRedraw      =   -1  'True
               Height          =   768
               Left            =   1092
               ScaleHeight     =   720
               ScaleWidth      =   2232
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   0
               Width           =   2280
               Begin VB.Image imgBitmap 
                  Height          =   348
                  Left            =   84
                  MousePointer    =   15  'Size All
                  Top             =   84
                  Width           =   348
               End
            End
            Begin VB.CommandButton cmdBitmapClear 
               Height          =   285
               Left            =   3444
               Picture         =   "ctxFormats.ctx":048D
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   0
               Width           =   300
            End
            Begin VB.CommandButton cmdBitmapIcon 
               Height          =   285
               Left            =   3780
               Picture         =   "ctxFormats.ctx":0A17
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   0
               Width           =   300
            End
            Begin VB.Label Label7 
               Caption         =   "&Bitmap:"
               Height          =   264
               Left            =   0
               TabIndex        =   43
               Top             =   0
               Width           =   1104
            End
         End
         Begin VB.TextBox txtGradAlpha 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   3528
            TabIndex        =   18
            ToolTipText     =   "Alpha: 0 to 255"
            Top             =   336
            Width           =   684
         End
         Begin VB.PictureBox picColorOffset 
            BorderStyle     =   0  'None
            Height          =   348
            Left            =   168
            ScaleHeight     =   348
            ScaleWidth      =   4128
            TabIndex        =   38
            Top             =   1176
            Width           =   4128
            Begin VB.TextBox txtGradBri 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   3360
               TabIndex        =   25
               ToolTipText     =   "Percent: -100 to 100"
               Top             =   0
               Width           =   684
            End
            Begin VB.TextBox txtGradSat 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   2604
               TabIndex        =   24
               ToolTipText     =   "Percent: -100 to 100"
               Top             =   0
               Width           =   684
            End
            Begin VB.TextBox txtGradHue 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   1092
               TabIndex        =   23
               ToolTipText     =   "Offset: 0 to 360"
               Top             =   0
               Width           =   684
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "Sat/Bri %: "
               Height          =   264
               Left            =   1680
               TabIndex        =   40
               Top             =   0
               Width           =   936
            End
            Begin VB.Label Label6 
               Caption         =   "Hue offset:"
               Height          =   264
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Width           =   1104
            End
         End
         Begin VB.ComboBox cobGradType 
            Height          =   288
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   336
            Width           =   1440
         End
         Begin VB.CommandButton cmdGradSecondColor 
            Height          =   285
            Left            =   3912
            Picture         =   "ctxFormats.ctx":0B61
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1176
            Width           =   285
         End
         Begin VB.ComboBox cobGradSecondColor 
            Height          =   315
            Left            =   1260
            TabIndex        =   21
            Text            =   "cobGradSecondColor"
            Top             =   1176
            Width           =   2616
         End
         Begin VB.CommandButton cmdGradColor 
            Height          =   285
            Left            =   3912
            Picture         =   "ctxFormats.ctx":0CAB
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   756
            Width           =   285
         End
         Begin VB.ComboBox cobGradColor 
            Height          =   315
            Left            =   1260
            TabIndex        =   19
            Text            =   "cobGradColor"
            Top             =   756
            Width           =   2616
         End
         Begin VB.CheckBox chkTileAbsolute 
            Caption         =   "Absolute"
            Height          =   264
            Left            =   2856
            TabIndex        =   47
            Top             =   336
            Width           =   1356
         End
         Begin VB.Label labGradAlpha 
            Alignment       =   1  'Right Justify
            Caption         =   "Alpha:"
            Height          =   264
            Left            =   2436
            TabIndex        =   41
            Top             =   336
            Width           =   1020
         End
         Begin VB.Label Label8 
            Caption         =   "T&ype:"
            Height          =   264
            Left            =   168
            TabIndex        =   30
            Top             =   336
            Width           =   1104
         End
         Begin VB.Label labGradSecondColor 
            Caption         =   "&SecondColor:"
            Height          =   264
            Left            =   168
            TabIndex        =   29
            Top             =   1176
            Width           =   1104
         End
         Begin VB.Label labGradColor 
            Caption         =   "Co&lor:"
            Height          =   264
            Left            =   168
            TabIndex        =   28
            Top             =   756
            Width           =   1104
         End
      End
      Begin VB.ComboBox cobBorder 
         Height          =   288
         Left            =   1344
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   924
         Width           =   1440
      End
      Begin VB.ComboBox cobHorAlignment 
         Height          =   288
         Left            =   1344
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1764
         Width           =   1440
      End
      Begin VB.CommandButton cmdForeColor 
         Height          =   285
         Left            =   3996
         Picture         =   "ctxFormats.ctx":0DF5
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   504
         Width           =   285
      End
      Begin VB.ComboBox cobForeColor 
         Height          =   288
         Left            =   1344
         TabIndex        =   6
         Text            =   "cobForeColor"
         Top             =   504
         Width           =   2616
      End
      Begin VB.CommandButton cmdFont 
         Height          =   285
         Left            =   3996
         Picture         =   "ctxFormats.ctx":0F3F
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   84
         Width           =   285
      End
      Begin VB.TextBox txtFont 
         BackColor       =   &H8000000F&
         Height          =   288
         Left            =   1344
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   84
         Width           =   2616
      End
      Begin VB.TextBox txtOffsetX 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   1344
         TabIndex        =   14
         Top             =   2184
         Width           =   684
      End
      Begin VB.TextBox txtOffsetY 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   2100
         TabIndex        =   15
         Top             =   2184
         Width           =   684
      End
      Begin VB.ComboBox cobBorderColor 
         Height          =   288
         Left            =   1344
         TabIndex        =   10
         Text            =   "cobBorderColor"
         Top             =   1344
         Width           =   2616
      End
      Begin VB.CommandButton cmdBorderColor 
         Height          =   285
         Left            =   3996
         Picture         =   "ctxFormats.ctx":1089
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1344
         Width           =   285
      End
      Begin VB.TextBox txtPadding 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   3612
         TabIndex        =   16
         Top             =   2184
         Width           =   684
      End
      Begin VB.Label Label5 
         Caption         =   "&Border:"
         Height          =   264
         Left            =   252
         TabIndex        =   37
         Top             =   924
         Width           =   1104
      End
      Begin VB.Label Label3 
         Caption         =   "&Alignment:"
         Height          =   264
         Left            =   252
         TabIndex        =   36
         Top             =   1764
         Width           =   1104
      End
      Begin VB.Label Label2 
         Caption         =   "Fore&Color:"
         Height          =   264
         Left            =   252
         TabIndex        =   35
         Top             =   504
         Width           =   1104
      End
      Begin VB.Label Label1 
         Caption         =   "F&ont:"
         Height          =   264
         Left            =   252
         TabIndex        =   34
         Top             =   84
         Width           =   1104
      End
      Begin VB.Label Label4 
         Caption         =   "Bo&rderColor:"
         Height          =   264
         Left            =   252
         TabIndex        =   33
         Top             =   1344
         Width           =   1104
      End
      Begin VB.Label Label9 
         Caption         =   "O&ffset:"
         Height          =   264
         Left            =   252
         TabIndex        =   32
         Top             =   2184
         Width           =   1104
      End
      Begin VB.Label Label11 
         Caption         =   "&Padding:"
         Height          =   264
         Left            =   2856
         TabIndex        =   31
         Top             =   2184
         Width           =   936
      End
   End
   Begin VB.Label Label10 
      Caption         =   "&Formats:"
      Height          =   312
      Left            =   84
      TabIndex        =   0
      Top             =   84
      Width           =   1404
   End
End
Attribute VB_Name = "ctxFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
'   You are free to use this source as long as this copyright message
'     appears on your program's "About" dialog:
'
'   Outlook Bar Project
'   Copyright (c) 2002 Vlad Vissoultchev (wqweto@myrealbox.com)
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "ctxFormats"

'=========================================================================
' Events
'=========================================================================

Event Changed()

'=========================================================================
' API
'=========================================================================

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Any) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Type OLERGBQUAD
    R As Byte
    G As Byte
    b As Byte
    a As Byte
End Type

'=========================================================================
' Constants and variables
'=========================================================================

Private Const CAP_MSG               As String = "Formats Property Page"
Private Const STR_ALIGNMENT_HOR     As String = "Left|0|Center|1|Right|2" ' |LeftOfText|3|RightOfText|4"
Private Const STR_ALIGNMENT_VERT    As String = "Top|0|Middle|1|Bottom|2"
Private Const STR_BORDER            As String = "None|0|Fixed|1|Single3D|2|Double3D|3"
Private Const STR_GRADIENT          As String = "Solid|0|Horizontal|1|Vertical|2|Blend|3|Transparent|4|Color Offset|5|Alpha Blend|6|Stretched|7|Tiled|8"
Private Const STR_UNDEFINED         As String = "<Undefined>"
Private Const STR_FILTER_OBF        As String = "Format files (*.obf)|*.obf|All files (*.*)|*.*"
Private Const STR_FILTER_IMAGES     As String = "Icons And Graphics (*.ico;*.bmp;*.gif;*.jpg)|*.ico;*.bmp;*.gif;*.jpg|All files (*.*)|*.*"

Private m_oControl              As Object
Private m_vFormatControls       As Variant
Private m_oFormat               As cFormatDef
Private WithEvents m_oFont      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_bInSet                As Boolean
Private m_lListIndex            As Long
Private m_bDrag                 As Boolean
Private m_sX                    As Single
Private m_sY                    As Single

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunc As String)
    PushError sFunc, MODULE_NAME
    PopRaiseError
End Sub

Private Function ShowError(sFunc As String) As VbMsgBoxResult
    PushError sFunc, MODULE_NAME
    ShowError = PopShowError(CAP_MSG)
End Function

'=========================================================================
' Properties
'=========================================================================

Property Let Changed(ByVal bValue As Boolean)
    RaiseEvent Changed
End Property

'=========================================================================
' Methods
'=========================================================================

Private Sub pvFillListBox()
    Const FUNC_NAME     As String = "pvFillListBox"
    Dim vDef            As Variant
    Dim oFmt            As cFormatDef
    
    On Error GoTo EH
    '--- fill treeview
    With trvFormats.Nodes
        .Clear
        For Each vDef In m_vFormatControls
            Set oFmt = vDef
            If oFmt.ParentFmt Is Nothing Then
                .Add(, , oFmt.FullName, oFmt.Name).Expanded = True
            Else
                .Add(oFmt.ParentFmt.FullName, tvwChild, oFmt.FullName, oFmt.Name).Expanded = True
            End If
        Next
    End With
    Set trvFormats.SelectedItem = trvFormats.Nodes(m_lListIndex + 1)
    trvFormats_NodeClick trvFormats.SelectedItem
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvLngToStr(ByVal lValue As Long) As String
    Const FUNC_NAME     As String = "pvLngToStr"
    
    On Error GoTo EH
    If lValue <> LNG_UNDEFINED Then
        pvLngToStr = lValue
    Else
        pvLngToStr = "" ' STR_UNDEFINED
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvStrToLng(sValue As String) As Long
    Const FUNC_NAME     As String = "pvStrToLng"
    
    On Error GoTo EH
    If IsNumeric(Trim(sValue)) Then
        pvStrToLng = Val(sValue)
    Else
        pvStrToLng = LNG_UNDEFINED
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvFillControls(ByVal oFmt As cFormatDef)
    Const FUNC_NAME     As String = "pvFillControls"
    
    On Error GoTo EH
    m_bInSet = True
    With oFmt
        '--- load font
        CopyFont m_oFont, .Font
        pvFillFont txtFont, m_oFont
        '--- load combos
        pvLookupCombo cobForeColor, .ForeColor, True
        pvColorComboClick cobForeColor, cmdForeColor
        pvLookupCombo cobBorderColor, .BorderColor, True
        pvColorComboClick cobBorderColor, cmdBorderColor
        chkSunken.Value = IIf(.BorderSunken = ucsTriTrue, vbChecked, IIf(.BorderSunken = ucsTriFalse, vbUnchecked, vbGrayed))
        pvLookupCombo cobHorAlignment, .HorAlignment
        pvLookupCombo cobVertAlignment, .VertAlignment
        pvLookupCombo cobBorder, .Border
        txtOffsetX = pvLngToStr(.OffsetX)
        txtOffsetY = pvLngToStr(.OffsetY)
        txtPadding = pvLngToStr(.Padding)
        pvLookupCombo cobGradType, .BackGradient.GradientType
        pvLookupCombo cobGradColor, .BackGradient.Color, True
        pvColorComboClick cobGradColor, cmdGradColor
        pvLookupCombo cobGradSecondColor, .BackGradient.SecondColor, True
        pvColorComboClick cobGradSecondColor, cmdGradSecondColor
        If .BackGradient.GradientType = ucsGrdColorOffset Then
            txtGradHue = .BackGradient.OffsetHue
            txtGradSat = .BackGradient.PercentSaturation
            txtGradBri = .BackGradient.PercentBrightness
        Else
            txtGradHue = 0
            txtGradSat = 0
            txtGradBri = 0
        End If
        If .BackGradient.GradientType = ucsGrdAlphaBlend Then
            txtGradAlpha = .BackGradient.Alpha
        Else
            txtGradAlpha = 255
        End If
        Select Case .BackGradient.GradientType
        Case ucsGrdStretchBitmap, ucsGrdTileBitmap
            Set imgBitmap = .BackGradient.Picture
            pvCenterBitmap
        Case Else
            Set imgBitmap = Nothing
        End Select
        If .BackGradient.GradientType = ucsGrdTileBitmap Then
            chkTileAbsolute.Value = Abs(.BackGradient.TileAbsolutePosition)
        End If
    End With
    m_bInSet = False
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvUploadColor(ByVal oCombo As ComboBox, Optional ByVal clrDefault As OLE_COLOR) As OLE_COLOR
    If oCombo.ListIndex >= 0 Then
        pvUploadColor = oCombo.ItemData(oCombo.ListIndex)
    Else
        If Left(oCombo.Text, 1) = "#" Then
            pvUploadColor = C2Lng("&H" & Replace(oCombo.Text, "#", ""))
        Else
            pvUploadColor = clrDefault
        End If
    End If
End Function

Private Sub pvUploadFormat()
    Const FUNC_NAME     As String = "pvUploadFormat"
    
    On Error GoTo EH
    If Not m_oFormat Is Nothing Then
        With m_oFormat
            Set .Font = m_oFont
            .ForeColor = pvUploadColor(cobForeColor, .ForeColor)
            If cobHorAlignment.ListIndex >= 0 Then
                .HorAlignment = cobHorAlignment.ItemData(cobHorAlignment.ListIndex)
            End If
            If cobVertAlignment.ListIndex >= 0 Then
                .VertAlignment = cobVertAlignment.ItemData(cobVertAlignment.ListIndex)
            End If
            If cobBorder.ListIndex >= 0 Then
                .Border = cobBorder.ItemData(cobBorder.ListIndex)
            End If
            .BorderColor = pvUploadColor(cobBorderColor, .BorderColor)
            .BorderSunken = IIf(chkSunken.Value = vbChecked, ucsTriTrue, IIf(chkSunken.Value = vbUnchecked, ucsTriFalse, ucsTri_Undefined))
            .OffsetX = pvStrToLng(txtOffsetX.Text)
            .OffsetY = pvStrToLng(txtOffsetY.Text)
            .Padding = pvStrToLng(txtPadding)
            If cobGradType.ListIndex >= 0 Then
                .BackGradient.GradientType = cobGradType.ItemData(cobGradType.ListIndex)
            End If
            .BackGradient.Color = pvUploadColor(cobGradColor, .BackGradient.Color)
            .BackGradient.SecondColor = pvUploadColor(cobGradSecondColor, .BackGradient.SecondColor)
            If .BackGradient.GradientType = ucsGrdColorOffset Then
                .BackGradient.OffsetHue = Val(txtGradHue.Text)
                .BackGradient.PercentSaturation = Val(txtGradSat.Text)
                .BackGradient.PercentBrightness = Val(txtGradBri.Text)
            Else
                .BackGradient.Alpha = Val(txtGradAlpha.Text)
            End If
            Select Case .BackGradient.GradientType
            Case ucsGrdStretchBitmap, ucsGrdTileBitmap
                Set .BackGradient.Picture = imgBitmap
            Case Else
                Set .BackGradient.Picture = Nothing
            End Select
            If .BackGradient.GradientType = ucsGrdTileBitmap Then
                .BackGradient.TileAbsolutePosition = (chkTileAbsolute.Value = vbChecked)
            End If
        End With
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvFillCombo(ByVal oCombo As ComboBox, sValues As String)
    Const FUNC_NAME     As String = "pvFillCombo"
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    vSplit = Split(sValues, "|")
    With oCombo
        .Clear
        .AddItem STR_UNDEFINED
        .ItemData(.NewIndex) = -1
        For lIdx = 0 To UBound(vSplit) Step 2
            .AddItem vSplit(lIdx + 1) & " - " & vSplit(lIdx)
            .ItemData(.NewIndex) = C2Lng(vSplit(lIdx + 1))
        Next
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvFillColorCombo(ByVal oCombo As ComboBox)
    Const FUNC_NAME     As String = "pvFillColorCombo"
    Dim lIdx            As Long
    
    On Error GoTo EH
    With oCombo
        .Clear
        .AddItem STR_UNDEFINED
        .ItemData(.NewIndex) = -1
'        .AddItem "<Custom...>"
        .AddItem "ScrollBars"
        .ItemData(.NewIndex) = vbScrollBars
        .AddItem "Desktop"
        .ItemData(.NewIndex) = vbDesktop
        .AddItem "ActiveTitleBar"
        .ItemData(.NewIndex) = vbActiveTitleBar
        .AddItem "InactiveTitleBar"
        .ItemData(.NewIndex) = vbInactiveTitleBar
        .AddItem "MenuBar"
        .ItemData(.NewIndex) = vbMenuBar
        .AddItem "WindowBackground"
        .ItemData(.NewIndex) = vbWindowBackground
        .AddItem "WindowFrame"
        .ItemData(.NewIndex) = vbWindowFrame
        .AddItem "MenuText"
        .ItemData(.NewIndex) = vbMenuText
        .AddItem "WindowText"
        .ItemData(.NewIndex) = vbWindowText
        .AddItem "ActiveTitleBarText"
        .ItemData(.NewIndex) = vbActiveTitleBarText
        .AddItem "TitleBarText"
        .ItemData(.NewIndex) = vbTitleBarText
        .AddItem "ActiveBorder"
        .ItemData(.NewIndex) = vbActiveBorder
        .AddItem "InactiveBorder"
        .ItemData(.NewIndex) = vbInactiveBorder
        .AddItem "ApplicationWorkspace"
        .ItemData(.NewIndex) = vbApplicationWorkspace
        .AddItem "Highlight"
        .ItemData(.NewIndex) = vbHighlight
        .AddItem "HighlightText"
        .ItemData(.NewIndex) = vbHighlightText
        .AddItem "ButtonFace / 3DFace"
        .ItemData(.NewIndex) = vbButtonFace
        .AddItem "ButtonShadow / 3DShadow"
        .ItemData(.NewIndex) = vbButtonShadow
        .AddItem "GrayText"
        .ItemData(.NewIndex) = vbGrayText
        .AddItem "ButtonText"
        .ItemData(.NewIndex) = vbButtonText
        .AddItem "InactiveCaptionText"
        .ItemData(.NewIndex) = vbInactiveCaptionText
        .AddItem "InactiveTitleBarText"
        .ItemData(.NewIndex) = vbInactiveTitleBarText
        .AddItem "3DHighlight"
        .ItemData(.NewIndex) = vb3DHighlight
        .AddItem "3DDKShadow"
        .ItemData(.NewIndex) = vb3DDKShadow
        .AddItem "3DLight"
        .ItemData(.NewIndex) = vb3DLight
        .AddItem "InfoText"
        .ItemData(.NewIndex) = vbInfoText
        .AddItem "InfoBackground"
        .ItemData(.NewIndex) = vbInfoBackground
        For lIdx = 1 To .ListCount - 1
            .List(lIdx) = .List(lIdx) & " - &H" & Right(Hex(.ItemData(lIdx)), 2)
        Next
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvLookupCombo( _
            ByVal oCombo As ComboBox, _
            ByVal lValue As Long, _
            Optional bFixColor As Boolean) As Boolean
    Const FUNC_NAME     As String = "pvLookupCombo"
    Dim lIdx            As Long
    
    On Error GoTo EH
    With oCombo
        For lIdx = 0 To .ListCount - 1
            If .ItemData(lIdx) = lValue Then
                .ListIndex = lIdx
                '--- success
                pvLookupCombo = True
                Exit Function
            End If
        Next
        '--- if color combo then need to set custom color
        If bFixColor Then
            .Text = pvColorToText(lValue)
        End If
    End With
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvFillFont(ByVal oText As TextBox, ByVal oFont As StdFont)
    Const FUNC_NAME     As String = "pvFillFont"
    Dim sDesc           As String
    
    On Error GoTo EH
    Set m_oFormat.Font = oFont
    sDesc = m_oFormat.FontDef.Description
    If Len(sDesc) = 0 Then
        sDesc = "None"
    End If
    If Not m_oFormat.ParentFmt Is Nothing Then
        oText = m_oFormat.ParentFmt.FullName & " + " & sDesc
    Else
        oText = sDesc
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvCloneFormat(ByVal oSrc As cFormatDef) As cFormatDef
    Const FUNC_NAME     As String = "pvCloneFormat"
    
    On Error GoTo EH
    Set pvCloneFormat = New cFormatDef
    pvCloneFormat.Contents = oSrc.Contents
    pvCloneFormat.Name = oSrc.Name
    If Not oSrc.ParentFmt Is Nothing Then
        Set pvCloneFormat.ParentFmt = oSrc.ParentFmt
    Else
        Set pvCloneFormat.ParentFont = oSrc.ParentFont
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvShowColorPicker(oCombo As ComboBox)
    Const FUNC_NAME     As String = "pvShowColorPicker"
    Dim clrColor        As OLE_COLOR
    Dim clrNew          As OLE_COLOR
    Dim lIdx            As Long
    
    On Error GoTo EH
    If oCombo.ListIndex > 0 Then
        clrColor = oCombo.ItemData(oCombo.ListIndex)
    Else
        clrColor = pvUploadColor(oCombo)
    End If
    If frmColorPicker.Init(clrColor, clrNew) Then
        '--- confirmed ok
        Call OleTranslateColor(clrColor, 0, clrColor)
        Call OleTranslateColor(clrNew, 0, clrNew)
        If clrColor <> clrNew Then
'            For lIdx = 1 To oCombo.ListCount - 1
'                Call OleTranslateColor(oCombo.ItemData(lIdx), 0, clrColor)
'                If clrColor = clrNew Then
'                    oCombo.ListIndex = lIdx
'                    Exit Sub
'                End If
'            Next
            oCombo.Text = pvColorToText(clrNew)
        End If
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvHex(ByVal lValue As Long, Optional lCount As Long = 2) As String
    On Error Resume Next
    pvHex = Right(String(lCount, "0") & Hex(lValue), lCount)
End Function

Private Function pvColorToText(ByVal clrValue As OLE_COLOR) As String
    Const FUNC_NAME     As String = "pvColorToText"
    Dim rgbColor        As OLERGBQUAD
    
    On Error GoTo EH
    OleTranslateColor clrValue, 0, rgbColor
    pvColorToText = "#" & pvHex(rgbColor.b) & pvHex(rgbColor.G) & pvHex(rgbColor.R)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvNavigateFormat(ByVal lIdx As Long)
    Const FUNC_NAME     As String = "pvNavigateFormat"
    
    On Error GoTo EH
    If lIdx >= LBound(m_vFormatControls) And _
            lIdx <= UBound(m_vFormatControls) Then
        '--- upload format settings
        If Not m_oFormat Is Nothing Then
            pvUploadFormat
        End If
        '--- save format object for upload
        Set m_oFormat = m_vFormatControls(lIdx)
        If Not m_oFormat Is Nothing Then
            '--- load controls
            pvFillControls m_oFormat
            '--- show controls
            picFormat.Visible = True
        Else
            picFormat.Visible = False
        End If
    Else
        '--- ops!!
        picFormat.Visible = False
        Set m_oFormat = Nothing
    End If
    m_lListIndex = lIdx
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvTextGotFocus(oText As TextBox)
    On Error Resume Next
    If oText = STR_UNDEFINED Then
        oText = ""
    End If
End Sub

Private Sub pvTextLostFocus(ByVal oText As TextBox)
    On Error Resume Next
    If Not IsNumeric(Trim(oText)) Then
        oText = "" ' STR_UNDEFINED
    End If
    oText.SelStart = 0
    oText.SelLength = Len(oText)
End Sub

Private Sub pvColorComboClick(ByVal oCombo As ComboBox, oCmd As CommandButton)
    On Error Resume Next
    If Not m_bInSet Then Changed = True
    If oCombo.ListIndex = 0 Then
        oCmd.BackColor = vbButtonFace
    Else
        oCmd.BackColor = pvUploadColor(oCombo)
    End If
    UpdateWindow oCmd.hwnd
End Sub

Private Sub pvCenterBitmap()
    On Error Resume Next
    imgBitmap.Move (picBitmap.ScaleWidth - imgBitmap.Width) \ 2, (picBitmap.ScaleHeight - imgBitmap.Height) \ 2
    picBitmap.Visible = False
    picBitmap.Visible = True
End Sub

Private Sub cmdOpen_Click()
    Const FUNC_NAME     As String = "cmdOpen_Click"
    Dim vDef            As Variant
    Dim oFmt            As cFormatDef
    Dim nFile           As Integer
    Dim bArray()        As Byte
    
    On Error GoTo EH_Cancel
    comDlgFile.Flags = cdlOFNHideReadOnly
    comDlgFile.Filter = STR_FILTER_OBF
    comDlgFile.ShowOpen
    On Error GoTo EH
    '--- load file
    nFile = FreeFile
    Open comDlgFile.FileName For Binary As #nFile
    ReDim bArray(0 To LOF(nFile) - 1)
    Get #nFile, , bArray
    Close #nFile
    '--- deserialize from propbag
    With New PropertyBag
        .Contents = bArray
        For Each vDef In m_vFormatControls
            Set oFmt = vDef
            oFmt.Contents = .ReadProperty(Replace(oFmt.Name, " ", ""), oFmt.Contents)
        Next
    End With
    '--- fill UI
    pvFillControls m_oFormat
    Changed = True
EH_Cancel:
    Exit Sub
EH:
    PushError FUNC_NAME, MODULE_NAME
    If nFile <> 0 Then
        Close #nFile
    End If
    Select Case PopShowError(CAP_MSG)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub cmdSave_Click()
    Const FUNC_NAME     As String = "cmdSave_Click"
    Dim vDef            As Variant
    Dim oFmt            As cFormatDef
    Dim nFile           As Integer
    Dim bArray()        As Byte
    
    On Error GoTo EH_Cancel
    comDlgFile.Flags = cdlOFNOverwritePrompt
    comDlgFile.Filter = STR_FILTER_OBF
    comDlgFile.ShowSave
    On Error GoTo EH
    '--- update current format from UI
    pvUploadFormat
    '--- persist in probbag
    With New PropertyBag
        For Each vDef In m_vFormatControls
            Set oFmt = vDef
            Call .WriteProperty(Replace(oFmt.Name, " ", ""), oFmt.Contents)
        Next
        bArray = .Contents
    End With
    '--- store in file
    nFile = FreeFile
    Open comDlgFile.FileName For Binary As #nFile
    Put #nFile, , bArray
    Close #nFile
EH_Cancel:
    Exit Sub
EH:
    PushError FUNC_NAME, MODULE_NAME
    If nFile <> 0 Then
        Close #nFile
    End If
    Select Case PopShowError(CAP_MSG)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub UserControl_Initialize()
    On Error Resume Next
    Set m_oFont = New StdFont
    Set m_oFormat = New cFormatDef
End Sub

Private Sub pvReload()
    picFormat.Visible = False
    pvFillCombo cobBorder, STR_BORDER
    pvFillCombo cobGradType, STR_GRADIENT
    pvFillCombo cobHorAlignment, STR_ALIGNMENT_HOR
    pvFillCombo cobVertAlignment, STR_ALIGNMENT_VERT
    pvFillColorCombo cobForeColor
    pvFillColorCombo cobBorderColor
    pvFillColorCombo cobGradColor
    pvFillColorCombo cobGradSecondColor
    m_vFormatControls = Array( _
            pvCloneFormat(m_oControl.FormatControl), _
            pvCloneFormat(m_oControl.FormatGroup), _
            pvCloneFormat(m_oControl.FormatGroupHover), _
            pvCloneFormat(m_oControl.FormatGroupPressed), _
            pvCloneFormat(m_oControl.FormatGroupSelected), _
            pvCloneFormat(m_oControl.FormatItem), _
            pvCloneFormat(m_oControl.FormatItemHover), _
            pvCloneFormat(m_oControl.FormatItemPressed), _
            pvCloneFormat(m_oControl.FormatItemSelected), _
            pvCloneFormat(m_oControl.FormatItemLargeIcons), _
            pvCloneFormat(m_oControl.FormatSmallIcon), _
            pvCloneFormat(m_oControl.FormatSmallIconHover), _
            pvCloneFormat(m_oControl.FormatSmallIconPressed), _
            pvCloneFormat(m_oControl.FormatSmallIconSelected), _
            pvCloneFormat(m_oControl.FormatLargeIcon), _
            pvCloneFormat(m_oControl.FormatLargeIconHover), _
            pvCloneFormat(m_oControl.FormatLargeIconPressed), _
            pvCloneFormat(m_oControl.FormatLargeIconSelected))
    pvFillListBox
End Sub

Public Sub SelectionChanged(SelectedControls As Object)
    Const FUNC_NAME     As String = "SelectionChanged"

    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    If Not m_oControl Is SelectedControls(0) Then
        Set m_oControl = SelectedControls(0)
        pvReload
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Sub ApplyChanges()
    Const FUNC_NAME     As String = "ApplyChanges"
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- upload current format settings if modified
    Screen.MousePointer = vbHourglass
    If Not m_oFormat Is Nothing Then
        pvUploadFormat
        pvFillControls m_oFormat
    End If
    If Not m_oControl Is Nothing Then
        With m_oControl
            '--- persist in controls props (source easy to edit/add formats)
            Set .FormatControl = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatGroup = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatGroupHover = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatGroupPressed = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatGroupSelected = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatItem = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatItemHover = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatItemPressed = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatItemSelected = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatItemLargeIcons = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatSmallIcon = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatSmallIconHover = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatSmallIconPressed = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatSmallIconSelected = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatLargeIcon = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatLargeIconHover = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatLargeIconPressed = m_vFormatControls(lIdx): lIdx = lIdx + 1
            Set .FormatLargeIconSelected = m_vFormatControls(lIdx): lIdx = lIdx + 1
            '--- force reload on SelectionChanged
            pvReload
        End With
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub cmdForeColor_Click()
    pvShowColorPicker cobForeColor
End Sub

Private Sub cmdBorderColor_Click()
    pvShowColorPicker cobBorderColor
End Sub

Private Sub cmdGradColor_Click()
    pvShowColorPicker cobGradColor
End Sub

Private Sub cmdGradSecondColor_Click()
    pvShowColorPicker cobGradSecondColor
End Sub

Private Sub cmdFont_Click()
    Const FUNC_NAME     As String = "cmdFont_Click"

    On Error GoTo EH
    With m_oFont
        comDlg.FontBold = .Bold
'        .Charset '--- !!! MISSING !!!
        comDlg.FontItalic = .Italic
        comDlg.FontName = .Name
        comDlg.FontSize = .Size
        comDlg.FontStrikethru = .Strikethrough
        comDlg.FontUnderline = .Underline
'        .Weight '--- !!! MISSING !!!
    End With
    On Error GoTo EH_Cancel
    comDlg.ShowFont
    On Error GoTo EH
    With m_oFont
        .Bold = comDlg.FontBold
'        .Charset '--- !!! MISSING !!!
        .Italic = comDlg.FontItalic
        .Name = comDlg.FontName
        .Size = comDlg.FontSize
        .Strikethrough = comDlg.FontStrikethru
        .Underline = comDlg.FontUnderline
'        .Weight '--- !!! MISSING !!!
    End With
    pvFillFont txtFont, m_oFont
EH_Cancel:
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub chkSunken_Click()
    If Not m_bInSet Then
        If chkSunken.Value = vbUnchecked And chkSunken.Tag = "" Then
            chkSunken.Value = vbGrayed
        End If
        Changed = True
    End If
    chkSunken.Tag = IIf(chkSunken.Value = vbGrayed, "grey", "")
End Sub

Private Sub cobBorder_Click()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub cobForeColor_Click()
    pvColorComboClick cobForeColor, cmdForeColor
End Sub

Private Sub cobForeColor_Change()
    cobForeColor_Click
End Sub

Private Sub cobBorderColor_Click()
    pvColorComboClick cobBorderColor, cmdBorderColor
End Sub

Private Sub cobBorderColor_Change()
    cobBorderColor_Click
End Sub

Private Sub cobGradColor_Click()
    pvColorComboClick cobGradColor, cmdGradColor
End Sub

Private Sub cobGradColor_Change()
    cobGradColor_Click
End Sub

Private Sub cobGradSecondColor_Click()
    pvColorComboClick cobGradSecondColor, cmdGradSecondColor
End Sub

Private Sub cobGradSecondColor_Change()
    cobGradSecondColor_Click
End Sub

Private Sub cobGradType_Click()
    Const FUNC_NAME     As String = "cobGradType_Click"
    Dim lNumColors      As Long
    Dim eItemData       As UcsGradientType
    
    On Error GoTo EH
    Select Case cobGradType.ItemData(cobGradType.ListIndex)
    Case ucsGrdSolid, ucsGrdColorOffset
        lNumColors = 1
    Case ucsGrdHorizontal, ucsGrdVertical, ucsGrdBlend, ucsGrdAlphaBlend
        lNumColors = 2
    End Select
    labGradColor.Visible = (lNumColors > 0)
    cobGradColor.Visible = (lNumColors > 0)
    cmdGradColor.Visible = (lNumColors > 0)
    labGradSecondColor.Visible = (lNumColors > 1)
    cobGradSecondColor.Visible = (lNumColors > 1)
    cmdGradSecondColor.Visible = (lNumColors > 1)
    eItemData = cobGradType.ItemData(cobGradType.ListIndex)
    picColorOffset.Visible = (eItemData = ucsGrdColorOffset)
    txtGradAlpha.Visible = (eItemData = ucsGrdAlphaBlend)
    labGradAlpha.Visible = (eItemData = ucsGrdAlphaBlend)
    picSelectBitmap.Visible = (eItemData = ucsGrdStretchBitmap Or eItemData = ucsGrdTileBitmap)
    chkTileAbsolute.Visible = (eItemData = ucsGrdTileBitmap)
    If Not m_bInSet Then Changed = True
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub cobHorAlignment_Click()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub cobVertAlignment_Click()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    If Not m_bInSet Then Changed = True
End Sub

Private Sub txtOffsetX_Change()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub txtOffsetY_Change()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub txtPadding_Change()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub txtGradBri_Change()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub txtGradHue_Change()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub txtGradSat_Change()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub txtGradAlpha_Change()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub chkTileAbsolute_Click()
    If Not m_bInSet Then Changed = True
End Sub

Private Sub trvFormats_NodeClick(ByVal Node As MSComctlLib.Node)
    Const FUNC_NAME     As String = "trvFormats_NodeClick"

    On Error GoTo EH
'    trvFormats.Drag vbBeginDrag
    Set trvFormats.DropHighlight = trvFormats.SelectedItem
'    trvFormats.Drag vbCancel
    pvNavigateFormat Node.Index - 1
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub txtOffsetX_GotFocus()
    pvTextGotFocus txtOffsetX
End Sub

Private Sub txtOffsetX_LostFocus()
    pvTextLostFocus txtOffsetX
End Sub

Private Sub txtOffsetY_GotFocus()
    pvTextGotFocus txtOffsetY
End Sub

Private Sub txtOffsetY_LostFocus()
    pvTextLostFocus txtOffsetY
End Sub

Private Sub txtPadding_GotFocus()
    pvTextGotFocus txtPadding
End Sub

Private Sub txtPadding_LostFocus()
    pvTextLostFocus txtPadding
End Sub

Private Sub trvFormats_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "trvFormats_OLEDragDrop"
    Dim oFmt            As cFormatDef
    
    On Error GoTo EH
    Set trvFormats.DropHighlight = trvFormats.HitTest(X, Y)
    If Not trvFormats.DropHighlight Is trvFormats.SelectedItem Then
        If MsgBox("Copy " & trvFormats.SelectedItem.Key & " to " & trvFormats.DropHighlight.Key & "?", vbQuestion + vbYesNo, CAP_MSG) = vbYes Then
            Set oFmt = m_vFormatControls(trvFormats.DropHighlight.Index - 1)
            oFmt.Contents = m_oFormat.Contents
            Changed = True
            Set trvFormats.SelectedItem = trvFormats.DropHighlight
            pvNavigateFormat trvFormats.DropHighlight.Index - 1
        Else
            Set trvFormats.DropHighlight = trvFormats.SelectedItem
        End If
    End If
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub trvFormats_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    On Error Resume Next
    Set trvFormats.DropHighlight = trvFormats.HitTest(X, Y)
End Sub

Private Sub imgBitmap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    m_bDrag = True
    m_sX = X: m_sY = Y
End Sub

Private Sub imgBitmap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If m_bDrag Then
        With imgBitmap
            .Move .Left + (X - m_sX), .Top + (Y - m_sY)
        End With
    End If
End Sub

Private Sub imgBitmap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    imgBitmap_MouseMove Button, Shift, X, Y
    With imgBitmap
        If .Width < picBitmap.ScaleWidth Then
            .Left = (picBitmap.ScaleWidth - .Width) \ 2
        Else
            If .Left < picBitmap.ScaleWidth - .Width Then
                .Left = picBitmap.ScaleWidth - .Width
            End If
            If .Left > 0 Then
                .Left = 0
            End If
        End If
        If .Height < picBitmap.ScaleHeight Then
            .Top = (picBitmap.ScaleHeight - .Height) \ 2
        Else
            If .Top < picBitmap.ScaleHeight - .Height Then
                .Top = picBitmap.ScaleHeight - .Height
            End If
            If .Top > 0 Then
                .Top = 0
            End If
        End If
    End With
    m_bDrag = False
End Sub

Private Sub cmdBitmapIcon_Click()
    On Error GoTo EH_Cancel
    comDlgFile.Flags = cdlOFNHideReadOnly
    comDlgFile.Filter = STR_FILTER_IMAGES
    comDlgFile.ShowOpen
    Set imgBitmap = LoadPicture(comDlgFile.FileName)
    Changed = True
    DoEvents
    pvCenterBitmap
EH_Cancel:
End Sub

Private Sub cmdBitmapClear_Click()
    On Error Resume Next
    Set imgBitmap = Nothing
    Changed = True
End Sub

