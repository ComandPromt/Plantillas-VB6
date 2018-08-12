VERSION 5.00
Begin VB.Form frmColorPicker 
   AutoRedraw      =   -1  'True
   Caption         =   "Color Picker"
   ClientHeight    =   4344
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   7392
   Icon            =   "frmColorPicker.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrKeyboard 
      Interval        =   500
      Left            =   6216
      Top             =   840
   End
   Begin VB.Frame fraColors 
      Caption         =   "Reference"
      Height          =   915
      Left            =   4560
      TabIndex        =   28
      Top             =   60
      Width           =   1515
      Begin VB.PictureBox picReference 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   180
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   852
         Begin VB.Label labOld 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   792
         End
         Begin VB.Label labNew 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   -180
            TabIndex        =   30
            Top             =   240
            Width           =   795
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   312
      Left            =   6180
      TabIndex        =   23
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   312
      Left            =   6180
      TabIndex        =   22
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame fraNumeric 
      Caption         =   "Components"
      Height          =   3015
      Left            =   4560
      TabIndex        =   24
      Top             =   1080
      Width           =   2715
      Begin VB.PictureBox picAdditional 
         BorderStyle     =   0  'None
         Height          =   1512
         Left            =   60
         ScaleHeight     =   1512
         ScaleWidth      =   2592
         TabIndex        =   32
         Top             =   1440
         Width           =   2592
         Begin VB.TextBox txtNewColor 
            Alignment       =   2  'Center
            Height          =   252
            Left            =   540
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtLabLuminance 
            Alignment       =   1  'Right Justify
            Height          =   252
            Left            =   540
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   14
            Top             =   60
            Width           =   615
         End
         Begin VB.TextBox txtLabA 
            Alignment       =   1  'Right Justify
            Height          =   252
            Left            =   540
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   15
            Top             =   420
            Width           =   615
         End
         Begin VB.TextBox txtLabB 
            Alignment       =   1  'Right Justify
            Height          =   252
            Left            =   540
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   16
            Top             =   780
            Width           =   615
         End
         Begin VB.TextBox txtCyan 
            Alignment       =   1  'Right Justify
            Height          =   252
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   18
            Top             =   60
            Width           =   615
         End
         Begin VB.TextBox txtMagenta 
            Alignment       =   1  'Right Justify
            Height          =   252
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   19
            Top             =   420
            Width           =   615
         End
         Begin VB.TextBox txtYellow 
            Alignment       =   1  'Right Justify
            Height          =   252
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   20
            Top             =   780
            Width           =   615
         End
         Begin VB.TextBox txtBlack 
            Alignment       =   1  'Right Justify
            Height          =   252
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   21
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "#"
            Height          =   252
            Left            =   300
            TabIndex        =   40
            Top             =   1200
            Width           =   252
         End
         Begin VB.Label Label4 
            Caption         =   "L:"
            Height          =   252
            Left            =   300
            TabIndex        =   39
            Top             =   60
            Width           =   372
         End
         Begin VB.Label Label5 
            Caption         =   "a:"
            Height          =   252
            Left            =   300
            TabIndex        =   38
            Top             =   420
            Width           =   372
         End
         Begin VB.Label Label6 
            Caption         =   "b:"
            Height          =   252
            Left            =   300
            TabIndex        =   37
            Top             =   780
            Width           =   372
         End
         Begin VB.Label Label7 
            Caption         =   "C:"
            Height          =   252
            Left            =   1680
            TabIndex        =   36
            Top             =   60
            Width           =   372
         End
         Begin VB.Label Label8 
            Caption         =   "M:"
            Height          =   252
            Left            =   1680
            TabIndex        =   35
            Top             =   420
            Width           =   372
         End
         Begin VB.Label Label9 
            Caption         =   "Y:"
            Height          =   252
            Left            =   1680
            TabIndex        =   34
            Top             =   780
            Width           =   372
         End
         Begin VB.Label Label10 
            Caption         =   "K:"
            Height          =   252
            Left            =   1680
            TabIndex        =   33
            Top             =   1140
            Width           =   372
         End
      End
      Begin VB.TextBox txtBlue 
         Alignment       =   1  'Right Justify
         Height          =   252
         Left            =   1980
         MaxLength       =   6
         TabIndex        =   13
         Top             =   1080
         Width           =   432
      End
      Begin VB.TextBox txtGreen 
         Alignment       =   1  'Right Justify
         Height          =   252
         Left            =   1980
         MaxLength       =   6
         TabIndex        =   12
         Top             =   720
         Width           =   432
      End
      Begin VB.TextBox txtRed 
         Alignment       =   1  'Right Justify
         Height          =   252
         Left            =   1980
         MaxLength       =   6
         TabIndex        =   11
         Top             =   360
         Width           =   432
      End
      Begin VB.TextBox txtBri 
         Alignment       =   1  'Right Justify
         Height          =   252
         Left            =   600
         MaxLength       =   6
         TabIndex        =   10
         Top             =   1080
         Width           =   432
      End
      Begin VB.TextBox txtSat 
         Alignment       =   1  'Right Justify
         Height          =   252
         Left            =   600
         MaxLength       =   6
         TabIndex        =   9
         Top             =   720
         Width           =   432
      End
      Begin VB.TextBox txtHue 
         Alignment       =   1  'Right Justify
         Height          =   252
         Left            =   600
         MaxLength       =   6
         TabIndex        =   8
         Top             =   360
         Width           =   432
      End
      Begin VB.OptionButton optGreen 
         Caption         =   "&G"
         Height          =   252
         Left            =   1500
         TabIndex        =   6
         Top             =   720
         Width           =   492
      End
      Begin VB.OptionButton optBlue 
         Caption         =   "&B"
         Height          =   252
         Left            =   1500
         TabIndex        =   7
         Top             =   1080
         Width           =   492
      End
      Begin VB.OptionButton optRed 
         Caption         =   "&R"
         Height          =   252
         Left            =   1500
         TabIndex        =   5
         Top             =   360
         Width           =   492
      End
      Begin VB.OptionButton optBri 
         Caption         =   "&B"
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   492
      End
      Begin VB.OptionButton optSat 
         Caption         =   "&S"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   492
      End
      Begin VB.OptionButton optHue 
         Caption         =   "&H"
         Height          =   252
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   492
      End
      Begin OBPropPages.ctxUpDown updBlue 
         Height          =   264
         Left            =   2400
         Top             =   1080
         Width           =   192
         _ExtentX        =   339
         _ExtentY        =   466
      End
      Begin OBPropPages.ctxUpDown updGreen 
         Height          =   264
         Left            =   2400
         Top             =   720
         Width           =   192
         _ExtentX        =   339
         _ExtentY        =   466
      End
      Begin OBPropPages.ctxUpDown updRed 
         Height          =   264
         Left            =   2400
         Top             =   360
         Width           =   192
         _ExtentX        =   339
         _ExtentY        =   466
      End
      Begin OBPropPages.ctxUpDown updBri 
         Height          =   255
         Left            =   1020
         Top             =   1080
         Width           =   192
         _ExtentX        =   339
         _ExtentY        =   445
      End
      Begin OBPropPages.ctxUpDown updSat 
         Height          =   255
         Left            =   1020
         Top             =   720
         Width           =   192
         _ExtentX        =   339
         _ExtentY        =   445
      End
      Begin OBPropPages.ctxUpDown updHue 
         Height          =   255
         Left            =   1020
         Top             =   360
         Width           =   192
         _ExtentX        =   339
         _ExtentY        =   445
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Left            =   1260
         TabIndex        =   27
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   255
         Left            =   1260
         TabIndex        =   26
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblH 
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   25
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CheckBox chkWebSafe 
      Caption         =   "Web colors only"
      Height          =   264
      Left            =   84
      TabIndex        =   0
      Top             =   4020
      Width           =   1632
   End
   Begin VB.CheckBox chkAccel 
      Caption         =   "Accelerate display"
      Enabled         =   0   'False
      Height          =   252
      Left            =   1800
      TabIndex        =   1
      Top             =   4020
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1776
   End
   Begin VB.Image imgBar 
      Height          =   3870
      Left            =   4080
      MousePointer    =   2  'Cross
      Top             =   60
      Width           =   240
   End
   Begin VB.Image imgRect 
      Height          =   3870
      Left            =   75
      MousePointer    =   2  'Cross
      Top             =   75
      Width           =   3870
   End
End
Attribute VB_Name = "frmColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

'--- set to 1 NOT to compile accelerated gradients for win98/2k
#Const NO_ACCELERATED_GRADIENTS = 0

'=========================================================================
' UDTs and Enums
'=========================================================================

Private Enum UcsRgbColorIdx
    ucsRgbRed
    ucsRgbGreen
    ucsRgbBlue
End Enum

Private Type UcsHsbColor
    Hue                 As Double
    Sat                 As Double
    Bri                 As Double
End Type

Private Type UcsXyzColor
    X                   As Double
    Y                   As Double
    Z                   As Double
End Type

Private Type UcsLabColor
    L                   As Double
    a                   As Double
    b                   As Double
End Type

Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    b                   As Byte
    a                   As Byte
End Type

Private Type UcsRgbTriple
    b                   As Byte
    G                   As Byte
    R                   As Byte
End Type

Private Type UcsColorGraphicsCache
    imgRect             As StdPicture
    imgBar              As StdPicture
    bWebSafe            As Boolean
    rgbColor            As UcsRgbQuad
    hsbColor            As UcsHsbColor
End Type

'=========================================================================
' API
'=========================================================================

'--- for GetSystemMetrics
Private Const SM_CYCAPTION              As Long = 4
Private Const SM_CYDLGFRAME             As Long = 8
Private Const SM_CXDLGFRAME             As Long = 7
'--- for SetStretchBltMode
Private Const HALFTONE                  As Long = 4
'--- for GradientFill
Private Const GRADIENT_FILL_RECT_H      As Long = 0
Private Const GRADIENT_FILL_RECT_V      As Long = 1
Private Const GRADIENT_FILL_TRIANGLE    As Long = 2

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Any) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GradientFill Lib "Msimg32.dll" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GradientFillRect Lib "Msimg32.dll" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1             As Long
    Vertex2             As Long
    Vertex3             As Long
End Type

Private Type GRADIENT_RECT
    UpperLeft           As Long
    LowerRight          As Long
End Type

Private Type TRIVERTEX
    X                   As Long
    Y                   As Long
    Red                 As Integer
    Green               As Integer
    Blue                As Integer
    Alpha               As Integer
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

'--- integer math precision constants
Private Const PREC_BRI              As Long = 8
Private Const PREC_BRI_255          As Long = PREC_BRI * 255
Private Const PREC_SAT              As Long = 8
Private Const PREC_SAT_255          As Long = PREC_SAT * 255
Private Const PREC_SAT_BRI_255      As Long = PREC_SAT * PREC_BRI * 255
Private Const PREC_HUE              As Long = 256
Private Const PREC_HUE_255          As Long = PREC_HUE * 255
Private Const PREC_HUE_SAT_255      As Long = PREC_HUE * PREC_SAT * 255
Private Const PREC_HUE_BRI_255      As Long = PREC_HUE * PREC_BRI * 255
Private Const PREC_HUE_SAT_BRI_255  As Long = PREC_HUE * PREC_SAT * PREC_BRI * 255
'--- color rect and color bar sizes
Private Const BAR_WIDTH             As Long = 16
'--- these used to be constants
Private RECT_WIDTH_STEP             As Long ' = 23
Private RECT_WIDTH                  As Long ' = 6 * RECT_WIDTH_STEP ' 258
Private RECT_HEIGHT                 As Long
Private BAR_HEIGHT                  As Long ' = RECT_HEIGHT
'--- keyboard input (timer) type
Private Const STR_TIMER_FROM_RGB    As String = "rgb"
Private Const STR_TIMER_FROM_HSB    As String = "hsb"
'--- misc
Private Const MASK_COLOR            As Long = &HFF00FF
Private Const GRID_SIZE             As Long = 5
Private Const LAB_CORELDRAW_NORMALIZE As Double = 2

Private m_bOk                   As Boolean
Private m_clrCurrent            As OLE_COLOR
Private m_clrOriginal           As OLE_COLOR
Private m_hsbCurrent            As UcsHsbColor
Private m_hsbPrevious           As UcsHsbColor
Private m_oHueCache             As UcsColorGraphicsCache
Private m_oSaturationCache      As UcsColorGraphicsCache
Private m_oBrightnessCache      As UcsColorGraphicsCache
Private m_oRedCache             As UcsColorGraphicsCache
Private m_oGreenCache           As UcsColorGraphicsCache
Private m_oBlueCache            As UcsColorGraphicsCache
Private m_imgRect               As StdPicture
Private m_imgBar                As StdPicture
Private m_aWebSafe(0 To 255)    As Byte
Private m_bWebSafeOnly          As Boolean
Private m_bBarPressed           As Boolean
Private m_bRectPressed          As Boolean
Private m_bInSet                As Boolean
Private m_imgBarSelector        As StdPicture
Private m_bAccelerateSupported  As Boolean
Private m_dblTimer              As Double
Private m_sNumericHeight        As Single

'=========================================================================
' Properties
'=========================================================================

Property Get Color() As OLE_COLOR
    Color = m_clrCurrent
End Property

Property Let Color(ByVal clrValue As OLE_COLOR)
    Dim rgbColor        As UcsRgbQuad
    Dim cmykColor       As UcsRgbQuad
    Dim labColor        As UcsHsbColor
    
    '--- do web colors conversion
    CopyMemory rgbColor, clrValue, 4
    If m_bWebSafeOnly Then
        With rgbColor
            .R = m_aWebSafe(.R)
            .G = m_aWebSafe(.G)
            .b = m_aWebSafe(.b)
        End With
        CopyMemory clrValue, rgbColor, 4
    End If
    '--- if anything changed
    If clrValue <> m_clrCurrent _
            Or Not pvIsEqualHsb(m_hsbPrevious, m_hsbCurrent) Then
        '--- save current color (and hsb representation)
        m_clrCurrent = clrValue
        m_hsbPrevious = m_hsbCurrent
        '--- modify UI
        labNew.BackColor = clrValue
        '--- prevent textbox's events from controling color
        m_bInSet = True
        '--- RGB
        With rgbColor
            pvSetText txtRed, .R
            pvSetText txtGreen, .G
            pvSetText txtBlue, .b
        End With
        '--- RGB -> HSB
        With m_hsbCurrent
            If .Hue < 0 Then
                m_hsbCurrent = pvRGBToHSB(clrValue)
            End If
            pvSetText txtHue, CLng(.Hue)
            pvSetText txtSat, CLng(.Sat)
            pvSetText txtBri, CLng(.Bri)
        End With
        '--- RGB -> CMYK
        With pvRGBToCMYK(clrValue)
            pvSetText txtCyan, .R
            pvSetText txtMagenta, .G
            pvSetText txtYellow, .b
            pvSetText txtBlack, .a
        End With
        '--- RGB -> XYZ -> L*a*b*
        With pvXYZToLAB(pvRGBToXYZ(clrValue))
            pvSetText txtLabLuminance, Format(.L, "0.0")
            pvSetText txtLabA, Format(.a, "0.0")
            pvSetText txtLabB, Format(.b, "0.0")
        End With
        '--- RGB -> Web
        pvSetText txtNewColor, pvHex(rgbColor.R) & pvHex(rgbColor.G) & pvHex(rgbColor.b)
        '--- end of prevention
        m_bInSet = False
        '--- set current graphics depending on current view
        If optHue Then
            pvSetHueCurrent pvInitHsb(m_hsbCurrent.Hue, 100, 100), m_bWebSafeOnly
        ElseIf optSat Then
            pvSetSaturationCurrent m_hsbCurrent, m_bWebSafeOnly
        ElseIf optBri Then
            pvSetBrightnessCurrent m_hsbCurrent, m_bWebSafeOnly
        ElseIf optRed Then
            pvSetRedCurrent rgbColor, m_bWebSafeOnly
        ElseIf optGreen Then
            pvSetGreenCurrent rgbColor, m_bWebSafeOnly
        ElseIf optBlue Then
            pvSetBlueCurrent rgbColor, m_bWebSafeOnly
        End If
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init( _
            ByVal clrColor As OLE_COLOR, _
            clrNew As OLE_COLOR) As Boolean
'--- retval: true - confirmed and clrNew is the new color, false - canceled
    On Error GoTo EH
    '--- translate input color
    If clrColor = -1 Then
        clrColor = 0
    Else
        OleTranslateColor clrColor, 0, clrColor
    End If
    '--- local vars
    m_hsbCurrent.Hue = -1
    labOld.BackColor = clrColor
    m_clrOriginal = clrColor
    Color = clrColor
    '--- UI handling
    pvRefresh
    m_bOk = False
    Show vbModal
    If m_bOk Then
        '--- confirmed ok
        clrNew = Color
        '--- success
        Init = True
    End If
    Exit Function
EH:
    MsgBox "Error: " & Error, vbCritical, Me.Caption
End Function

Private Sub pvSetHueCurrent( _
            hsbColor As UcsHsbColor, _
            ByVal bWebSafe As Boolean)
    With m_oHueCache
        If Not pvIsEqualHsb(.hsbColor, hsbColor) _
                Or .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgRect, RECT_WIDTH, RECT_HEIGHT) Then
#If NO_ACCELERATED_GRADIENTS = 0 Then
            If Not bWebSafe And m_bAccelerateSupported Then
                Set .imgRect = pvCreateRectHueAccel(hsbColor)
            Else
                Set .imgRect = pvCreateRectHue(hsbColor)
            End If
#Else
            Set .imgRect = pvCreateRectHue(hsbColor)
#End If
        End If
        If .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgBar, BAR_WIDTH, BAR_HEIGHT) Then
            Set .imgBar = pvCreateBarHue()
        End If
        .hsbColor = hsbColor
        .bWebSafe = m_bWebSafeOnly
        Set m_imgRect = .imgRect
        Set m_imgBar = .imgBar
    End With
End Sub

Private Sub pvSetSaturationCurrent( _
            hsbColor As UcsHsbColor, _
            ByVal bWebSafe As Boolean)
    With m_oSaturationCache
        If .hsbColor.Sat <> hsbColor.Sat _
                Or .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgRect, RECT_WIDTH, RECT_HEIGHT) Then
#If NO_ACCELERATED_GRADIENTS = 0 Then
            If Not bWebSafe And m_bAccelerateSupported Then
                Set .imgRect = pvCreateRectSaturationAccel(hsbColor.Sat)
            Else
                Set .imgRect = pvCreateRectSaturation(hsbColor.Sat)
            End If
#Else
            Set .imgRect = pvCreateRectSaturation(hsbColor.Sat)
#End If
        End If
        If .hsbColor.Hue <> hsbColor.Hue _
                Or .hsbColor.Bri <> hsbColor.Bri _
                Or .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgBar, BAR_WIDTH, BAR_HEIGHT) Then
            Set .imgBar = pvCreateBarSaturation(hsbColor.Hue, hsbColor.Bri)
        End If
        .hsbColor = hsbColor
        .bWebSafe = m_bWebSafeOnly
        Set m_imgRect = .imgRect
        Set m_imgBar = .imgBar
    End With
End Sub

Private Sub pvSetBrightnessCurrent( _
            hsbColor As UcsHsbColor, _
            ByVal bWebSafe As Boolean)
    With m_oBrightnessCache
        If .hsbColor.Bri <> hsbColor.Bri _
                Or .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgRect, RECT_WIDTH, RECT_HEIGHT) Then
#If NO_ACCELERATED_GRADIENTS = 0 Then
            If Not bWebSafe And m_bAccelerateSupported Then
                Set .imgRect = pvCreateRectBrightnessAccel(hsbColor.Bri)
            Else
                Set .imgRect = pvCreateRectBrightness(hsbColor.Bri)
            End If
#Else
            Set .imgRect = pvCreateRectBrightness(hsbColor.Bri)
#End If
        End If
        If .hsbColor.Hue <> hsbColor.Hue _
                Or .hsbColor.Sat <> hsbColor.Sat _
                Or .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgBar, BAR_WIDTH, BAR_HEIGHT) Then
            Set .imgBar = pvCreateBarBrightness(hsbColor.Hue, hsbColor.Sat)
        End If
        .hsbColor = hsbColor
        .bWebSafe = m_bWebSafeOnly
        Set m_imgRect = .imgRect
        Set m_imgBar = .imgBar
    End With
End Sub

Private Sub pvSetRedCurrent( _
            rgbColor As UcsRgbQuad, _
            ByVal bWebSafe As Boolean)
    With m_oRedCache
        If .rgbColor.R <> rgbColor.R _
                Or .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgRect, RECT_WIDTH, RECT_HEIGHT) Then
#If NO_ACCELERATED_GRADIENTS = 0 Then
            If Not bWebSafe And m_bAccelerateSupported Then
                Set .imgRect = pvCreateRectRGBAccel(rgbColor.R, ucsRgbRed)
            Else
                Set .imgRect = pvCreateRectRGB(rgbColor.R, ucsRgbRed)
            End If
#Else
            Set .imgRect = pvCreateRectRGB(rgbColor.R, ucsRgbRed)
#End If
        End If
        If .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgBar, BAR_WIDTH, BAR_HEIGHT) Then
            Set .imgBar = pvCreateBarRGB(ucsRgbRed)
        End If
        .rgbColor = rgbColor
        .bWebSafe = m_bWebSafeOnly
        Set m_imgRect = .imgRect
        Set m_imgBar = .imgBar
    End With
End Sub

Private Sub pvSetGreenCurrent( _
            rgbColor As UcsRgbQuad, _
            ByVal bWebSafe As Boolean)
    With m_oGreenCache
        If .rgbColor.G <> rgbColor.G _
                Or .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgRect, RECT_WIDTH, RECT_HEIGHT) Then
#If NO_ACCELERATED_GRADIENTS = 0 Then
            If Not bWebSafe And m_bAccelerateSupported Then
                Set .imgRect = pvCreateRectRGBAccel(rgbColor.G, ucsRgbGreen)
            Else
                Set .imgRect = pvCreateRectRGB(rgbColor.G, ucsRgbGreen)
            End If
#Else
            Set .imgRect = pvCreateRectRGB(rgbColor.G, ucsRgbGreen)
#End If
        End If
        If .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgBar, BAR_WIDTH, BAR_HEIGHT) Then
            Set .imgBar = pvCreateBarRGB(ucsRgbGreen)
        End If
        .rgbColor = rgbColor
        .bWebSafe = m_bWebSafeOnly
        Set m_imgRect = .imgRect
        Set m_imgBar = .imgBar
    End With
End Sub

Private Sub pvSetBlueCurrent( _
            rgbColor As UcsRgbQuad, _
            ByVal bWebSafe As Boolean)
    With m_oBlueCache
        If .rgbColor.b <> rgbColor.b _
                Or .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgRect, RECT_WIDTH, RECT_HEIGHT) Then
#If NO_ACCELERATED_GRADIENTS = 0 Then
            If Not bWebSafe And m_bAccelerateSupported Then
                Set .imgRect = pvCreateRectRGBAccel(rgbColor.b, ucsRgbBlue)
            Else
                Set .imgRect = pvCreateRectRGB(rgbColor.b, ucsRgbBlue)
            End If
#Else
            Set .imgRect = pvCreateRectRGB(rgbColor.b, ucsRgbBlue)
#End If
        End If
        If .bWebSafe <> bWebSafe _
                Or Not pvCheckDimensions(.imgBar, BAR_WIDTH, BAR_HEIGHT) Then
            Set .imgBar = pvCreateBarRGB(ucsRgbBlue)
        End If
        .rgbColor = rgbColor
        .bWebSafe = m_bWebSafeOnly
        Set m_imgRect = .imgRect
        Set m_imgBar = .imgBar
    End With
End Sub

Private Function pvCreateRectHue(hsbColor As UcsHsbColor) As StdPicture
'--- based on a submission to PSC by Saifudheen A.A.
    Dim lX              As Long
    Dim lY              As Long
    Dim rgbColor        As UcsRgbQuad
    Dim lRedBri         As Long
    Dim lGreenBri       As Long
    Dim lBlueBri        As Long
    Dim lRedSat         As Long
    Dim lGreenSat       As Long
    Dim lBlueSat        As Long
    Dim lIdx            As Long
    Dim clrColor        As OLE_COLOR
    Dim lArea           As Long
    
    On Error GoTo EH
    ReDim rgbLine(0 To RECT_WIDTH) As UcsRgbTriple
    '--- include padding
    ReDim aBits(0 To pvPadScanline(RECT_WIDTH * 3) * RECT_HEIGHT) As Byte
    m_dblTimer = Timer: Debug.Print "pvCreateRectHue "; m_dblTimer;
    clrColor = pvHSBToRGB(hsbColor)
    Call OleTranslateColor(clrColor, 0, rgbColor)
    lArea = (RECT_HEIGHT - 1) * (RECT_WIDTH - 1)
    For lY = 0 To RECT_HEIGHT - 1
        lRedBri = rgbColor.R * lY * (RECT_WIDTH - 1)
        lGreenBri = rgbColor.G * lY * (RECT_WIDTH - 1)
        lBlueBri = rgbColor.b * lY * (RECT_WIDTH - 1)
        lRedSat = (255 - rgbColor.R) * lY
        lGreenSat = (255 - rgbColor.G) * lY
        lBlueSat = (255 - rgbColor.b) * lY
        For lX = 0 To (RECT_WIDTH - 1)
            With rgbLine(RECT_WIDTH - 1 - lX)
                If m_bWebSafeOnly Then
                    .b = m_aWebSafe((lBlueBri + lX * lBlueSat) \ lArea)
                    .G = m_aWebSafe((lGreenBri + lX * lGreenSat) \ lArea)
                    .R = m_aWebSafe((lRedBri + lX * lRedSat) \ lArea)
                Else
                    .b = (lBlueBri + lX * lBlueSat) \ lArea
                    .G = (lGreenBri + lX * lGreenSat) \ lArea
                    .R = (lRedBri + lX * lRedSat) \ lArea
                End If
            End With
        Next
        CopyMemory aBits(lIdx), rgbLine(0), 3 * RECT_WIDTH
        '--- perform padding of DIB scanline
        lIdx = pvPadScanline(lIdx + 3 * RECT_WIDTH)
    Next
    '--- success
    Set pvCreateRectHue = pvExtractRect(aBits)
    Debug.Print Format(Timer - m_dblTimer, "#,##0.0000")
    Exit Function
EH:
    Resume Next
End Function

Private Function pvCreateBarHue() As StdPicture
    Dim lX              As Long
    Dim lY              As Long
    Dim lIdx            As Long
    Dim rgbColor        As UcsRgbQuad
    Dim hsbColor        As UcsHsbColor
    
    ReDim rgbLine(0 To BAR_WIDTH) As UcsRgbTriple
    '--- include padding
    ReDim aBits(0 To pvPadScanline(BAR_WIDTH * 3) * BAR_HEIGHT) As Byte
    hsbColor.Sat = 100
    hsbColor.Bri = 100
    For lY = 0 To BAR_HEIGHT - 1
        '--- floating point division
        hsbColor.Hue = lY * 359 / (BAR_HEIGHT - 1)
        OleTranslateColor pvHSBToRGB(hsbColor), 0, rgbColor
        With rgbLine(0)
            If m_bWebSafeOnly Then
                .R = m_aWebSafe(rgbColor.R)
                .G = m_aWebSafe(rgbColor.G)
                .b = m_aWebSafe(rgbColor.b)
            Else
                .R = rgbColor.R
                .G = rgbColor.G
                .b = rgbColor.b
            End If
        End With
        For lX = 1 To BAR_WIDTH - 1
            rgbLine(lX) = rgbLine(0)
        Next
        CopyMemory aBits(lIdx), rgbLine(0), 3 * BAR_WIDTH
        '--- perform padding of DIB scanline
        lIdx = pvPadScanline(lIdx + 3 * BAR_WIDTH)
    Next
    '--- success
    Set pvCreateBarHue = pvExtractBar(aBits)
End Function

Private Function pvCreateRectSaturation(ByVal dblSat As Double) As StdPicture
    Dim lIdx            As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim nB              As Long
    Dim nS              As Long
    Dim nF              As Long
    Dim bytR            As Byte
    Dim bytG            As Byte
    Dim bytB            As Byte
    
    ReDim rgbLine(0 To RECT_WIDTH) As UcsRgbTriple
    '--- include padding
    ReDim aBits(0 To pvPadScanline(RECT_WIDTH * 3) * RECT_HEIGHT) As Byte
    m_dblTimer = Timer: Debug.Print "pvCreateRectSaturation "; m_dblTimer;
    nS = dblSat * PREC_SAT_255 \ 100
    For lY = 0 To RECT_HEIGHT - 1
        nB = lY * PREC_BRI_255 \ (RECT_HEIGHT - 1)
        For lX = 0 To RECT_WIDTH_STEP - 1
            nF = (lX * PREC_HUE \ RECT_WIDTH_STEP) - (lX \ RECT_WIDTH_STEP) * PREC_HUE
            If m_bWebSafeOnly Then
                bytR = m_aWebSafe(nB \ PREC_BRI)
                bytG = m_aWebSafe(nB * (PREC_HUE_SAT_255 - nS * (PREC_HUE - nF)) \ PREC_HUE_SAT_BRI_255)
                bytB = m_aWebSafe(nB * (PREC_SAT_255 - nS) \ PREC_SAT_BRI_255)
            Else
                bytR = nB \ PREC_BRI
                bytG = nB * (PREC_HUE_SAT_255 - nS * (PREC_HUE - nF)) \ PREC_HUE_SAT_BRI_255
                bytB = nB * (PREC_SAT_255 - nS) \ PREC_SAT_BRI_255
            End If
            With rgbLine(lX)
                .R = bytR
                .G = bytG
                .b = bytB
            End With
            With rgbLine(2 * RECT_WIDTH_STEP - lX - 1)
                .R = bytG
                .G = bytR
                .b = bytB
            End With
            With rgbLine(2 * RECT_WIDTH_STEP + lX)
                .R = bytB
                .G = bytR
                .b = bytG
            End With
            With rgbLine(4 * RECT_WIDTH_STEP - lX - 1)
                .R = bytB
                .G = bytG
                .b = bytR
            End With
            With rgbLine(4 * RECT_WIDTH_STEP + lX)
                .R = bytG
                .G = bytB
                .b = bytR
            End With
            With rgbLine(6 * RECT_WIDTH_STEP - lX - 1)
                .R = bytR
                .G = bytB
                .b = bytG
            End With
        Next
        CopyMemory aBits(lIdx), rgbLine(0), 3 * RECT_WIDTH
        '--- perform padding on scanline
        lIdx = pvPadScanline(lIdx + 3 * RECT_WIDTH)
    Next
    '--- success
    Set pvCreateRectSaturation = pvExtractRect(aBits)
    Debug.Print Format(Timer - m_dblTimer, "#,##0.0000")
End Function

Private Function pvCreateBarSaturation( _
            ByVal dblHue As Double, _
            ByVal dblBri As Double) As StdPicture
    Dim lX              As Long
    Dim lY              As Long
    Dim lIdx            As Long
    Dim hsbColor        As UcsHsbColor
    Dim rgbColor        As UcsRgbQuad
    Dim rgbLine(0 To BAR_WIDTH) As UcsRgbTriple
    
    '--- include padding
    ReDim aBits(0 To pvPadScanline(BAR_WIDTH * 3) * BAR_HEIGHT) As Byte
    hsbColor.Hue = dblHue
    hsbColor.Bri = dblBri
    For lY = 0 To BAR_HEIGHT - 1
        '--- floating point division
        hsbColor.Sat = lY * 100 / (BAR_HEIGHT - 1)
        Call OleTranslateColor(pvHSBToRGB(hsbColor), 0, rgbColor)
        With rgbLine(0)
            If m_bWebSafeOnly Then
                .R = m_aWebSafe(rgbColor.R)
                .G = m_aWebSafe(rgbColor.G)
                .b = m_aWebSafe(rgbColor.b)
            Else
                .R = rgbColor.R
                .G = rgbColor.G
                .b = rgbColor.b
            End If
        End With
        For lX = 1 To BAR_WIDTH - 1
            rgbLine(lX) = rgbLine(0)
        Next
        CopyMemory aBits(lIdx), rgbLine(0), BAR_WIDTH * 3
        '--- perform padding on scanline
        lIdx = pvPadScanline(lIdx + BAR_WIDTH * 3)
    Next
    '--- success
    Set pvCreateBarSaturation = pvExtractBar(aBits)
End Function

Private Function pvCreateRectBrightness(ByVal dblBri As Double) As StdPicture
    Dim lIdx            As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim nB              As Long
    Dim nS              As Long
    Dim nF              As Long
    Dim bytR            As Byte
    Dim bytG            As Byte
    Dim bytB            As Byte
    
    ReDim rgbLine(0 To RECT_WIDTH) As UcsRgbTriple
    '--- include padding
    ReDim aBits(0 To pvPadScanline(RECT_WIDTH * 3) * RECT_HEIGHT) As Byte
    m_dblTimer = Timer: Debug.Print "pvCreateRectBrightness "; m_dblTimer;
    nB = dblBri * PREC_BRI_255 \ 100
    For lY = 0 To RECT_HEIGHT - 1
        nS = lY * PREC_SAT_255 \ (RECT_HEIGHT - 1)
        For lX = 0 To RECT_WIDTH_STEP - 1
            nF = (lX * PREC_HUE \ RECT_WIDTH_STEP) - (lX \ RECT_WIDTH_STEP) * PREC_HUE
            If m_bWebSafeOnly Then
                bytR = m_aWebSafe(nB \ PREC_BRI)
                bytG = m_aWebSafe(nB * (PREC_HUE_SAT_255 - nS * (PREC_HUE - nF)) \ PREC_HUE_SAT_BRI_255)
                bytB = m_aWebSafe(nB * (PREC_SAT_255 - nS) \ PREC_SAT_BRI_255)
            Else
                bytR = nB \ PREC_BRI
                bytG = nB * (PREC_HUE_SAT_255 - nS * (PREC_HUE - nF)) \ PREC_HUE_SAT_BRI_255
                bytB = nB * (PREC_SAT_255 - nS) \ PREC_SAT_BRI_255
            End If
            With rgbLine(lX)
                .R = bytR
                .G = bytG
                .b = bytB
            End With
            With rgbLine(2 * RECT_WIDTH_STEP - lX - 1)
                .R = bytG
                .G = bytR
                .b = bytB
            End With
            With rgbLine(2 * RECT_WIDTH_STEP + lX)
                .R = bytB
                .G = bytR
                .b = bytG
            End With
            With rgbLine(4 * RECT_WIDTH_STEP - lX - 1)
                .R = bytB
                .G = bytG
                .b = bytR
            End With
            With rgbLine(4 * RECT_WIDTH_STEP + lX)
                .R = bytG
                .G = bytB
                .b = bytR
            End With
            With rgbLine(6 * RECT_WIDTH_STEP - lX - 1)
                .R = bytR
                .G = bytB
                .b = bytG
            End With
        Next
        CopyMemory aBits(lIdx), rgbLine(0), 3 * RECT_WIDTH
        '--- perform padding on scanline
        lIdx = pvPadScanline(lIdx + 3 * RECT_WIDTH)
    Next
    '--- success
    Set pvCreateRectBrightness = pvExtractRect(aBits)
    Debug.Print Format(Timer - m_dblTimer, "#,##0.0000")
End Function

Private Function pvCreateBarBrightness( _
            ByVal dblHue As Double, _
            ByVal dblSat As Double) As StdPicture
    Dim lX              As Long
    Dim lY              As Long
    Dim lIdx            As Long
    Dim hsbColor        As UcsHsbColor
    Dim rgbColor        As UcsRgbQuad
    Dim rgbLine(0 To BAR_WIDTH) As UcsRgbTriple
    
    '--- include padding
    ReDim aBits(0 To pvPadScanline(BAR_WIDTH * 3) * BAR_HEIGHT) As Byte
    hsbColor.Hue = dblHue
    hsbColor.Sat = dblSat
    For lY = 0 To BAR_HEIGHT - 1
        '--- floating point division
        hsbColor.Bri = lY * 100 / (BAR_HEIGHT - 1)
        Call OleTranslateColor(pvHSBToRGB(hsbColor), 0, rgbColor)
        With rgbLine(0)
            If m_bWebSafeOnly Then
                .R = m_aWebSafe(rgbColor.R)
                .G = m_aWebSafe(rgbColor.G)
                .b = m_aWebSafe(rgbColor.b)
            Else
                .b = rgbColor.b
                .G = rgbColor.G
                .R = rgbColor.R
            End If
        End With
        For lX = 1 To BAR_WIDTH
            rgbLine(lX) = rgbLine(0)
        Next
        CopyMemory aBits(lIdx), rgbLine(0), BAR_WIDTH * 3
        '--- perform padding on scanline
        lIdx = pvPadScanline(lIdx + 3 * BAR_WIDTH)
    Next
    '--- success
    Set pvCreateBarBrightness = pvExtractBar(aBits)
End Function

Private Function pvCreateRectRGB( _
            ByVal lValue As Long, _
            ByVal eType As UcsRgbColorIdx) As StdPicture
    Dim lX              As Long
    Dim lY              As Long
    Dim lIdx            As Long
    
    ReDim rgbLine(0 To RECT_WIDTH) As UcsRgbTriple
    '--- include padding
    ReDim aBits(0 To pvPadScanline(RECT_WIDTH * 3) * RECT_HEIGHT) As Byte
    m_dblTimer = Timer: Debug.Print "pvCreateRectRGB "; m_dblTimer;
    For lY = 0 To RECT_HEIGHT - 1
        If eType = ucsRgbRed Then
            For lX = 0 To RECT_WIDTH - 1
                With rgbLine(lX)
                    If m_bWebSafeOnly Then
                        .R = m_aWebSafe(lValue)
                        .G = m_aWebSafe(lX * 255 \ (RECT_WIDTH - 1))
                        .b = m_aWebSafe(lY * 255 \ (RECT_HEIGHT - 1))
                    Else
                        .R = lValue
                        .G = lX * 255 \ (RECT_WIDTH - 1)
                        .b = lY * 255 \ (RECT_HEIGHT - 1)
                    End If
                End With
            Next
        ElseIf eType = ucsRgbGreen Then
            For lX = 0 To RECT_WIDTH - 1
                With rgbLine(lX)
                    If m_bWebSafeOnly Then
                        .G = m_aWebSafe(lValue)
                        .R = m_aWebSafe(lX * 255 \ (RECT_WIDTH - 1))
                        .b = m_aWebSafe(lY * 255 \ (RECT_HEIGHT - 1))
                    Else
                        .G = lValue
                        .R = lX * 255 \ (RECT_WIDTH - 1)
                        .b = lY * 255 \ (RECT_HEIGHT - 1)
                    End If
                End With
            Next
        Else '--- eType = ucsRgbBlue
            For lX = 0 To RECT_WIDTH - 1
                With rgbLine(lX)
                    If m_bWebSafeOnly Then
                        .b = m_aWebSafe(lValue)
                        .R = m_aWebSafe(lX * 255 \ (RECT_WIDTH - 1))
                        .G = m_aWebSafe(lY * 255 \ (RECT_HEIGHT - 1))
                    Else
                        .b = lValue
                        .R = lX * 255 \ (RECT_WIDTH - 1)
                        .G = lY * 255 \ (RECT_HEIGHT - 1)
                    End If
                End With
            Next
        End If
        CopyMemory aBits(lIdx), rgbLine(0), 3 * RECT_WIDTH
        '--- perform padding on scanline
        lIdx = pvPadScanline(lIdx + 3 * RECT_WIDTH)
    Next
    '--- success
    Set pvCreateRectRGB = pvExtractRect(aBits)
    Debug.Print Format(Timer - m_dblTimer, "#,##0.0000")
End Function

Private Function pvCreateBarRGB(ByVal eType As UcsRgbColorIdx) As StdPicture
    Dim lX              As Long
    Dim lY              As Long
    Dim lIdx            As Long
    Dim rgbLine(0 To BAR_WIDTH) As UcsRgbTriple
    
    '--- include padding
    ReDim aBits(0 To pvPadScanline(BAR_WIDTH * 3) * BAR_HEIGHT) As Byte
    For lY = 0 To BAR_HEIGHT - 1
        If eType = ucsRgbRed Then
            If m_bWebSafeOnly Then
                rgbLine(0).R = m_aWebSafe(lY * 255 \ (BAR_HEIGHT - 1))
            Else
                rgbLine(0).R = lY * 255 \ (BAR_HEIGHT - 1)
            End If
        ElseIf eType = ucsRgbGreen Then
            If m_bWebSafeOnly Then
                rgbLine(0).G = m_aWebSafe(lY * 255 \ (BAR_HEIGHT - 1))
            Else
                rgbLine(0).G = lY * 255 \ (BAR_HEIGHT - 1)
            End If
        Else '--- eType = ucsRgbBlue
            If m_bWebSafeOnly Then
                rgbLine(0).b = m_aWebSafe(lY * 255 \ (BAR_HEIGHT - 1))
            Else
                rgbLine(0).b = lY * 255 \ (BAR_HEIGHT - 1)
            End If
        End If
        For lX = 1 To BAR_WIDTH - 1
            rgbLine(lX) = rgbLine(0)
        Next
        CopyMemory aBits(lIdx), rgbLine(0), BAR_WIDTH * 3
        '--- perform padding on scanline
        lIdx = pvPadScanline(lIdx + BAR_WIDTH * 3)
    Next
    '--- success
    Set pvCreateBarRGB = pvExtractBar(aBits)
End Function

#If NO_ACCELERATED_GRADIENTS = 0 Then
Private Function pvCreateRectHueAccel(hsbColor As UcsHsbColor) As StdPicture
    Dim lY              As Long
    Dim hsbC1           As UcsHsbColor
    Dim hsbC2           As UcsHsbColor
    Dim pVert(0 To 1)   As TRIVERTEX
    Dim pMesh           As GRADIENT_RECT
    Dim rgbColor        As UcsRgbQuad
        
    m_dblTimer = Timer: Debug.Print "pvCreateRectHueAccel "; m_dblTimer;
    pVert(1).X = RECT_WIDTH
    hsbC1 = pvInitHsb(hsbColor.Hue, 0, 0)
    hsbC2 = pvInitHsb(hsbColor.Hue, 100, 0)
    pMesh.UpperLeft = 0
    pMesh.LowerRight = 1
    With New cMemDC
        .Init RECT_WIDTH, RECT_HEIGHT
        For lY = 0 To RECT_HEIGHT - 1
            '--- floating point division
            hsbC1.Bri = 100 - lY * 100 / (RECT_HEIGHT - 1)
            OleTranslateColor pvHSBToRGB(hsbC1), 0, rgbColor
            With pVert(0)
                .Y = lY
                .Red = pvDWord(256& * rgbColor.R)
                .Green = pvDWord(256& * rgbColor.G)
                .Blue = pvDWord(256& * rgbColor.b)
            End With
            hsbC2.Bri = hsbC1.Bri
            OleTranslateColor pvHSBToRGB(hsbC2), 0, rgbColor
            With pVert(1)
                .Y = lY + 1
                .Red = pvDWord(256& * rgbColor.R)
                .Green = pvDWord(256& * rgbColor.G)
                .Blue = pvDWord(256& * rgbColor.b)
            End With
            GradientFillRect .hDC, pVert(0), 2, pMesh, 1, GRADIENT_FILL_RECT_H
        Next
        '--- success
        Set pvCreateRectHueAccel = .Image
    End With
    Debug.Print Format(Timer - m_dblTimer, "#,##0.0000")
End Function

Private Function pvCreateRectSaturationAccel(ByVal dblSat As Double) As StdPicture
    Dim lX              As Long
    Dim hsbC1           As UcsHsbColor
    Dim hsbC2           As UcsHsbColor
    Dim pVert(0 To 1)   As TRIVERTEX
    Dim pMesh           As GRADIENT_RECT
    Dim rgbColor        As UcsRgbQuad
        
    m_dblTimer = Timer: Debug.Print "pvCreateRectSaturationAccel "; m_dblTimer;
    pVert(1).Y = RECT_HEIGHT
    hsbC1 = pvInitHsb(0, dblSat, 100)
    hsbC2 = pvInitHsb(0, dblSat, 0)
    pMesh.UpperLeft = 0
    pMesh.LowerRight = 1
    With New cMemDC
        .Init RECT_WIDTH, RECT_HEIGHT
        For lX = 0 To RECT_WIDTH - 1
            '--- floating point division
            hsbC1.Hue = lX * 359 / (RECT_WIDTH - 1)
            OleTranslateColor pvHSBToRGB(hsbC1), 0, rgbColor
            With pVert(0)
                .X = lX
                .Red = pvDWord(256& * rgbColor.R)
                .Green = pvDWord(256& * rgbColor.G)
                .Blue = pvDWord(256& * rgbColor.b)
            End With
            hsbC2.Hue = hsbC1.Hue
            OleTranslateColor pvHSBToRGB(hsbC2), 0, rgbColor
            With pVert(1)
                .X = lX + 1
                .Red = pvDWord(256& * rgbColor.R)
                .Green = pvDWord(256& * rgbColor.G)
                .Blue = pvDWord(256& * rgbColor.b)
            End With
            GradientFillRect .hDC, pVert(0), 2, pMesh, 1, GRADIENT_FILL_RECT_V
        Next
        '--- success
        Set pvCreateRectSaturationAccel = .Image
    End With
    Debug.Print Format(Timer - m_dblTimer, "#,##0.0000")
End Function

Private Function pvCreateRectBrightnessAccel(ByVal dblBri As Double) As StdPicture
    Dim lX              As Long
    Dim hsbC1           As UcsHsbColor
    Dim hsbC2           As UcsHsbColor
    Dim pVert(0 To 1)   As TRIVERTEX
    Dim pMesh           As GRADIENT_RECT
    Dim rgbColor        As UcsRgbQuad
        
    m_dblTimer = Timer: Debug.Print "pvCreateRectBrightnessAccel "; m_dblTimer;
    pVert(1).Y = RECT_HEIGHT
    hsbC1 = pvInitHsb(0, 100, dblBri)
    hsbC2 = pvInitHsb(0, 0, dblBri)
    pMesh.UpperLeft = 0
    pMesh.LowerRight = 1
    With New cMemDC
        .Init RECT_WIDTH, RECT_HEIGHT
        For lX = 0 To RECT_WIDTH - 1
            '--- floating point division
            hsbC1.Hue = lX * 359 / (RECT_WIDTH - 1)
            OleTranslateColor pvHSBToRGB(hsbC1), 0, rgbColor
            With pVert(0)
                .X = lX
                .Red = pvDWord(256& * rgbColor.R)
                .Green = pvDWord(256& * rgbColor.G)
                .Blue = pvDWord(256& * rgbColor.b)
            End With
            hsbC2.Hue = hsbC1.Hue
            OleTranslateColor pvHSBToRGB(hsbC2), 0, rgbColor
            With pVert(1)
                .X = lX + 1
                .Red = pvDWord(256& * rgbColor.R)
                .Green = pvDWord(256& * rgbColor.G)
                .Blue = pvDWord(256& * rgbColor.b)
            End With
            GradientFillRect .hDC, pVert(0), 2, pMesh, 1, GRADIENT_FILL_RECT_V
        Next
        '--- success
        Set pvCreateRectBrightnessAccel = .Image
    End With
    Debug.Print Format(Timer - m_dblTimer, "#,##0.0000")
End Function

Private Function pvCreateRectRGBAccel(ByVal lValue As Long, ByVal eType As UcsRgbColorIdx) As StdPicture
    Dim lY              As Long
    Dim pVert(0 To 1)   As TRIVERTEX
    Dim pMesh           As GRADIENT_RECT
        
    m_dblTimer = Timer: Debug.Print "pvCreateRectRGBAccel "; m_dblTimer;
    With pVert(0)
        .Red = pvDWord(lValue * &HFF)
        .Green = pvDWord(lValue * &HFF)
        .Blue = pvDWord(lValue * &HFF)
    End With
    With pVert(1)
        .X = RECT_WIDTH
        .Red = pvDWord(lValue * &HFF)
        .Green = pvDWord(lValue * &HFF)
        .Blue = pvDWord(lValue * &HFF)
    End With
    pMesh.UpperLeft = 0
    pMesh.LowerRight = 1
    With New cMemDC
        .Init RECT_WIDTH, RECT_HEIGHT
        For lY = 0 To RECT_HEIGHT - 1
            If eType = ucsRgbRed Then
                pVert(0).Green = 0
                pVert(0).Blue = pvDWord((RECT_HEIGHT - 1 - lY) * 255 * 255 \ (RECT_HEIGHT - 1))
                pVert(1).Green = &HFF00
                pVert(1).Blue = pVert(0).Blue
            ElseIf eType = ucsRgbGreen Then
                pVert(0).Red = 0
                pVert(0).Blue = pvDWord((RECT_HEIGHT - 1 - lY) * 255 * 255 \ (RECT_HEIGHT - 1))
                pVert(1).Red = &HFF00
                pVert(1).Blue = pVert(0).Blue
            Else ' --- eType = ucsRgbBlue
                pVert(0).Green = pvDWord((RECT_HEIGHT - 1 - lY) * 255 * 255 \ (RECT_HEIGHT - 1))
                pVert(0).Red = 0
                pVert(1).Green = pVert(0).Green
                pVert(1).Red = &HFF00
            End If
            pVert(0).Y = lY
            pVert(1).Y = lY + 1
            GradientFillRect .hDC, pVert(0), 2, pMesh, 1, GRADIENT_FILL_RECT_H
        Next
        '--- success
        Set pvCreateRectRGBAccel = .Image
    End With
    Debug.Print Format(Timer - m_dblTimer, "#,##0.0000")
End Function
#End If

Private Sub pvPaintForm()
    Dim lLeft           As Long
    Dim lTop            As Long
    Dim lRight          As Long
    Dim lBottom         As Long
    Dim rgbColor        As UcsRgbQuad
    Dim clrCircle       As Long
    
    With New cMemDC
        .Init ScaleWidth, ScaleHeight, , hDC
        '--- cleanup (especially markers)
        .Cls BackColor
        '--- paint color rectangle
        lLeft = imgRect.Left - 2: lTop = imgRect.Top - 2
        lRight = lLeft + RECT_WIDTH + 4
        lBottom = lTop + RECT_HEIGHT + 4
        .DrawEdge lLeft, lTop, lRight, lBottom
        .FrameRect lLeft + 1, lTop + 1, lRight - 1, lBottom - 1, vbBlack
        .PaintPicture m_imgRect, lLeft + 2, lTop + 2
        '--- paint color bar
        lLeft = imgBar.Left + GRID_SIZE - 2: lTop = imgBar.Top - 2
        lRight = lLeft + BAR_WIDTH + 4
        lBottom = lTop + BAR_HEIGHT + 4
        .DrawEdge lLeft, lTop, lRight, lBottom
        .FrameRect lLeft + 1, lTop + 1, lRight - 1, lBottom - 1, vbBlack
        .PaintPicture m_imgBar, lLeft + 2, lTop + 2
        '--- calc markers positions (left,top) -> rect, (right,bottom) -> bar
        OleTranslateColor m_clrCurrent, 0, rgbColor
        lRight = imgBar.Left + GRID_SIZE - 7
        If optHue Then
            lLeft = imgRect.Left + m_hsbCurrent.Sat * (RECT_WIDTH - 1) \ 100
            lTop = imgRect.Top + (100 - m_hsbCurrent.Bri) * (RECT_HEIGHT - 1) \ 100
            lBottom = imgBar.Top + (359 - m_hsbCurrent.Hue) * (BAR_HEIGHT - 1) \ 359 - 3
        ElseIf optSat Then
            lLeft = imgRect.Left + m_hsbCurrent.Hue * (RECT_WIDTH - 1) \ 359
            lTop = imgRect.Top + (100 - m_hsbCurrent.Bri) * (RECT_HEIGHT - 1) \ 100
            lBottom = imgBar.Top + (100 - m_hsbCurrent.Sat) * (BAR_HEIGHT - 1) \ 100 - 3
        ElseIf optBri Then
            lLeft = imgRect.Left + m_hsbCurrent.Hue * (RECT_WIDTH - 1) \ 359
            lTop = imgRect.Top + (100 - m_hsbCurrent.Sat) * (RECT_HEIGHT - 1) \ 100
            lBottom = imgBar.Top + (100 - m_hsbCurrent.Bri) * (BAR_HEIGHT - 1) \ 100 - 3
        ElseIf optRed Then
            lLeft = imgRect.Left + rgbColor.G * (RECT_WIDTH - 1) \ 255
            lTop = imgRect.Top + (255 - rgbColor.b) * (RECT_HEIGHT - 1) \ 255
            lBottom = imgBar.Top + (255 - rgbColor.R) * (BAR_HEIGHT - 1) \ 255 - 3
        ElseIf optGreen Then
            lLeft = imgRect.Left + rgbColor.R * (RECT_WIDTH - 1) \ 255
            lTop = imgRect.Top + (255 - rgbColor.b) * (RECT_HEIGHT - 1) \ 255
            lBottom = imgBar.Top + (255 - rgbColor.G) * (BAR_HEIGHT - 1) \ 255 - 3
        ElseIf optBlue Then
            lLeft = imgRect.Left + rgbColor.R * (RECT_WIDTH - 1) \ 255
            lTop = imgRect.Top + (255 - rgbColor.G) * (RECT_HEIGHT - 1) \ 255
            lBottom = imgBar.Top + (255 - rgbColor.b) * (BAR_HEIGHT - 1) \ 255 - 3
        End If
        '--- paint rectangle marker
        OleTranslateColor m_clrCurrent, 0, rgbColor
        '--- try to figure intensity (formula based on glimpses of memory;-))
        If rgbColor.R * 0.299 + rgbColor.G * 0.587 + rgbColor.b * 0.114 > 127 Then
            clrCircle = vbBlack
        Else
            clrCircle = vbWhite
        End If
        .DrawEllipse lLeft - 3, lTop - 3, lLeft + 3, lTop + 3, vbWhite - clrCircle
        .DrawEllipse lLeft - 4, lTop - 4, lLeft + 4, lTop + 4, clrCircle
        .DrawEllipse lLeft - 5, lTop - 5, lLeft + 5, lTop + 5, vbWhite - clrCircle
        '--- paint bar marker
        .PaintPicture m_imgBarSelector, lRight, lBottom
        .Destroy
    End With
    '--- flush memory dc bitmap
    If AutoRedraw Then
        Refresh
    End If
End Sub

Private Function pvHSBToRGB(hsbColor As UcsHsbColor) As Long
'--- based on *cool* code by Branco Medeiros (http://www.myrealbox.com/branco_medeiros)
'--- Converts an HSB value to the RGB color model. Adapted from Java.awt.Color.java
    Dim nH              As Double
    Dim nS              As Double
    Dim nL              As Double
    Dim nF              As Double
    Dim nP              As Double
    Dim nQ              As Double
    Dim nT              As Double
    Dim lH              As Long
    Dim clrConv         As UcsRgbQuad

    With clrConv
        If hsbColor.Sat > 0 Then
            nH = hsbColor.Hue / 60
            nL = hsbColor.Bri / 100
            nS = hsbColor.Sat / 100
            lH = Int(nH)
            nF = nH - lH
            nP = nL * (1 - nS)
            nQ = nL * (1 - nS * nF)
            nT = nL * (1 - nS * (1 - nF))
            Select Case lH
            Case 0
                .R = nL * 255
                .G = nT * 255
                .b = nP * 255
            Case 1
                .R = nQ * 255
                .G = nL * 255
                .b = nP * 255
            Case 2
                .R = nP * 255
                .G = nL * 255
                .b = nT * 255
            Case 3
                .R = nP * 255
                .G = nQ * 255
                .b = nL * 255
            Case 4
                .R = nT * 255
                .G = nP * 255
                .b = nL * 255
            Case 5
                .R = nL * 255
                .G = nP * 255
                .b = nQ * 255
            End Select
        Else
            .R = (hsbColor.Bri * 255) / 100
            .G = .R
            .b = .R
        End If
    End With
    '--- return long
    CopyMemory lH, clrConv, 4
    pvHSBToRGB = lH
End Function

Private Function pvRGBToHSB(ByVal clrValue As OLE_COLOR) As UcsHsbColor
'--- based on *cool* code by Branco Medeiros (http://www.myrealbox.com/branco_medeiros)
'--- Converts an RGB value to the HSB color model. Adapted from Java.awt.Color.java
    Dim nTemp           As Double
    Dim lMin            As Long
    Dim lMax            As Long
    Dim lDelta          As Long
    Dim rgbValue        As UcsRgbQuad
  
    Call OleTranslateColor(clrValue, 0, rgbValue)
    lMax = IIf(rgbValue.R > rgbValue.G, IIf(rgbValue.R > rgbValue.b, rgbValue.R, rgbValue.b), IIf(rgbValue.G > rgbValue.b, rgbValue.G, rgbValue.b))
    lMin = IIf(rgbValue.R < rgbValue.G, IIf(rgbValue.R < rgbValue.b, rgbValue.R, rgbValue.b), IIf(rgbValue.G < rgbValue.b, rgbValue.G, rgbValue.b))
    lDelta = lMax - lMin
    pvRGBToHSB.Bri = (lMax * 100) / 255
    If lMax > 0 Then
        pvRGBToHSB.Sat = (lDelta / lMax) * 100
        If lDelta > 0 Then
            If lMax = rgbValue.R Then
                nTemp = (CLng(rgbValue.G) - rgbValue.b) / lDelta
            ElseIf lMax = rgbValue.G Then
                nTemp = 2 + (CLng(rgbValue.b) - rgbValue.R) / lDelta
            Else
                nTemp = 4 + (CLng(rgbValue.R) - rgbValue.G) / lDelta
            End If
            pvRGBToHSB.Hue = nTemp * 60
            If pvRGBToHSB.Hue < 0 Then
                pvRGBToHSB.Hue = pvRGBToHSB.Hue + 360
            End If
        End If
    End If
End Function

Private Function pvRGBToCMYK(ByVal clrValue As OLE_COLOR) As UcsRgbQuad
'--- retval: CMYK encoded in RGBA
    Dim lK              As Long
    Dim rgbColor        As UcsRgbQuad
    
    OleTranslateColor clrValue, 0, rgbColor
    With rgbColor
        lK = pvMin(pvMin((255 - .R) * 100 \ 255, _
                         (255 - .G) * 100 \ 255), _
                    (255 - .b) * 100 \ 255)
        pvRGBToCMYK.R = (255 - .R) * 100 \ 255 - lK
        pvRGBToCMYK.G = (255 - .G) * 100 \ 255 - lK
        pvRGBToCMYK.b = (255 - .b) * 100 \ 255 - lK
        pvRGBToCMYK.a = lK
    End With
End Function

Private Function pvRGBToXYZ(ByVal clrValue As OLE_COLOR) As UcsXyzColor
'--- multiplication matrix values are from ITU reference
    Dim rgbColor        As UcsRgbQuad
    Dim xyzColor        As UcsXyzColor
    
    OleTranslateColor clrValue, 0, rgbColor
    With xyzColor
        .X = pvRGBToXYZHelper(rgbColor.R / 255#)
        .Y = pvRGBToXYZHelper(rgbColor.G / 255#)
        .Z = pvRGBToXYZHelper(rgbColor.b / 255#)
        pvRGBToXYZ.X = (0.412453 * .X + 0.35758 * .Y + 0.180423 * .Z)
        pvRGBToXYZ.Y = (0.212671 * .X + 0.71516 * .Y + 0.072169 * .Z)
        pvRGBToXYZ.Z = (0.019334 * .X + 0.119193 * .Y + 0.950227 * .Z)
    End With
End Function

Private Function pvRGBToXYZHelper(dblT As Double) As Double
'    If dblT > 0.03928 Then
'        pvRGBToXYZHelper = ((dblT + 0.055) / 1.055) ^ 2.4
'    Else
'        pvRGBToXYZHelper = dblT / 12.92
'    End If
    pvRGBToXYZHelper = dblT ^ (1 / 0.45)
End Function

Private Function pvXYZToLAB(xyzValue As UcsXyzColor) As UcsLabColor
    Dim xyzColor        As UcsXyzColor
    
    With xyzColor
        .X = pvRGBToLABHelper(xyzValue.X / 0.950456)
        .Y = pvRGBToLABHelper(xyzValue.Y / 1#)
        .Z = pvRGBToLABHelper(xyzValue.Z / 1.088754)
        If xyzValue.Y < 0.008856 Then
            pvXYZToLAB.L = 903.3 * xyzValue.Y
        Else
            pvXYZToLAB.L = 116 * .Y - 16
        End If
        pvXYZToLAB.a = 500 * (.X - .Y) / LAB_CORELDRAW_NORMALIZE
        pvXYZToLAB.b = 200 * (.Y - .Z) / LAB_CORELDRAW_NORMALIZE
    End With
End Function

Private Function pvRGBToLABHelper(dblT As Double) As Double
'    If dblT > 0.008856 Then
'        pvRGBToLABHelper = dblT ^ (1# / 3)
'    Else
'        pvRGBToLABHelper = 7.787 * dblT + 16 / 116
'    End If
    pvRGBToLABHelper = dblT ^ (1# / 3)
End Function

Private Sub pvRefresh()
    Dim clrCurrent      As Long
    
    clrCurrent = m_clrCurrent
    m_clrCurrent = -1
    m_hsbCurrent.Hue = -1
    Color = clrCurrent
    pvInvalidate
End Sub
   
Private Sub pvInvalidate()
    Dim rc As RECT

    AutoRedraw = False
    GetClientRect hwnd, rc
    InvalidateRect hwnd, rc, 1
End Sub

Private Function pvExtractRect(aBits() As Byte) As StdPicture
'--- extract "Rect" StdPicture from DIBs
    With New cMemDC
        .Init RECT_WIDTH, RECT_HEIGHT
        '--- take care of 256 color displays
        Call SetStretchBltMode(.hDC, HALFTONE)
        .SetDIBits 0, 0, RECT_WIDTH, RECT_HEIGHT, aBits
        Set pvExtractRect = .Image
    End With
End Function

Private Function pvExtractBar(aBits() As Byte) As StdPicture
'--- extract "Bar" StdPicture from DIBs
    With New cMemDC
        .Init BAR_WIDTH, BAR_HEIGHT
        '--- take care of 256 color displays
        Call SetStretchBltMode(.hDC, HALFTONE)
        .SetDIBits 0, 0, BAR_WIDTH, BAR_HEIGHT, aBits
        Set pvExtractBar = .Image
    End With
End Function

Private Sub pvSetText(oCtl As TextBox, ByVal sText As String)
'--- set text to TextBox and select all -- much like a regular win32 edit control
    With oCtl
        .Text = sText
        If Not ActiveControl Is oCtl Then
            .SelStart = 0
            .SelLength = Len(sText)
        End If
    End With
End Sub

Private Sub pvResetTimer(sMode As String)
'--- reset timer i.e. start ticking from 0 if in the middle of timeout
    Dim sText           As String
    
    On Error Resume Next
    With tmrKeyboard
        '--- flush timer event if other color space input pending
        If .Tag <> "" And .Tag <> sMode Then
            '--- dont lose current textbox value
            sText = ActiveControl.Text
            tmrKeyboard_Timer
            ActiveControl.Text = sText
        End If
        .Tag = sMode
        .Enabled = False
        .Enabled = True
    End With
End Sub

'= Utility private methods ===============================================

Private Function pvHex(ByVal lValue As Long, Optional lCount As Long = 2) As String
'--- convert hex and pad with zeroes
    pvHex = Right(String(lCount, "0") & Hex(lValue), lCount)
End Function

Private Function pvPadScanline(ByVal lOffset As Long)
'--- DIB section horizontal scanline padding to dword
    pvPadScanline = (lOffset + 3) And (Not 3)
End Function

Private Function pvLimit( _
            ByVal dblValue As Double, _
            ByVal dblMin As Double, _
            ByVal dblMax As Double) As Double
'--- limit double value to upper and lower bound
    If dblValue < dblMin Then
        pvLimit = dblMin
    ElseIf dblValue > dblMax Then
        pvLimit = dblMax
    Else
        pvLimit = dblValue
    End If
End Function

Private Function pvIsEqualHsb( _
            oC1 As UcsHsbColor, _
            oC2 As UcsHsbColor) As Boolean
'--- compare HSB colors for equality (and inequality)
    pvIsEqualHsb = (oC1.Hue = oC2.Hue) _
                And (oC1.Sat = oC2.Sat) _
                And (oC1.Bri = oC2.Bri)
End Function

Private Function pvInitHsb( _
            ByVal dblHue As Double, _
            ByVal dblSat As Double, _
            ByVal dblBri As Double) As UcsHsbColor
'--- "class factory" for HSB colors
    With pvInitHsb
        .Hue = dblHue
        .Sat = dblSat
        .Bri = dblBri
    End With
End Function

Private Function pvMax(ByVal lA As Long, ByVal lB As Long) As Long
'--- retval: maximum of both arguments
    pvMax = IIf(lA > lB, lA, lB)
End Function

Private Function pvMin(ByVal lA As Long, ByVal lB As Long) As Long
'--- retval: minimum of both arguments
    pvMin = IIf(lA < lB, lA, lB)
End Function

Private Function pvDWord(ByVal lValue As Long) As Integer
'--- long to unsigned dword conversion
    If lValue >= &H8000& Then
        pvDWord = lValue - &H10000
    Else
        pvDWord = lValue
    End If
End Function

Private Function pvHM2Pix(ByVal Value As Double) As Double
'--- himetric to pixels conversion
   pvHM2Pix = Value * 1440 / 2540 / Screen.TwipsPerPixelX
End Function

Private Function pvCheckDimensions( _
            ByVal imgPic As StdPicture, _
            ByVal lWidth As Long, _
            ByVal lHeight As Long) As Boolean
'--- retval: true - cached image dimensions are current, false - repaint needed
    If Not imgPic Is Nothing Then
        pvCheckDimensions = Abs(lWidth - pvHM2Pix(imgPic.Width)) < 1 _
                    And Abs(lHeight - pvHM2Pix(imgPic.Height)) < 1
    End If
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_Load()
    Dim lIdx            As Long
        
    On Error Resume Next
    '--- reference colors border
    labOld.Move 2, 2, picReference.ScaleWidth - 4, picReference.ScaleHeight \ 2 - 2
    labNew.Move 2, picReference.ScaleHeight \ 2, picReference.ScaleWidth - 4, picReference.ScaleHeight - picReference.ScaleHeight \ 2 - 2
    With New cMemDC
        .Init picReference.ScaleWidth, picReference.ScaleHeight, , picReference.hDC
        .DrawEdge
        .FrameRect 1, 1, .Width - 1, .Height - 1, vbBlack
    End With
    '--- precalculate safe-colors array
    For lIdx = 0 To 255
        m_aWebSafe(lIdx) = CByte((lIdx + 25) \ 51) * 51
    Next
    '--- draw bar selector in mem dc
    With New cMemDC
        .Init BAR_WIDTH + 13, 7
        .Cls MASK_COLOR
        For lIdx = 0 To 3
            .DrawLine lIdx, lIdx, lIdx, 7 - lIdx
            .DrawLine BAR_WIDTH + 12 - lIdx, lIdx, BAR_WIDTH + 12 - lIdx, 7 - lIdx
        Next
        Set m_imgBarSelector = .ExtractIcon(MASK_COLOR)
    End With
    '--- for resize
    m_sNumericHeight = fraNumeric.Height
    Form_Resize
#If NO_ACCELERATED_GRADIENTS = 0 Then
    '--- check is acceleareted gradients supported by os
    Dim pVert As TRIVERTEX
    Dim pMesh As GRADIENT_TRIANGLE
    GradientFill 0, pVert, 0, pMesh, 0, 0
    '--- possible Err.Number, Err.Description:
    '---   453, Can't find DLL entry point GradientFill in Msimg32.dll
    '---   53, File not found: Msimg32.dll
    If Err.Number = 0 Then
        m_bAccelerateSupported = True
        chkAccel.Enabled = True
        chkAccel.Visible = True
    End If
#End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Visible = False
        DoEvents
    End If
End Sub

Private Sub Form_Resize()
    Dim lIdx        As Long
    
    On Error Resume Next
    '--- calc "constants" :-))
    RECT_WIDTH = ScaleWidth - fraNumeric.Width - BAR_WIDTH - 12 * GRID_SIZE
    RECT_WIDTH_STEP = RECT_WIDTH \ 6
    RECT_WIDTH = RECT_WIDTH_STEP * 6
    RECT_HEIGHT = ScaleHeight - 4 * GRID_SIZE - chkWebSafe.Height
    BAR_HEIGHT = RECT_HEIGHT
    '--- move click images
    imgRect.Move 2 * GRID_SIZE + 2, 2 * GRID_SIZE + 2, RECT_WIDTH, RECT_HEIGHT
    imgBar.Move RECT_WIDTH + 4 * GRID_SIZE + 2, 2 * GRID_SIZE + 2, BAR_WIDTH + 2 * GRID_SIZE, BAR_HEIGHT
    '--- move controls around
    chkWebSafe.Move imgRect.Left, imgRect.Top + imgRect.Height + GRID_SIZE
    chkAccel.Move chkWebSafe.Left + chkWebSafe.Width, chkWebSafe.Top
    lIdx = imgBar.Left + imgBar.Width + 4 * GRID_SIZE
    fraColors.Move lIdx, GRID_SIZE
    fraNumeric.Move lIdx, fraColors.Top + fraColors.Height + GRID_SIZE
    lIdx = fraNumeric.Left + fraNumeric.Width - cmdCancel.Width
    cmdOk.Move lIdx, 2 * GRID_SIZE
    cmdCancel.Move lIdx, cmdOk.Top + cmdOk.Height + GRID_SIZE
    If fraNumeric.Top + m_sNumericHeight > ScaleHeight Then
        fraNumeric.Height = picAdditional.Top \ Screen.TwipsPerPixelY
        picAdditional.Visible = False
    Else
        fraNumeric.Height = m_sNumericHeight
        picAdditional.Visible = True
    End If
    pvRefresh
End Sub

Private Sub Form_Paint()
    AutoRedraw = True
    pvPaintForm
End Sub

Private Sub cmdOk_Click()
    m_bOk = True
    Visible = False
    DoEvents
End Sub

Private Sub cmdCancel_Click()
    Visible = False
    DoEvents
End Sub

'= imgBar mouse selection ================================================

Private Sub imgBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bBarPressed = True
    Call imgBar_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--- note: divisions are floating point (i.e. "/" instead of "\")
    Dim rgbColor        As UcsRgbQuad
    
    On Error Resume Next
    If m_bBarPressed Then
        If optHue Then
            m_hsbCurrent.Hue = 359 - pvLimit((Y / Screen.TwipsPerPixelY) * 359 / (BAR_HEIGHT - 1), 0, 359)
            Color = pvHSBToRGB(m_hsbCurrent)
        ElseIf optSat Then
            m_hsbCurrent.Sat = 100 - pvLimit((Y / Screen.TwipsPerPixelY) * 100 / (BAR_HEIGHT - 1), 0, 100)
            Color = pvHSBToRGB(m_hsbCurrent)
        ElseIf optBri Then
            m_hsbCurrent.Bri = 100 - pvLimit((Y / Screen.TwipsPerPixelY) * 100 / (BAR_HEIGHT - 1), 0, 100)
            Color = pvHSBToRGB(m_hsbCurrent)
        ElseIf optRed Then
            OleTranslateColor m_clrCurrent, 0, rgbColor
            m_hsbCurrent.Hue = -1
            Color = RGB(255 - pvLimit((Y / Screen.TwipsPerPixelY) * 255 / (BAR_HEIGHT - 1), 0, 255), rgbColor.G, rgbColor.b)
        ElseIf optGreen Then
            OleTranslateColor m_clrCurrent, 0, rgbColor
            m_hsbCurrent.Hue = -1
            Color = RGB(rgbColor.R, 255 - pvLimit((Y / Screen.TwipsPerPixelY) * 255 / (BAR_HEIGHT - 1), 0, 255), rgbColor.b)
        ElseIf optBlue Then
            OleTranslateColor m_clrCurrent, 0, rgbColor
            m_hsbCurrent.Hue = -1
            Color = RGB(rgbColor.R, rgbColor.G, 255 - pvLimit((Y / Screen.TwipsPerPixelY) * 255 / (BAR_HEIGHT - 1), 0, 255))
        End If
        AutoRedraw = False
        Refresh
    End If
End Sub

Private Sub imgBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgBar_MouseMove(Button, Shift, X, Y)
    m_bBarPressed = False
End Sub

'= imgRect mouse selection ===============================================

Private Sub imgRect_DblClick()
    cmdOk.Value = True
End Sub

Private Sub imgRect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bRectPressed = True
    Call imgRect_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgRect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--- note: divisions are floating point (i.e. "/" instead of "\")
    Dim rgbColor        As UcsRgbQuad
    
    On Error Resume Next
    If m_bRectPressed Then
        If optHue Then
            m_hsbCurrent.Sat = pvLimit((X / Screen.TwipsPerPixelX) * 100 / (RECT_WIDTH - 1), 0, 100)
            m_hsbCurrent.Bri = 100 - pvLimit((Y / Screen.TwipsPerPixelY) * 100 / (RECT_HEIGHT - 1), 0, 100)
            Color = pvHSBToRGB(m_hsbCurrent)
        ElseIf optSat Then
            m_hsbCurrent.Hue = pvLimit((X / Screen.TwipsPerPixelX) * 359 / (RECT_WIDTH - 1), 0, 359)
            m_hsbCurrent.Bri = 100 - pvLimit((Y / Screen.TwipsPerPixelY) * 100 / (RECT_HEIGHT - 1), 0, 100)
            Color = pvHSBToRGB(m_hsbCurrent)
        ElseIf optBri Then
            m_hsbCurrent.Hue = pvLimit((X / Screen.TwipsPerPixelX) * 359 / (RECT_WIDTH - 1), 0, 359)
            m_hsbCurrent.Sat = 100 - pvLimit((Y / Screen.TwipsPerPixelY) * 100 / (RECT_HEIGHT - 1), 0, 100)
            Color = pvHSBToRGB(m_hsbCurrent)
        ElseIf optRed Then
            OleTranslateColor m_clrCurrent, 0, rgbColor
            m_hsbCurrent.Hue = -1
            Color = RGB(rgbColor.R, _
                        pvLimit((X / Screen.TwipsPerPixelX) * 255 / (RECT_WIDTH - 1), 0, 255), _
                        255 - pvLimit((Y / Screen.TwipsPerPixelY) * 255 / (RECT_HEIGHT - 1), 0, 255))
        ElseIf optGreen Then
            OleTranslateColor m_clrCurrent, 0, rgbColor
            m_hsbCurrent.Hue = -1
            Color = RGB(pvLimit((X / Screen.TwipsPerPixelX) * 255 / (RECT_WIDTH - 1), 0, 255), _
                        rgbColor.G, _
                        255 - pvLimit((Y / Screen.TwipsPerPixelY) * 255 / (RECT_HEIGHT - 1), 0, 255))
        ElseIf optBlue Then
            OleTranslateColor m_clrCurrent, 0, rgbColor
            m_hsbCurrent.Hue = -1
            Color = RGB(pvLimit((X / Screen.TwipsPerPixelX) * 255 / (RECT_WIDTH - 1), 0, 255), _
                        255 - pvLimit((Y / Screen.TwipsPerPixelY) * 255 / (RECT_HEIGHT - 1), 0, 255), _
                        rgbColor.b)
        End If
        AutoRedraw = False
        Refresh
    End If
End Sub

Private Sub imgRect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgRect_MouseMove(Button, Shift, X, Y)
    m_bRectPressed = False
End Sub

'= current view switch ===================================================

Private Sub optBri_Click()
    pvRefresh
End Sub

Private Sub optHue_Click()
    pvRefresh
End Sub

Private Sub optSat_Click()
    pvRefresh
End Sub

Private Sub optRed_Click()
    pvRefresh
End Sub

Private Sub optGreen_Click()
    pvRefresh
End Sub

Private Sub optBlue_Click()
    pvRefresh
End Sub

'= user keyboard input ===================================================

Private Sub txtHue_Change()
    If Not m_bInSet Then
        pvResetTimer STR_TIMER_FROM_HSB
    End If
End Sub

Private Sub txtSat_Change()
    If Not m_bInSet Then
        pvResetTimer STR_TIMER_FROM_HSB
    End If
End Sub

Private Sub txtBri_Change()
    If Not m_bInSet Then
        pvResetTimer STR_TIMER_FROM_HSB
    End If
End Sub

Private Sub txtRed_Change()
    If Not m_bInSet Then
        pvResetTimer STR_TIMER_FROM_RGB
    End If
End Sub

Private Sub txtGreen_Change()
    If Not m_bInSet Then
        pvResetTimer STR_TIMER_FROM_RGB
    End If
End Sub

Private Sub txtBlue_Change()
    If Not m_bInSet Then
        pvResetTimer STR_TIMER_FROM_RGB
    End If
End Sub

Private Sub updHue_Change(ByVal lValue As Long)
    txtHue = pvLimit(Val(txtHue) + lValue, 0, 359)
    tmrKeyboard_Timer
End Sub

Private Sub updSat_Change(ByVal lValue As Long)
    txtSat = pvLimit(Val(txtSat) + lValue, 0, 100)
    tmrKeyboard_Timer
End Sub

Private Sub updBri_Change(ByVal lValue As Long)
    txtBri = pvLimit(Val(txtBri) + lValue, 0, 100)
    tmrKeyboard_Timer
End Sub

Private Sub updRed_Change(ByVal lValue As Long)
    txtRed = pvLimit(Val(txtRed) + lValue, 0, 255)
    tmrKeyboard_Timer
End Sub

Private Sub updGreen_Change(ByVal lValue As Long)
    txtGreen = pvLimit(Val(txtGreen) + lValue, 0, 255)
    tmrKeyboard_Timer
End Sub

Private Sub updBlue_Change(ByVal lValue As Long)
    txtBlue = pvLimit(Val(txtBlue) + lValue, 0, 255)
    tmrKeyboard_Timer
End Sub

Private Sub tmrKeyboard_Timer()
    Dim rgbValue        As UcsRgbQuad
    Dim clrValue        As Long
    
    '--- check keyboard input mode
    If tmrKeyboard.Tag = STR_TIMER_FROM_RGB Then
        OleTranslateColor Color, 0, rgbValue
        rgbValue.R = pvLimit(Val(txtRed), 0, 255)
        rgbValue.G = pvLimit(Val(txtGreen), 0, 255)
        rgbValue.b = pvLimit(Val(txtBlue), 0, 255)
        CopyMemory clrValue, rgbValue, 4
        m_hsbCurrent.Hue = -1
        Color = clrValue
        pvRefresh
    ElseIf tmrKeyboard.Tag = STR_TIMER_FROM_HSB Then
        m_hsbCurrent.Hue = pvLimit(Val(txtHue), 0, 359)
        m_hsbCurrent.Sat = pvLimit(Val(txtSat), 0, 100)
        m_hsbCurrent.Bri = pvLimit(Val(txtBri), 0, 100)
        Color = pvHSBToRGB(m_hsbCurrent)
        pvInvalidate
    End If
    '--- stop timer
    tmrKeyboard.Enabled = False
    tmrKeyboard.Tag = ""
End Sub

'= misc ==================================================================

Private Sub labOld_Click()
'--- undocumented feature: restore orig color upon click ;-))
    Color = m_clrOriginal
    pvRefresh
End Sub

Private Sub chkAccel_Click()
    '--- cleanup cache (only "accelerated" bitmaps)
    Set m_oHueCache.imgRect = Nothing
    Set m_oSaturationCache.imgRect = Nothing
    Set m_oBrightnessCache.imgRect = Nothing
    Set m_oRedCache.imgRect = Nothing
    Set m_oGreenCache.imgRect = Nothing
    Set m_oBlueCache.imgRect = Nothing
    m_bAccelerateSupported = chkAccel.Value = vbChecked
    pvRefresh
End Sub

Private Sub chkWebSafe_Click()
    m_bWebSafeOnly = (chkWebSafe.Value = vbChecked)
    chkAccel.Enabled = Not m_bWebSafeOnly
    pvRefresh
End Sub
