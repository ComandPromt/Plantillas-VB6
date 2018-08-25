VERSION 5.00
Begin VB.UserControl ThumbNailList 
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   Picture         =   "ThumbNailList.ctx":0000
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ToolboxBitmap   =   "ThumbNailList.ctx":0342
   Begin VB.PictureBox Gradient 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.PictureBox PreviewPicture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      DataField       =   "Thumb"
      DataSource      =   "DataPictures"
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   120
      Picture         =   "ThumbNailList.ctx":0654
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   3030
      Begin VB.PictureBox P_load 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   90
         ScaleHeight     =   53
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   77
         TabIndex        =   6
         Top             =   30
         Visible         =   0   'False
         Width           =   1185
         Begin VB.FileListBox File1 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Left            =   90
            Pattern         =   "*.jpg;*.gif;*.Bmp"
            TabIndex        =   7
            Top             =   60
            Visible         =   0   'False
            Width           =   825
         End
      End
   End
   Begin VB.VScrollBar V_Scroll 
      Enabled         =   0   'False
      Height          =   3720
      LargeChange     =   66
      Left            =   3420
      Max             =   0
      SmallChange     =   66
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   210
      Width           =   225
   End
   Begin VB.PictureBox ThumbOuterFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   226
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   3420
      Begin VB.PictureBox ThumbInnerframe 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   3585
         Index           =   0
         Left            =   90
         ScaleHeight     =   239
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   217
         TabIndex        =   1
         Top             =   90
         Width           =   3255
         Begin VB.PictureBox ThumbPage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   0
            Left            =   90
            ScaleHeight     =   29
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   33
            TabIndex        =   4
            Top             =   60
            Visible         =   0   'False
            Width           =   525
         End
      End
   End
   Begin VB.PictureBox HtmlPreviewpicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      DataField       =   "Thumb"
      DataSource      =   "DataPictures"
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   870
      Picture         =   "ThumbNailList.ctx":32FF
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.Label LbInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Info:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   225
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3675
   End
End
Attribute VB_Name = "ThumbNailList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function BmpToJpeg Lib "JPeg32.dll" (ByVal BmpFilename As String, ByVal JpegFilename As String, ByVal Quality As Integer) As Integer
Private Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer
Const GFSR_SYSTEMRESOURCES = &H0



Enum ThumbID
     X_1
     y_1
     X_2
     y_2
     File_name
     pagina
End Enum

Dim Thumb_Index()
Dim ThumbSelected

Enum EnumLang
     English = 0
     Dutch = 1
End Enum
Dim m_Language As EnumLang
Const m_def_Language = 1

Enum EnumPrintResolution
     DPI_25 = 25
     DPI_50 = 50
     DPI_75 = 75
     DPI_100 = 100
     DPI_125 = 125
     DPI_150 = 150
     DPI_175 = 175
     DPI_200 = 200
     DPI_225 = 225
     DPI_250 = 250
     DPI_275 = 275
     DPI_300 = 300
End Enum
Dim m_PrintResolution
Const m_def_PrintResolution = 200

Enum EnumPrint_Zoomfactor
     x10 = 10
     x20 = 20
     x30 = 30
     x40 = 40
     x50 = 50
     x60 = 60
     x70 = 70
     x80 = 80
     x90 = 90
     x100 = 100
     x110 = 110
     x120 = 120
     x130 = 130
     x140 = 140
     x150 = 150
     x160 = 160
     x170 = 170
     x180 = 180
     x190 = 190
     x200 = 200
End Enum
Const m_def_Print_Zoomfactor = 100
Dim m_Print_Zoomfactor As EnumPrint_Zoomfactor
Dim Pagina_Hoogte
Dim Pagina_Breedte
Dim Pagina_Seperator

Enum EnumShadowstyle
     None = 0
     XSmall = 4
     Small = 8
     Medium = 14
     Large = 20
     XLarge = 30
End Enum

Enum EnumPrintThumbsPerCol
     One_Pic_Row = 1
     Two_Pics_Row = 2
     Three_Pics_Row = 3
     Foure_Pics_Row = 4
     Five_Pics_Row = 5
     Six_Pics_Row = 6
     Seven_Pics_Row = 7
     Eight_Pics_Row = 8
     Nine_Pics_Row = 9
     Ten_Pics_Row = 10
End Enum
Const m_def_Print_NrThumbCols = 5
Dim m_Print_NrThumbCols As Integer
Dim PrintPreviewMode As Boolean
Dim HtmlPreviewmode As Boolean

Enum ThumbSize
    px_0 = 0
    px_10 = 10
    px_20 = 20
    px_30 = 30
    px_40 = 40
    px_50 = 50
    px_60 = 60
    px_70 = 70
    px_80 = 80
    px_90 = 90
    px_100 = 100
    px_110 = 110
    px_120 = 120
    px_130 = 130
    px_140 = 140
    px_150 = 150
    px_160 = 160
    px_170 = 170
    px_180 = 180
    px_190 = 190
    px_200 = 200
    px_210 = 210
    px_220 = 220
    px_230 = 230
    px_240 = 240
    px_250 = 250
    px_260 = 260
    px_270 = 270
    px_280 = 280
    px_290 = 290
    px_300 = 300
End Enum
Dim m_ThumbNailSize As ThumbSize
Const m_def_ThumbNailSize = 60

Enum EnumPresetColor
   None = 0
   Black = 1
   Antraciet = 2
   Dark_Grey = 3
   Middle_Grey = 4
   Light_Grey = 5
   White_Grey = 6
End Enum
Const m_def_ThumbPageColorPresets = 0
Dim m_ThumbPageColorPresets As EnumPresetColor

'----------------
Dim Thumb_Size
Dim Thumb_Width
Dim Thumb_Height


Dim Thumbs_PerRij
Dim Bezig_met_Laden As Boolean
Dim m_ResetLoading As Boolean
Dim ItemsGeladen

Dim m_InfobarVisible As Boolean

Dim m_CompressToJPG As Boolean
Dim m_Compress_NO_GIF As Boolean

Dim m_DestinationPath
Dim OpenFnr As Integer
Dim TextPlaced As Boolean
Dim Html_BodyText As String
Dim Html_CustomHeaderlinks As String
Dim HtmlThumbFilenames()

Dim BinnenMarge
Dim BuitenMarge

Dim ThumbFramesLoaded ' pagina`s geladen
Dim ThumbPagesLoaded
Dim ThumbsPerPage

Dim Can_Print As Boolean

Dim m_ThumbNailColor As OLE_COLOR

Const m_def_Html_ThumbnailBorderThicknes = 0
Const m_def_Html_ThumbnailSize = 200

Dim m_Html_ThumbnailBorderThicknes As EnumShadowstyle
Dim m_Html_ThumbnailSize As ThumbSize

Dim FnameClicked
Dim CurrentThumbnail
Dim PreviousThumbnail
Dim PreviousPage

Const m_def_ThumbPageColor = 0
Const m_def_ThumbNailSelectColor = &HFF00& ' Felgroen
Const m_def_ThumbNailBorderColor = 0

Dim m_ThumbPageColor As OLE_COLOR
Dim m_ThumbNailSelectColor As OLE_COLOR
Dim m_ThumbNailBorderColor As OLE_COLOR
Dim m_ThumbPage_FrameColor As OLE_COLOR

Const m_def_Html_ThumbnailBorderColor = &H808080
Const m_def_Html_PageBackcolor = &H353535
Const m_def_Html_PageForecolor = &H808080
Const m_def_Html_PageLineColor = &H330099
Const m_def_Html_PageLinkColor = &HC0C0C0
Const m_def_Html_PageLinkHoverForeColor = &HFFFFFF
Const m_def_Html_PageLinkHoverBackColor = &H330099
Const m_def_Html_PageFotoframeBackColor = &H575757
Const m_def_Html_PageFotoframeBorderColor = &H808080
Const m_def_Html_PageFotoLinkstyle = 0
Const m_def_Html_PageColorPresets = 0

Dim m_Html_ThumbnailBorderColor As OLE_COLOR
Dim m_Html_PageBackcolor As OLE_COLOR
Dim m_Html_PageForecolor As OLE_COLOR
Dim m_Html_PageLineColor As OLE_COLOR
Dim m_Html_PageLinkColor As OLE_COLOR
Dim m_Html_PageLinkHoverForeColor As OLE_COLOR
Dim m_Html_PageLinkHoverBackColor As OLE_COLOR
Dim m_Html_PageFotoframeBackColor As OLE_COLOR
Dim m_Html_PageFotoframeBorderColor As OLE_COLOR
Dim m_Html_PageFotoLinkstyle As EnumLang
Dim m_Html_PageColorPresets As EnumPresetColor
'Default Property Values:
Const m_def_ThumbnailExtraWidth = 0
Const m_def_ThumbnailExtraHeight = 0
Const m_def_MaxThumbnailpages = 100
Const m_def_FloodFillColor = &HFF
Const m_def_AutoUpdateOnPathchange = False
'Property Variables:
Dim m_ThumbnailExtraWidth As ThumbSize
Dim m_ThumbnailExtraHeight As ThumbSize
Dim m_MaxThumbnailpages As Long
Dim m_FloodFillColor As OLE_COLOR
Dim m_AutoUpdateOnPathchange As Boolean


Event ThumbClick(Index As Long, Filename As String)
Event ThumbDblClick(Index As Long, Filename As String)
Event PathChange(Path As String)
Event HtmLGaleryCreated(Sucses As Boolean, PicsGenerated As Long, TimeUsed As String)
Event ThumbPageStatus(ThumbNailSize, ThumbsPerPage, PagesNeeded, nmfiles)




'############## U S E R C O N T R O L  -  P R O P E R T Y ' S ###############

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    File1.Path = PropBag.ReadProperty("Path", "")
    m_ThumbPage_FrameColor = PropBag.ReadProperty("ThumbPage_FrameColor", &H404040)
    LbInfo.BackColor = PropBag.ReadProperty("InfoBackColor", &H0&)
    LbInfo.Forecolor = PropBag.ReadProperty("InfoForeColor", &H8080&)
    LbInfo.Visible = PropBag.ReadProperty("InfobarVisible", True)
    m_Language = PropBag.ReadProperty("Language", m_def_Language)
    m_CompressToJPG = PropBag.ReadProperty("CompressToJPG", True)
    m_Compress_NO_GIF = PropBag.ReadProperty("Compress_NO_GIF", True)
    m_Print_NrThumbCols = PropBag.ReadProperty("Print_NrThumbCols", m_def_Print_NrThumbCols)
    m_Print_Zoomfactor = PropBag.ReadProperty("Print_Zoomfactor", m_def_Print_Zoomfactor)
    m_PrintResolution = PropBag.ReadProperty("PrintResolution", m_def_PrintResolution)
    m_ThumbNailSize = PropBag.ReadProperty("ThumbNailSize", m_def_ThumbNailSize)
    m_ThumbNailColor = PropBag.ReadProperty("ThumbNailColor", &H808080)
    m_ThumbPageColor = PropBag.ReadProperty("ThumbPageColor", m_def_ThumbPageColor)
    m_ThumbNailSelectColor = PropBag.ReadProperty("ThumbNailSelectColor", m_def_ThumbNailSelectColor)
    m_ThumbNailBorderColor = PropBag.ReadProperty("ThumbNailBorderColor", m_def_ThumbNailBorderColor)
    m_ThumbPageColorPresets = PropBag.ReadProperty("ThumbPageColorPresets", m_def_ThumbPageColorPresets)
    m_Html_ThumbnailBorderThicknes = PropBag.ReadProperty("Html_ThumbnailBorderThicknes", m_def_Html_ThumbnailBorderThicknes)
    m_Html_ThumbnailSize = PropBag.ReadProperty("Html_ThumbnailSize", m_def_Html_ThumbnailSize)
    m_Html_ThumbnailBorderColor = PropBag.ReadProperty("Html_ThumbnailBorderColor", m_def_Html_ThumbnailBorderColor)
    m_Html_PageBackcolor = PropBag.ReadProperty("Html_PageBackcolor", m_def_Html_PageBackcolor)
    m_Html_PageForecolor = PropBag.ReadProperty("Html_PageForecolor", m_def_Html_PageForecolor)
    m_Html_PageLineColor = PropBag.ReadProperty("Html_PageLineColor", m_def_Html_PageLineColor)
    m_Html_PageLinkColor = PropBag.ReadProperty("Html_PageLinkColor", m_def_Html_PageLinkColor)
    m_Html_PageLinkHoverForeColor = PropBag.ReadProperty("Html_PageLinkHoverForeColor", m_def_Html_PageLinkHoverForeColor)
    m_Html_PageLinkHoverBackColor = PropBag.ReadProperty("Html_PageLinkHoverBackColor", m_def_Html_PageLinkHoverBackColor)
    m_Html_PageFotoframeBackColor = PropBag.ReadProperty("Html_PageFotoframeBackColor", m_def_Html_PageFotoframeBackColor)
    m_Html_PageFotoframeBorderColor = PropBag.ReadProperty("Html_PageFotoframeBorderColor", m_def_Html_PageFotoframeBorderColor)
    m_Html_PageFotoLinkstyle = PropBag.ReadProperty("Html_PageFotoLinkstyle", m_def_Html_PageFotoLinkstyle)
    m_Html_PageColorPresets = PropBag.ReadProperty("Html_PageColorPresets", m_def_Html_PageColorPresets)
    '--------------
    ThumbPage(0).BackColor = m_ThumbPageColor
    ThumbInnerframe(0).BackColor = m_ThumbPage_FrameColor
    ThumbOuterFrame.BackColor = m_ThumbPage_FrameColor
    m_AutoUpdateOnPathchange = PropBag.ReadProperty("AutoUpdateOnPathchange", m_def_AutoUpdateOnPathchange)
    m_FloodFillColor = PropBag.ReadProperty("FloodFillColor", m_def_FloodFillColor)
    m_MaxThumbnailpages = PropBag.ReadProperty("MaxThumbnailpages", m_def_MaxThumbnailpages)
    m_ThumbnailExtraWidth = PropBag.ReadProperty("ThumbnailExtraWidth", m_def_ThumbnailExtraWidth)
    m_ThumbnailExtraHeight = PropBag.ReadProperty("ThumbnailExtraHeight", m_def_ThumbnailExtraHeight)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Path", File1.Path, "")
    Call PropBag.WriteProperty("ThumbPage_FrameColor", m_ThumbPage_FrameColor, &H404040)
    Call PropBag.WriteProperty("ThumbNailColor", m_ThumbNailColor, &H808080)
    Call PropBag.WriteProperty("InfoBackColor", LbInfo.BackColor, &H0&)
    Call PropBag.WriteProperty("InfoForeColor", LbInfo.Forecolor, &H8080&)
    Call PropBag.WriteProperty("InfobarVisible", LbInfo.Visible, True)
    Call PropBag.WriteProperty("ThumbNailSize", m_ThumbNailSize, m_def_ThumbNailSize)
    Call PropBag.WriteProperty("PrintResolution", m_PrintResolution, m_def_PrintResolution)
    Call PropBag.WriteProperty("Language", m_Language, m_def_Language)
    Call PropBag.WriteProperty("CompressToJPG", m_CompressToJPG, True)
    Call PropBag.WriteProperty("Compress_NO_GIF", m_Compress_NO_GIF, True)
    Call PropBag.WriteProperty("Print_NrThumbCols", m_Print_NrThumbCols, m_def_Print_NrThumbCols)
    Call PropBag.WriteProperty("Print_Zoomfactor", m_Print_Zoomfactor, m_def_Print_Zoomfactor)
    Call PropBag.WriteProperty("Html_ThumbnailBorderThicknes", m_Html_ThumbnailBorderThicknes, m_def_Html_ThumbnailBorderThicknes)
    Call PropBag.WriteProperty("Html_ThumbnailSize", m_Html_ThumbnailSize, m_def_Html_ThumbnailSize)
    Call PropBag.WriteProperty("ThumbPageColor", m_ThumbPageColor, m_def_ThumbPageColor)
    Call PropBag.WriteProperty("ThumbNailSelectColor", m_ThumbNailSelectColor, m_def_ThumbNailSelectColor)
    Call PropBag.WriteProperty("ThumbNailBorderColor", m_ThumbNailBorderColor, m_def_ThumbNailBorderColor)
    Call PropBag.WriteProperty("ThumbPageColorPresets", m_ThumbPageColorPresets, m_def_ThumbPageColorPresets)
    Call PropBag.WriteProperty("Html_ThumbnailBorderColor", m_Html_ThumbnailBorderColor, m_def_Html_ThumbnailBorderColor)
    Call PropBag.WriteProperty("Html_PageBackcolor", m_Html_PageBackcolor, m_def_Html_PageBackcolor)
    Call PropBag.WriteProperty("Html_PageForecolor", m_Html_PageForecolor, m_def_Html_PageForecolor)
    Call PropBag.WriteProperty("Html_PageLineColor", m_Html_PageLineColor, m_def_Html_PageLineColor)
    Call PropBag.WriteProperty("Html_PageLinkColor", m_Html_PageLinkColor, m_def_Html_PageLinkColor)
    Call PropBag.WriteProperty("Html_PageLinkHoverForeColor", m_Html_PageLinkHoverForeColor, m_def_Html_PageLinkHoverForeColor)
    Call PropBag.WriteProperty("Html_PageLinkHoverBackColor", m_Html_PageLinkHoverBackColor, m_def_Html_PageLinkHoverBackColor)
    Call PropBag.WriteProperty("Html_PageFotoframeBackColor", m_Html_PageFotoframeBackColor, m_def_Html_PageFotoframeBackColor)
    Call PropBag.WriteProperty("Html_PageFotoframeBorderColor", m_Html_PageFotoframeBorderColor, m_def_Html_PageFotoframeBorderColor)
    Call PropBag.WriteProperty("Html_PageFotoLinkstyle", m_Html_PageFotoLinkstyle, m_def_Html_PageFotoLinkstyle)
    Call PropBag.WriteProperty("Html_PageColorPresets", m_Html_PageColorPresets, m_def_Html_PageColorPresets)
    Call PropBag.WriteProperty("AutoUpdateOnPathchange", m_AutoUpdateOnPathchange, m_def_AutoUpdateOnPathchange)
    Call PropBag.WriteProperty("FloodFillColor", m_FloodFillColor, m_def_FloodFillColor)
    Call PropBag.WriteProperty("MaxThumbnailpages", m_MaxThumbnailpages, m_def_MaxThumbnailpages)
    Call PropBag.WriteProperty("ThumbnailExtraWidth", m_ThumbnailExtraWidth, m_def_ThumbnailExtraWidth)
    Call PropBag.WriteProperty("ThumbnailExtraHeight", m_ThumbnailExtraHeight, m_def_ThumbnailExtraHeight)
End Sub

Public Property Get ThumbPage_FrameColor() As OLE_COLOR
    ThumbPage_FrameColor = m_ThumbPage_FrameColor
End Property
        Public Property Let ThumbPage_FrameColor(ByVal New_ThumbPage_FrameColor As OLE_COLOR)
            On Error Resume Next
            m_ThumbPage_FrameColor = New_ThumbPage_FrameColor
            ThumbOuterFrame.BackColor = New_ThumbPage_FrameColor
            ThumbInnerframe(0).BackColor = New_ThumbPage_FrameColor
            PropertyChanged "ThumbPage_FrameColor"
            If Ambient.UserMode = False Then Call Makegallery
        End Property

Public Property Get ThumbNailColor() As OLE_COLOR
    ThumbNailColor = m_ThumbNailColor
End Property
        Public Property Let ThumbNailColor(ByVal New_ThumbNailColor As OLE_COLOR)
            m_ThumbNailColor = New_ThumbNailColor
            PropertyChanged "ThumbNailColor"
            If Ambient.UserMode = False Then Call Makegallery
        End Property

Public Property Get ThumbPageColor() As OLE_COLOR
    ThumbPageColor = m_ThumbPageColor
End Property
        Public Property Let ThumbPageColor(ByVal New_ThumbPageColor As OLE_COLOR)
            m_ThumbPageColor = New_ThumbPageColor
            ThumbPage(0).BackColor = New_ThumbPageColor
            PropertyChanged "ThumbPageColor"
            If Ambient.UserMode = False Then Call Makegallery
        End Property

Public Property Get ThumbNailSelectColor() As OLE_COLOR
    ThumbNailSelectColor = m_ThumbNailSelectColor
End Property
        Public Property Let ThumbNailSelectColor(ByVal New_ThumbNailSelectColor As OLE_COLOR)
            m_ThumbNailSelectColor = New_ThumbNailSelectColor
            PropertyChanged "ThumbNailSelectColor"
            If Ambient.UserMode = False Then Call Makegallery
        End Property

Public Property Get ThumbNailBorderColor() As OLE_COLOR
    ThumbNailBorderColor = m_ThumbNailBorderColor
End Property
        Public Property Let ThumbNailBorderColor(ByVal New_ThumbNailBorderColor As OLE_COLOR)
            m_ThumbNailBorderColor = New_ThumbNailBorderColor
            PropertyChanged "ThumbNailBorderColor"
            If Ambient.UserMode = False Then Call Makegallery
        End Property

Public Property Get InfoBackColor() As OLE_COLOR
    InfoBackColor = LbInfo.BackColor
End Property
        Public Property Let InfoBackColor(ByVal New_InfoBackColor As OLE_COLOR)
            LbInfo.BackColor = New_InfoBackColor
            PropertyChanged "InfoBackColor"
        End Property

Public Property Get InfoForeColor() As OLE_COLOR
    InfoForeColor = LbInfo.Forecolor
End Property
        Public Property Let InfoForeColor(ByVal New_InfoForeColor As OLE_COLOR)
            LbInfo.Forecolor = New_InfoForeColor
            PropertyChanged "InfoForeColor"
        End Property

Public Property Get CompressToJPG() As Boolean
    CompressToJPG = m_CompressToJPG
End Property
        Public Property Let CompressToJPG(ByVal New_CompressToJPG As Boolean)
            m_CompressToJPG = New_CompressToJPG
            PropertyChanged "CompressToJPG"
        End Property

Public Property Get Compress_NO_GIF() As Boolean
    Compress_NO_GIF = m_Compress_NO_GIF
End Property
        Public Property Let Compress_NO_GIF(ByVal New_Compress_NO_GIF As Boolean)
            m_Compress_NO_GIF = New_Compress_NO_GIF
            PropertyChanged "Compress_NO_GIF"
        End Property
    
Public Property Get InfobarVisible() As Boolean
    InfobarVisible = LbInfo.Visible
End Property
        Public Property Let InfobarVisible(ByVal New_InfobarVisible As Boolean)
            m_InfobarVisible = New_InfobarVisible
            LbInfo.Visible = m_InfobarVisible
            PropertyChanged "InfobarVisible"
            Call UserControl_Resize
        End Property

Public Property Get PrintResolution() As EnumPrintResolution
    PrintResolution = m_PrintResolution
End Property
        Public Property Let PrintResolution(ByVal New_PrintResolution As EnumPrintResolution)
            m_PrintResolution = New_PrintResolution
            PropertyChanged "PrintResolution"
        End Property

Public Property Get ThumbNailSize() As ThumbSize
    ThumbNailSize = m_ThumbNailSize
End Property
        Public Property Let ThumbNailSize(ByVal New_ThumbNailSize As ThumbSize)
            m_ThumbNailSize = New_ThumbNailSize
            PropertyChanged "ThumbNailSize"
            If Ambient.UserMode = False Then Call Makegallery
        End Property

Public Property Get Print_NrThumbCols() As EnumPrintThumbsPerCol
    Print_NrThumbCols = m_Print_NrThumbCols
End Property
        Public Property Let Print_NrThumbCols(ByVal New_Print_NrThumbCols As EnumPrintThumbsPerCol)
            m_Print_NrThumbCols = New_Print_NrThumbCols
            PropertyChanged "Print_NrThumbCols"
            If Bezig_met_Laden = True Then m_ResetLoading = True
        End Property


Public Property Get Print_Zoomfactor() As EnumPrint_Zoomfactor
    Print_Zoomfactor = m_Print_Zoomfactor
End Property
        Public Property Let Print_Zoomfactor(ByVal New_Print_Zoomfactor As EnumPrint_Zoomfactor)
            If New_Print_Zoomfactor >= 10 And New_Print_Zoomfactor <= 200 Then
               m_Print_Zoomfactor = New_Print_Zoomfactor
               PropertyChanged "Print_Zoomfactor"
               If Ambient.UserMode = False Then
                  If Bezig_met_Laden = True Then m_ResetLoading = True
                  Call PrintPreview
               End If
            End If
        End Property

Public Property Get Html_ThumbnailBorderThicknes() As EnumShadowstyle
    Html_ThumbnailBorderThicknes = m_Html_ThumbnailBorderThicknes
End Property
        Public Property Let Html_ThumbnailBorderThicknes(ByVal New_Html_ThumbnailBorderThicknes As EnumShadowstyle)
            m_Html_ThumbnailBorderThicknes = New_Html_ThumbnailBorderThicknes
            PropertyChanged "Html_ThumbnailBorderThicknes"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Html_ThumbnailSize() As ThumbSize
    Html_ThumbnailSize = m_Html_ThumbnailSize
End Property
        Public Property Let Html_ThumbnailSize(ByVal New_Html_ThumbnailSize As ThumbSize)
            m_Html_ThumbnailSize = New_Html_ThumbnailSize
            PropertyChanged "Html_ThumbnailSize"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Html_ThumbnailBorderColor() As OLE_COLOR
    Html_ThumbnailBorderColor = m_Html_ThumbnailBorderColor
End Property
        Public Property Let Html_ThumbnailBorderColor(ByVal New_Html_ThumbnailBorderColor As OLE_COLOR)
            m_Html_ThumbnailBorderColor = New_Html_ThumbnailBorderColor
            PropertyChanged "Html_ThumbnailBorderColor"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Html_PageBackcolor() As OLE_COLOR
    Html_PageBackcolor = m_Html_PageBackcolor
End Property
        Public Property Let Html_PageBackcolor(ByVal New_Html_PageBackcolor As OLE_COLOR)
            m_Html_PageBackcolor = New_Html_PageBackcolor
            PropertyChanged "Html_PageBackcolor"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Html_PageForecolor() As OLE_COLOR
    Html_PageForecolor = m_Html_PageForecolor
End Property
        Public Property Let Html_PageForecolor(ByVal New_Html_PageForecolor As OLE_COLOR)
            m_Html_PageForecolor = New_Html_PageForecolor
            PropertyChanged "Html_PageForecolor"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Html_PageLineColor() As OLE_COLOR
    Html_PageLineColor = m_Html_PageLineColor
End Property
        Public Property Let Html_PageLineColor(ByVal New_Html_PageLineColor As OLE_COLOR)
            m_Html_PageLineColor = New_Html_PageLineColor
            PropertyChanged "Html_PageLineColor"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Html_PageLinkColor() As OLE_COLOR
    Html_PageLinkColor = m_Html_PageLinkColor
End Property
        Public Property Let Html_PageLinkColor(ByVal New_Html_PageLinkColor As OLE_COLOR)
            m_Html_PageLinkColor = New_Html_PageLinkColor
            PropertyChanged "Html_PageLinkColor"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Html_PageLinkHoverForeColor() As OLE_COLOR
    Html_PageLinkHoverForeColor = m_Html_PageLinkHoverForeColor
End Property
        Public Property Let Html_PageLinkHoverForeColor(ByVal New_Html_PageLinkHoverForeColor As OLE_COLOR)
            m_Html_PageLinkHoverForeColor = New_Html_PageLinkHoverForeColor
            PropertyChanged "Html_PageLinkHoverForeColor"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Html_PageLinkHoverBackColor() As OLE_COLOR
    Html_PageLinkHoverBackColor = m_Html_PageLinkHoverBackColor
End Property
        Public Property Let Html_PageLinkHoverBackColor(ByVal New_Html_PageLinkHoverBackColor As OLE_COLOR)
            m_Html_PageLinkHoverBackColor = New_Html_PageLinkHoverBackColor
            PropertyChanged "Html_PageLinkHoverBackColor"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Html_PageFotoframeBackColor() As OLE_COLOR
    Html_PageFotoframeBackColor = m_Html_PageFotoframeBackColor
End Property
        Public Property Let Html_PageFotoframeBackColor(ByVal New_Html_PageFotoframeBackColor As OLE_COLOR)
            m_Html_PageFotoframeBackColor = New_Html_PageFotoframeBackColor
            PropertyChanged "Html_PageFotoframeBackColor"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Html_PageFotoframeBorderColor() As OLE_COLOR
    Html_PageFotoframeBorderColor = m_Html_PageFotoframeBorderColor
End Property
    Public Property Let Html_PageFotoframeBorderColor(ByVal New_Html_PageFotoframeBorderColor As OLE_COLOR)
        m_Html_PageFotoframeBorderColor = New_Html_PageFotoframeBorderColor
        PropertyChanged "Html_PageFotoframeBorderColor"
        If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
    End Property

Public Property Get Html_PageFotoLinkstyle() As EnumLang
    Html_PageFotoLinkstyle = m_Html_PageFotoLinkstyle
End Property
        Public Property Let Html_PageFotoLinkstyle(ByVal New_Html_PageFotoLinkstyle As EnumLang)
            m_Html_PageFotoLinkstyle = New_Html_PageFotoLinkstyle
            PropertyChanged "Html_PageFotoLinkstyle"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Html_PageColorPresets() As EnumPresetColor
    Html_PageColorPresets = m_Html_PageColorPresets
End Property
        Public Property Let Html_PageColorPresets(ByVal New_Html_PageColorPresets As EnumPresetColor)
            m_Html_PageColorPresets = New_Html_PageColorPresets
            ' KOMT NOG
            'Call Set_Html_PageColorpresets
            PropertyChanged "Html_PageColorPresets"
            If Ambient.UserMode = False Then Call Html_Thumbnail_Preview
        End Property

Public Property Get Path() As String
    Path = File1.Path
End Property
        Public Property Let Path(ByVal New_Path As String)
            File1.Path = New_Path
            PropertyChanged "Path"
            If m_AutoUpdateOnPathchange = True Then Call Makegallery
        End Property

Public Property Get Language() As EnumLang
    Language = m_Language
End Property
        Public Property Let Language(ByVal New_language As EnumLang)
            m_Language = New_language
            PropertyChanged "Language"
        End Property

Public Property Get ThumbPageColorPresets() As EnumPresetColor
    ThumbPageColorPresets = m_ThumbPageColorPresets
End Property
        Public Property Let ThumbPageColorPresets(ByVal New_ThumbPageColorPresets As EnumPresetColor)
            m_ThumbPageColorPresets = New_ThumbPageColorPresets
            Call SetPresetcolors(m_ThumbPageColorPresets)
            PropertyChanged "ThumbPageColorPresets"
        End Property
Public Property Get AutoUpdateOnPathchange() As Boolean
    AutoUpdateOnPathchange = m_AutoUpdateOnPathchange
End Property
        Public Property Let AutoUpdateOnPathchange(ByVal New_AutoUpdateOnPathchange As Boolean)
            m_AutoUpdateOnPathchange = New_AutoUpdateOnPathchange
            PropertyChanged "AutoUpdateOnPathchange"
        End Property
Public Property Get FloodFillColor() As OLE_COLOR
    FloodFillColor = m_FloodFillColor
End Property
        Public Property Let FloodFillColor(ByVal New_FloodFillColor As OLE_COLOR)
            m_FloodFillColor = New_FloodFillColor
            PropertyChanged "FloodFillColor"
            If Ambient.UserMode = False Then Call Makegallery
        End Property
Public Property Get MaxThumbnailpages() As Long
    MaxThumbnailpages = m_MaxThumbnailpages
End Property
        Public Property Let MaxThumbnailpages(ByVal New_MaxThumbnailpages As Long)
            m_MaxThumbnailpages = New_MaxThumbnailpages
            PropertyChanged "MaxThumbnailpages"
        End Property



'################### U S E R C O N T R O L  -  E V E N T S #######################

Private Sub UserControl_Resize()
  BinnenMarge = 5
  BuitenMarge = 0
  
  LbInfo.Top = 0
  LbInfo.Left = 0
  Gradient.Left = 0
  Gradient.Top = LbInfo.Height - Gradient.Height
  '===============
  If LbInfo.Visible = False Then
     ThumbOuterFrame.Top = 0
     ThumbOuterFrame.Left = 0
     ThumbOuterFrame.Height = UserControl.ScaleHeight - BuitenMarge
  Else
     ThumbOuterFrame.Top = LbInfo.Height
     ThumbOuterFrame.Left = 0
     ThumbOuterFrame.Height = (UserControl.ScaleHeight - LbInfo.Height) - BuitenMarge
  End If
  ThumbOuterFrame.Width = (UserControl.ScaleWidth - V_Scroll.Width) - BuitenMarge
  LbInfo.Width = ThumbOuterFrame.Width + V_Scroll.Width
  Gradient.Width = LbInfo.Width
  '-----
  ThumbInnerframe(0).Left = BinnenMarge
  ThumbInnerframe(0).Top = BinnenMarge
  If ThumbInnerframe(0).Height < ThumbOuterFrame.Height - (BinnenMarge * 2.5) Then ThumbInnerframe(0).Height = ThumbOuterFrame.Height - (BinnenMarge * 2.5)
  ThumbInnerframe(0).Width = ThumbOuterFrame.Width - (BinnenMarge * 2.5)
  '-----
  V_Scroll.Top = ThumbOuterFrame.Top
  V_Scroll.Left = ThumbOuterFrame.Width
  V_Scroll.Height = ThumbOuterFrame.Height
  '-----
  If Ambient.UserMode = False Then
     Call Makegallery
  End If
End Sub
Private Sub UserControl_InitProperties()
    m_ThumbNailSize = m_def_ThumbNailSize
    m_PrintResolution = m_def_PrintResolution
    m_Print_NrThumbCols = m_def_Print_NrThumbCols
    m_Print_Zoomfactor = m_def_Print_Zoomfactor
    m_Html_ThumbnailBorderThicknes = m_def_Html_ThumbnailBorderThicknes
'    m_Html_ThumbnailBorderColor = m_def_Html_ThumbnailBorderColor
    m_Html_ThumbnailSize = m_def_Html_ThumbnailSize
    m_ThumbPageColor = m_def_ThumbPageColor
    m_ThumbNailSelectColor = m_def_ThumbNailSelectColor
    m_ThumbNailBorderColor = m_def_ThumbNailBorderColor
    m_ThumbPageColorPresets = m_def_ThumbPageColorPresets
    m_Html_ThumbnailBorderColor = m_def_Html_ThumbnailBorderColor
    m_Html_PageBackcolor = m_def_Html_PageBackcolor
    m_Html_PageForecolor = m_def_Html_PageForecolor
    m_Html_PageLineColor = m_def_Html_PageLineColor
    m_Html_PageLinkColor = m_def_Html_PageLinkColor
    m_Html_PageLinkHoverForeColor = m_def_Html_PageLinkHoverForeColor
    m_Html_PageLinkHoverBackColor = m_def_Html_PageLinkHoverBackColor
    m_Html_PageFotoframeBackColor = m_def_Html_PageFotoframeBackColor
    m_Html_PageFotoframeBorderColor = m_def_Html_PageFotoframeBorderColor
    m_Html_PageFotoLinkstyle = m_def_Html_PageFotoLinkstyle
    m_Html_PageColorPresets = m_def_Html_PageColorPresets
    m_AutoUpdateOnPathchange = m_def_AutoUpdateOnPathchange
    m_FloodFillColor = m_def_FloodFillColor
    m_MaxThumbnailpages = m_def_MaxThumbnailpages
    m_ThumbnailExtraWidth = m_def_ThumbnailExtraWidth
    m_ThumbnailExtraHeight = m_def_ThumbnailExtraHeight
End Sub

' ############   H U L P  -  R O U T I N E 'S #############


Public Function Convert_Decimal_To_Hex(Number As Long, IsHtmlHex As Boolean) As String
    Dim DevideTree As Long
    Dim Alpha As String
    Dim Base As Single
    Dim lTemp As Single
 
    DevideTree = 1
    Do Until DevideTree > Number
       DevideTree = DevideTree * 16
    Loop
    DevideTree = DevideTree \ 16
    Do Until DevideTree = 0
       lTemp = Number \ DevideTree
       Number = Number - (DevideTree * lTemp)
       Alpha = lTemp
       DTHEX = DTHEX & GetHex(Alpha)
       DevideTree = DevideTree \ 16
    Loop
    ' ================
    If IsHtmlHex = True Then
        lengte = Len(Trim(DTHEX))
        Select Case lengte
        Case 2: DTHEX = DTHEX & "0000"
        Case 4:
              links = Left$(DTHEX, 2)
              Rechts = Right$(DTHEX, 2)
              DTHEX = Rechts & links & "00"
        Case 6: DTHEX = DraaiHexOm(DTHEX)
        Case Else
              DTHEX = "000000"
        End Select
    Else
        lengte = Len(Trim(DTHEX))
        Select Case lengte
        Case 0: DTHEX = "000000"
        Case 2: DTHEX = "0000" & DTHEX
        Case 4: DTHEX = DTHEX & "00"
        Case Else
        End Select
    End If
   Convert_Decimal_To_Hex = DTHEX
End Function
Private Function GetHex(Number As String)
    Select Case Number
    Case 10:    GetHex = "A"
    Case 11:    GetHex = "B"
    Case 12:    GetHex = "C"
    Case 13:    GetHex = "D"
    Case 14:    GetHex = "E"
    Case 15:    GetHex = "F"
    Case Else:  GetHex = Number
    End Select
End Function
Private Function DraaiHexOm(HexString_in)
  Tmp = Trim(HexString_in)
  links = Left(Tmp, 2)
  Midden = Mid(Tmp, 3, 2)
  Rechts = Right(Tmp, 2)
  DraaiHexOm = Rechts & Midden & links
End Function
Private Function CheckStriem(Variabele)
   If Left$(Variabele, 1) <> "#" Then
      V2 = "#" & Variabele
      Variabele = V2
   End If
   CheckStriem = Variabele
End Function

Public Function SystemResources()
    SystemResources = GetFreeSystemResources(GFSR_SYSTEMRESOURCES)
End Function

Private Function Dir_CreateDir_2(strDir As String) As Boolean
On Error Resume Next
    Dim bytMax As Byte
    Dim bytNdx As Byte
    Dim strDirLevel As String
    If Right(strDir, 1) <> "\" Then
       strDir = strDir & "\"
    End If
    bytMax = Len(strDir)
    For bytNdx = 4 To bytMax
        If (Mid(strDir, bytNdx, 1) = "\") Then
            strDirLevel = Left(strDir, bytNdx - 1)
            If Dir(strDirLevel, vbDirectory) = "" Then
               MkDir strDirLevel
            End If
        End If
    Next
    If Dir(strDir, vbDirectory) <> "" Then
       Dir_CreateDir_2 = True ' Succeeded creating directory
    Else
       Dir_CreateDir_2 = False ' Failed creating directory
    End If
End Function

Private Function Dir_GetLastSubdir(s As String)
   If Right$(s, 1) <> "\" Then s = s & "\"
   For I = 1 To Len(s)
       If Mid(s, I, 1) = "\" Then t = t + 1
   Next
   pos = t
   t = 0
   For I = 1 To Len(s)
       Char$ = Mid(s, I, 1)
       If Char$ = "\" Then
          t = t + 1
          If t <> 1 Then Sign$ = Sign$ & "   "
       End If
       If t = pos - 1 Then ok = 1
       If t = pos Then ok = 2
       If ok = 1 Then
          If Char$ <> "\" Then ts$ = ts$ & Char$
       End If
   Next
   Dir_GetLastSubdir = Sign$ & ts$
End Function

Private Function LegalFilename(Name_In)
Dim temp
 temp = Trim(Name_In): Name_In = temp
 For I = 1 To Len(Name_In)
   Char$ = Mid(Name_In, I, 1)
   If Char$ = " " Then Char$ = "_"
   If InStr("\/:;*?,<>|%+=" & Chr$(34) & Space$(1), Char$) <> 0 Then Char$ = "_"
   t$ = t$ & Char$
   If Len(t$) > 210 Then Exit For
 Next
 LegalFilename = t$
End Function


Sub RunFile(ByVal File, FilePath, RunStyle)
    Const MB_ICONSTOP = 16
    Dim temp, Msg As String
    Dim X
    temp = GetActiveWindow()
    X = ShellExecute(temp, "Open", File, "", FilePath, RunStyle)
    If X < 32 Then MsgBox "Fout met openen van de browser !!"
End Sub

Private Sub ThumbPage_DblClick(Index As Integer)
  If PrintPreviewMode = False Then
     If FnameClicked <> "" Then
        RaiseEvent ThumbDblClick(CLng(CurrentThumbnail), CStr(FnameClicked))
        FnameClicked = ""
     End If
  End If
End Sub

Private Sub ThumbPage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    If PrintPreviewMode = True Then Exit Sub
    If HtmlPreviewmode = True Then Exit Sub
     
     P = PreviousThumbnail
     ThumbPage(PreviousPage).Line (Thumb_Index(P, ThumbID.X_1) - 1, Thumb_Index(P, ThumbID.y_1) - 1)-(Thumb_Index(P, ThumbID.X_2) + 1, Thumb_Index(P, ThumbID.y_2) + 1), m_ThumbNailBorderColor, B
  
     For I = 0 To File1.ListCount - 1
         If Thumb_Index(I, ThumbID.pagina) = Index Then
            X1 = Thumb_Index(I, ThumbID.X_1)
            Y1 = Thumb_Index(I, ThumbID.y_1)
            X2 = Thumb_Index(I, ThumbID.X_2)
            Y2 = Thumb_Index(I, ThumbID.y_2)
            If X >= X1 And X <= X2 And y >= Y1 And y <= Y2 Then
               FnameClicked = Thumb_Index(I, ThumbID.File_name)
               CurrentThumbnail = I
               PreviousThumbnail = I
               PreviousPage = Index
               ThumbPage(Index).Line (Thumb_Index(I, ThumbID.X_1) - 1, Thumb_Index(I, ThumbID.y_1) - 1)-(Thumb_Index(I, ThumbID.X_2) + 1, Thumb_Index(I, ThumbID.y_2) + 1), m_ThumbNailSelectColor, B
               RaiseEvent ThumbClick(CLng(CurrentThumbnail), CStr(FnameClicked))
               Exit For
            End If
         End If
     Next
     If Button = 1 Then
     
     End If
End Sub

Private Sub V_Scroll_Change()
       ThumbInnerframe(0).Top = -V_Scroll.Value
End Sub
Private Sub V_Scroll_Scroll()
    V_Scroll_Change
End Sub


'######################################################################################################################

Private Sub File1_PathChange()
 RaiseEvent PathChange(File1.Path)
End Sub


'###################### H T M L - G A L E R Y - M A K E R ##########################

Public Sub Html_Thumbnail_Preview()
  HtmlPreviewmode = True
   
  On Error Resume Next
   
' m_Html_ThumbnailBorderColor
' m_Html_PageBackcolor
' m_Html_PageForecolor
' m_Html_PageLineColor
' m_Html_PageLinkColor
' m_Html_PageLinkHoverForeColor
' m_Html_PageLinkHoverBackColor
' m_Html_PageFotoframeBackColor
' m_Html_PageFotoframeBorderColor
  
  Call Unload_ThumbInnerframes
  Call Unload_ThumbPages
  
  ThumbInnerframe(0).Visible = True
  ThumbOuterFrame.BackColor = m_Html_PageBackcolor
  ThumbInnerframe(0).BackColor = m_Html_PageBackcolor
  ThumbPage(0).Visible = False
  
  
  Aantal = 8  ' 5 thumbnailpreviews
  afstand = 3 ' 3 pixels tussenruimte
  Bovenmarge = 20
  FrameRand = 13
  celgroote = (m_Html_ThumbnailSize + FrameRand) + m_Html_ThumbnailBorderThicknes
  links = afstand
  '---------------------------------
  ThumbInnerframe(0).CurrentX = links
  ThumbInnerframe(0).CurrentY = links
  ThumbInnerframe(0).Forecolor = m_Html_PageLinkColor
  ThumbInnerframe(0).Print " vorige [ index ] volgende  "
  
  For I = 1 To Aantal
      If links = afstand Then ' .. pixels van de kantlijn, plaatje printen
         y = Bovenmarge + ((afstand + celgroote) * tel)
         Call Load_ThumbPage(I, links, y, celgroote, celgroote)
         ThumbPage(I).Visible = True
         ThumbPage(I).BackColor = m_Html_PageFotoframeBackColor
         '--------
         Call Verschaal_Picture(m_Html_ThumbnailSize, m_Html_ThumbnailSize, Verschaalde_Breedte, Verschaalde_Hoogte, m_Html_ThumbnailBorderThicknes)
         X1 = (celgroote \ 2) - (Verschaalde_Breedte \ 2)
         Y1 = (celgroote \ 2) - (Verschaalde_Hoogte \ 2)
         X2 = X1 + Verschaalde_Breedte
         Y2 = Y1 + Verschaalde_Hoogte
         BThk = m_Html_ThumbnailBorderThicknes
         '--------
         If BThk <> 0 Then ThumbPage(I).Line (X1 - BThk, Y1 - BThk)-(X2 + BThk, Y2 + BThk), m_Html_ThumbnailBorderColor, BF
         ThumbPage(I).PaintPicture HtmlPreviewpicture, X1, Y1, Verschaalde_Breedte, Verschaalde_Hoogte
         links = celgroote + (afstand * 2) ' naar rechts opschuiven voor de volgende lus
         tel = tel + 1
      Else
         Call Load_ThumbPage(I, links, y, celgroote, celgroote)
         ThumbPage(I).Visible = True
         ThumbPage(I).BackColor = m_Html_PageFotoframeBackColor
         links = afstand ' weer naar links opschuiven voor de volgende lus
      End If
      ' lichte lijntjes om de cellen
      ThumbPage(I).Line (0, 0)-(0, ThumbPage(I).Width), m_Html_PageFotoframeBorderColor
      ThumbPage(I).Line (0, 0)-(ThumbPage(I).Height, 0), m_Html_PageFotoframeBorderColor
  Next
  
End Sub

Public Sub MakeHtmlGallery(GalleryRoot, BodyText, CustomHeaderlinks)
  
  Gradient.Visible = True
  
  HtmlPreviewmode = True
  
  If GalleryRoot = "" Then GalleryRoot = App.Path
  If BodyText <> "" Then Html_BodyText = BodyText
  If CustomHeaderlinks <> "" Then Html_CustomHeaderlinks = CustomHeaderlinks
  If Dir_CreateDir_2(GalleryRoot & "\Gallery") = False Then Exit Sub
  m_DestinationPath = GalleryRoot & "\Gallery"
  m_Def_DestinationPath = m_DestinationPath & Mid(File1.Path, 3, Len(File1.Path))
  If Dir_CreateDir_2(CStr(m_Def_DestinationPath & "\OriginalPics")) = False Then
     MsgBox "Kan direcory : " & m_Def_DestinationPath & vbCrLf & "Niet maken"
     Exit Sub
  End If
  pad = File1.Path: If Right$(pad, 1) <> "\" Then pad = pad & "\"
    
  ThumbInnerframe(0).Visible = True
    
  m_ResetLoading = False
  Bezig_met_Laden = True
  
  Hoogte = ThumbInnerframe(0).Height = ThumbOuterFrame.Height - (BinnenMarge * 2.5)
  Breedte = ThumbInnerframe(0).Width = ThumbOuterFrame.Width - (BinnenMarge * 2.5)
  ThumbInnerframe(0).Top = 0 ' frame terugscrollen
  
  Aantal = File1.ListCount
  Call Master_Openpage(m_Def_DestinationPath)
  
  ReDim HtmlThumbFilenames(File1.ListCount - 1, 6)
  Laatste = ThumbPagesLoaded + 1
  Call Load_ThumbPage(Laatste, -200, -200, 1, 1) ' tijdelijke ThumbPage
  ThumbPage(Laatste).Visible = True
  ThumbPage(Laatste).BackColor = m_Html_ThumbnailBorderColor
  ThumbPage(Laatste).ZOrder 0
  
  marge = m_Html_ThumbnailBorderThicknes
  
  If m_Html_ThumbnailBorderColor > 0 Then
     'Als de fotoborder is ingesteld een extra page laden
     FotometBorder = Laatste + 1
     Call Load_ThumbPage(FotometBorder, -100, 0, 1, 1) ' tijdelijke ThumbPage
     ThumbPage(FotometBorder).Visible = False
     ThumbPage(FotometBorder).BackColor = m_Html_ThumbnailBorderColor
  End If
  
  '=============================================
  Screen.MousePointer = vbHourglass
  '---
  
  For I = 0 To Aantal - 1
      DoEvents
      If m_ResetLoading = True Then Exit For
      Fnm = File1.List(I)
      Bron_PadEnFile = pad & Fnm
      If Fnm = "" Then GoTo skip
      P_load.Picture = LoadPicture(Bron_PadEnFile)
      '---
      Call Verschaal_Picture(m_Html_ThumbnailSize, m_Html_ThumbnailSize, Verschaalde_Breedte, Verschaalde_Hoogte, marge)
      '---
      ThumbPage(Laatste).Cls
      ThumbPage(Laatste).Width = Verschaalde_Breedte + marge
      ThumbPage(Laatste).Height = Verschaalde_Hoogte + marge
      ThumbPage(Laatste).Top = CInt((ThumbOuterFrame.Height / 2) - (ThumbPage(Laatste).Height / 2))
      ThumbPage(Laatste).Left = CInt((ThumbOuterFrame.Width / 2) - (ThumbPage(Laatste).Width / 2))
      '---
      If m_Language = EnumLang.Dutch Then LbInfo.Caption = " Bezig met maken html Master pagina, afbeelding: " & CStr(I + 1) & " van: " & CStr(Aantal)
      If m_Language = EnumLang.English Then LbInfo.Caption = " Creating html Master page, processing Picture: " & CStr(I + 1) & " from: " & CStr(Aantal)
      '---
      ThumbPage(Laatste).PaintPicture P_load.Picture, marge \ 2, marge \ 2, Verschaalde_Breedte, Verschaalde_Hoogte
      FnmNew = LCase(LegalFilename(Fnm)) ' geen gedonder met spaties en hoofdletters
      '---
      If m_CompressToJPG = True Then ' JPG COMPRESSIE Thumbs en Origineel
         '---
         JpgQuality = 90:             '== Comprimeer de ThumbPage
         ThumbFnm = "tn_" & Mid(FnmNew, 1, Len(FnmNew) - 4) & ".jpg"
         Thumb_PadEnFile = m_Def_DestinationPath & "\" & ThumbFnm
         Call SavePicture(ThumbPage(Laatste).Image, App.Path & "\Tmp.bmp")
         Call BmpToJpeg(App.Path & "\Tmp.bmp", Thumb_PadEnFile, JpgQuality)
         '== Comprimeer het origineel indien nodig
         If m_Compress_NO_GIF = True Then
            If Right$(FnmNew, 4) = ".gif" Then
               Exitstentie = ".gif": Compress = False
            Else
               Exitstentie = ".jpg": Compress = True
            End If
         Else
            Exitstentie = ".jpg": Compress = True
         End If
         FnmNew = Mid(FnmNew, 1, Len(FnmNew) - 4) & Exitstentie
         Original_PadEnFile = m_Def_DestinationPath & "\OriginalPics\" & FnmNew
         If Compress = True Then
            If m_Html_ThumbnailBorderThicknes > 0 Then
               'Ook een foto border om de kopie van de originelen
               'Kan alleen bij jpg compressiemodus
               W_Border = (marge * (P_load.Width \ Verschaalde_Breedte))
               H_Border = (marge * (P_load.Height \ Verschaalde_Hoogte))
               ThumbPage(FotometBorder).Cls
               ThumbPage(FotometBorder).BackColor = m_Html_ThumbnailBorderColor
               ThumbPage(FotometBorder).Width = P_load.Width + W_Border
               ThumbPage(FotometBorder).Height = P_load.Height + H_Border
               ThumbPage(FotometBorder).PaintPicture P_load.Picture, (W_Border \ 2), (H_Border \ 2), P_load.Width, P_load.Height
               Call SavePicture(ThumbPage(FotometBorder).Image, App.Path & "\Tmp.bmp")
            Else
               Call SavePicture(P_load.Image, App.Path & "\Tmp.bmp")
            End If
            If m_Language = EnumLang.Dutch Then LbInfo.Caption = " Bezig met comprimeren naar jpg formaat, afbeelding: " & CStr(I + 1) & " van: " & CStr(Aantal)
            If m_Language = EnumLang.English Then LbInfo.Caption = " Compressing picture to jpg format, Picture: " & CStr(I + 1) & " from: " & CStr(Aantal)
            Call BmpToJpeg(App.Path & "\Tmp.bmp", Original_PadEnFile, JpgQuality)
            Kill App.Path & "\Tmp.bmp"
         Else
            Call FileCopy(Bron_PadEnFile, m_Def_DestinationPath & "\OriginalPics\" & FnmNew)
         End If
         '---
      ElseIf JPGCompressie = False Then
         '---
         ThumbFnm = "tn_" & Mid(FnmNew, 1, Len(FnmNew) - 4) & ".bmp"
         Thumb_PadEnFile = m_Def_DestinationPath & "\" & ThumbFnm
         Call SavePicture(ThumbPage(Laatste).Image, Thumb_PadEnFile)
         Call FileCopy(Bron_PadEnFile, m_Def_DestinationPath & "\OriginalPics\" & FnmNew)
         '---
      End If
      '---
      fileinfo = Make_File_info_string(m_Def_DestinationPath & "\OriginalPics\", FnmNew)
      Url = "OriginalPics\" & Mid(FnmNew, 1, Len(FnmNew) - 4) & ".htm"
      '---
      HtmlThumbFilenames(I, 1) = m_Def_DestinationPath & "\OriginalPics\" & Mid(FnmNew, 1, Len(FnmNew) - 4) & ".htm"
      HtmlThumbFilenames(I, 2) = Mid(FnmNew, 1, Len(FnmNew) - 4) & ".htm"
      HtmlThumbFilenames(I, 3) = FnmNew
      HtmlThumbFilenames(I, 4) = fileinfo
      HtmlThumbFilenames(I, 5) = P_load.ScaleWidth
      HtmlThumbFilenames(I, 6) = P_load.ScaleHeight
      '---
      Call Master_Print_PictureCel(ThumbFnm, Url, fileinfo, Verschaalde_Hoogte, Verschaalde_Breedte)
skip:
      Call Make_Gradient(Gradient, I, Aantal - 1, m_FloodFillColor, LbInfo.BackColor)
   Next
   
   Call Master_Closepage
   '---
   For j = 0 To Aantal - 1
       DoEvents
       If m_Language = EnumLang.Dutch Then LbInfo.Caption = " Bezig met maken html Detail pagina`s, afbeelding: " & CStr(j + 1) & " van: " & CStr(Aantal)
       If m_Language = EnumLang.English Then LbInfo.Caption = " Creating html Detail pages, processing Picture: " & CStr(j + 1) & " from: " & CStr(Aantal)
       '---
       If j > 0 And Aantal - 1 Then
          PrevLink = HtmlThumbFilenames(j - 1, 2)
       End If
       If j < Aantal - 1 Then
          NextLink = HtmlThumbFilenames(j + 1, 2)
       ElseIf j = Aantal - 1 Then
          NextLink = ""
       End If
       Picturename = HtmlThumbFilenames(j, 3)
       CreateFilename = HtmlThumbFilenames(j, 1)
       PictureInfo = HtmlThumbFilenames(j, 4)
       m_Width = CInt(HtmlThumbFilenames(j, 5))
       m_Height = CInt(HtmlThumbFilenames(j, 6))
       Call create_Detailpage(PictureInfo, m_Width, m_Height, CreateFilename, Picturename, PrevLink, NextLink)
   Next j
   '---
   RaiseEvent HtmLGaleryCreated(True, CLng(Aantal), CStr(TimeUsed))
   '---
   If m_Language = EnumLang.Dutch Then LbInfo.Caption = " Klaar !!"
   If m_Language = EnumLang.English Then LbInfo.Caption = " All done !!"
   Bezig_met_Laden = False
   m_ResetLoading = False
   ReDim HtmlThumbFilenames(1)
   ThumbPage(Laatste).Visible = False
   Screen.MousePointer = vbDefault
   On Error Resume Next
   Unload ThumbPage(Laatste)
   Unload ThumbPage(FotometBorder)
  
   '-----------------------------
   
   Filename = m_Def_DestinationPath & "\index.htm"
   RunStyle = 1
   Call RunFile(Filename, m_Def_DestinationPath, 1)
  
   HtmlPreviewmode = False
  
   Gradient.Visible = False
  
  
End Sub

Private Sub Master_Openpage(pad)
    OpenFnr = FreeFile
    LastFoldername = Trim(Dir_GetLastSubdir(CStr(pad)))
    Titel = LastFoldername
    Filename = pad & "\index.htm"
    Open Filename For Output As #OpenFnr
    Print #OpenFnr, "<Html>"
    Print #OpenFnr, "<head>"
    Print #OpenFnr, "<title>" & Titel & "</title>"
    If Customstyle = "" Then Customstyle = Create_Default_Css
    Print #OpenFnr, Customstyle
    Print #OpenFnr, "</head>"
    Print #OpenFnr, "<body>"
    Print #OpenFnr, "<p align=" & Chr(34) & "center" & Chr(34) & "><font size=" & Chr(34) & "4" & Chr(34) & ">" & Titel & "</font></p>"
    If Html_CustomHeaderlinks <> "" Then
       Print #OpenFnr, "<font size=1>" & Html_CustomHeaderlinks & "</font>"
    End If
    Print #OpenFnr, "<hr>"
    Print #OpenFnr, " <div align=" & Chr(34) & "left" & Chr(34) & ">"
    If Html_BodyText <> "" Then
       TableWidth = 600
    Else
       TableWidth = 420
    End If
    Print #OpenFnr, "  <table border=" & Chr(34) & "1" & Chr(34) & " width=" & Chr(34) & TableWidth & Chr(34) & " cellpadding=" & Chr(34) & "11" & Chr(34) & " cellspacing=" & Chr(34) & "8" & Chr(34) & " bordercolorlight=" & Chr(34) & "#000000" & Chr(34) & " bordercolordark=" & Chr(34) & "#808080" & Chr(34) & " bordercolor=" & Chr(34) & "#000000" & Chr(34) & ">"
    TextPlaced = False
End Sub
Private Sub Master_Print_PictureCel(F_Name, Url, Txt, Hgt, Wht)
    Print #OpenFnr, "    <tr>"
    Print #OpenFnr, "      <td width=" & Chr(34) & "200" & Chr(34) & " bgcolor=" & Chr(34) & "#575757" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & "><a href=" & Chr(34) & Url & Chr(34) & "><img border=" & Chr(34) & "0" & Chr(34) & " src=" & Chr(34) & F_Name & Chr(34) & " width=" & Chr(34) & Wht & Chr(34) & " height=" & Chr(34) & Hgt & Chr(34) & "></a></td>"
    Print #OpenFnr, "        <td width=" & Chr(34) & "140" & Chr(34) & " bgcolor=" & Chr(34) & "#575757" & Chr(34) & " align=" & Chr(34) & "left" & Chr(34) & ">"
    Print #OpenFnr, "          <dl>"
    Print #OpenFnr, "            <dt><font size=" & Chr(34) & "1" & Chr(34) & ">" & Txt & "</font></dt>"
    Print #OpenFnr, "        </dl>"
    Print #OpenFnr, "      </td>"
    If Html_BodyText <> "" And TextPlaced = False Then Call Master_BodyText
    Print #OpenFnr, "    </tr>"
End Sub
Private Sub Master_BodyText()
    Print #OpenFnr, "      <td width=" & Chr(34) & "160" & Chr(34) & " bgcolor=" & Chr(34) & "#353535" & Chr(34) & " align=" & Chr(34) & "left" & Chr(34) & " rowspan=" & Chr(34) & "1000" & Chr(34) & " valign=" & Chr(34) & "top" & Chr(34) & ">"
    Print #OpenFnr, "        <dl>"
    Print #OpenFnr, "          <dt><font size=" & Chr(34) & "1" & Chr(34) & ">" & Html_BodyText & "</font></dt>"
    Print #OpenFnr, "        </dl>"
    Print #OpenFnr, "      </td>"
    TextPlaced = True
End Sub
Private Sub Master_Closepage()
    Print #OpenFnr, "  </Table>"
    Print #OpenFnr, " </div>"
    Print #OpenFnr, " <hr size=" & Chr(34) & "1" & Chr(34) & " color=" & Chr(34) & "#990033" & Chr(34) & ">"
    Print #OpenFnr, "</Body>"
    Print #OpenFnr, "</Html>"
    Close #OpenFnr
End Sub
Private Sub create_Detailpage(PictureInfo, m_Width, m_Height, CreateFilename, Picturename, PrevLink, NextLink)
    Fnr = FreeFile
    Titel = Picturename
    Open CreateFilename For Output As #Fnr
    Print #Fnr, "<Html>"
    Print #Fnr, "<head>"
    Print #Fnr, "<title>" & Titel & "</title>"
    If Customstyle = "" Then Customstyle = Create_Default_Css
    Print #Fnr, Customstyle
    Print #Fnr, "</head>"
    Print #Fnr, "<body>"
    Print #Fnr, "<p align=" & Chr(34) & "center" & Chr(34) & "><font size=" & Chr(34) & "4" & Chr(34) & ">" & Titel & "</font></p>"
    IndexLink = "../index.htm"
    If m_Language = EnumLang.Dutch Then
        If PrevLink <> "" Then
           st = "<a href=" & PrevLink & ">Vorige</a> | "
        Else
           st = "Vorige | "
        End If
        st = st & "<a href=" & IndexLink & ">Inhoud</a> | "
        If NextLink <> "" Then
           st = st & "<a href=" & NextLink & ">Volgende</a>  "
        Else
           st = st & "Volgende"
        End If
    ElseIf m_Language = EnumLang.English Then
        If PrevLink <> "" Then
           st = "<a href=" & PrevLink & ">Previous</a> | "
        Else
           st = "Previous | "
        End If
        st = st & "<a href=" & IndexLink & ">Index</a> | "
        If NextLink <> "" Then
           st = st & "<a href=" & NextLink & ">Next</a>"
        Else
           st = st & "Next"
        End If
    End If
    Print #Fnr, "<font size=1 >" & st & "</font>"
    Print #Fnr, "<hr>"
    Print #Fnr, "<p> "
    Print #Fnr, "</p> "
    Print #Fnr, "  <Table border=" & Chr(34) & "1" & Chr(34) & "align=" & Chr(34) & "center" & Chr(34) & _
                " width=" & Chr(34) & TableWidth & Chr(34) & " cellpadding=" & Chr(34) & "11" & Chr(34) & _
                " cellspacing=" & Chr(34) & "8" & Chr(34) & _
                " bordercolorlight=" & Chr(34) & "#000000" & Chr(34) & _
                " bordercolordark=" & Chr(34) & "#808080" & Chr(34) & _
                " bordercolor=" & Chr(34) & "#000000" & Chr(34) & ">"
    Print #Fnr, "    <tr>"
    Print #Fnr, "      <td width=" & Chr(34) & "2" & Chr(34) & _
                         " bgcolor=" & Chr(34) & "#575757" & Chr(34) & _
                         " align=" & Chr(34) & "center" & Chr(34) & _
                         "><img border=" & Chr(34) & "0" & Chr(34) & _
                         " src=" & Chr(34) & Picturename & Chr(34) & _
                         " width=" & Chr(34) & m_Width & Chr(34) & _
                         " height=" & Chr(34) & m_Height & Chr(34) & "></a></td>"
    Print #Fnr, "    </tr>"
    Print #Fnr, "  </Table>"
    Print #Fnr, "<p> "
    Print #Fnr, "</p> "
    Print #Fnr, "<p> "
    Print #Fnr, "</p> "

    Print #Fnr, "  <Table border=" & Chr(34) & "1" & Chr(34) & "align=" & Chr(34) & "center" & Chr(34) & _
                " width=" & Chr(34) & TableWidth & Chr(34) & " cellpadding=" & Chr(34) & "11" & Chr(34) & _
                " cellspacing=" & Chr(34) & "8" & Chr(34) & _
                " bordercolorlight=" & Chr(34) & "#000000" & Chr(34) & _
                " bordercolordark=" & Chr(34) & "#808080" & Chr(34) & _
                " bordercolor=" & Chr(34) & "#000000" & Chr(34) & ">"
    Print #Fnr, "    <tr>"
    Print #Fnr, "      <td width=" & Chr(34) & "160" & Chr(34) & " bgcolor=" & Chr(34) & "#353535" & Chr(34) & " align=" & Chr(34) & "left" & Chr(34) & " rowspan=" & Chr(34) & "1000" & Chr(34) & " valign=" & Chr(34) & "top" & Chr(34) & ">"
    Print #Fnr, "        <dl>"
    Print #Fnr, "          <dt><font size=" & Chr(34) & "1" & Chr(34) & ">" & PictureInfo & "</font></dt>"
    Print #Fnr, "        </dl>"
    Print #Fnr, "      </td>"
    Print #Fnr, "    </tr>"
    Print #Fnr, "  </Table>"
    Print #Fnr, "</Body>"
    Print #Fnr, "</Html>"
    Close #Fnr
End Sub
Private Function Create_Default_Css()
      st = "<style>"
 st = st & "  BODY    { FONT-SIZE: 8pt; BACKGROUND-IMAGE: bg.jpg); COLOR: #999999;    FONT-FAMILY: Verdana; BACKGROUND-COLOR: #353535}" & vbCrLf
 st = st & "  A:link  { FONT-SIZE: 8pt; COLOR: #C0C0C0;}" & vbCrLf
 st = st & "  A:visit { FONT-SIZE: 8pt; COLOR: #C0C0C0;}" & vbCrLf
 st = st & "  A:hover { FONT-SIZE: 8pt; BORDER-TOP: #990033 1px solid; COLOR: #FFFFFF; BORDER-BOTTOM: #990033 1px solid; BACKGROUND-COLOR: #990033; text-decoration: none;}" & vbCrLf
 st = st & "  HR      { height:1 ; color:#990033}" & vbCrLf
 st = st & "</style>"
  Create_Default_Css = st
End Function


' ########################## P R I N T E N ########################

Public Sub PrintPicture()
    Can_Print = True
    Call PrintPreview
End Sub
Public Sub PrintPreview()
    
    Call ResetLoading
    
    PrintPreviewMode = True
   
    Gradient.Visible = True
    
    Dim Aantal, R As Integer, c As Integer, I As Integer, Bladzijde As Integer, Klr(30) As Long
    '---
    Bezig_met_Laden = True
    '---
    Pagina_Seperator = 20
    '---
    Printer.PrintQuality = m_PrintResolution
    Printer.ScaleMode = vbCentimeters
    '---
    If Pagina_Hoogte = 0 Then Pagina_Hoogte = ThumbOuterFrame.Height - (Pagina_Seperator * 2)
    Pagina_Breedte = CLng((Printer.ScaleWidth / Printer.ScaleHeight) * Pagina_Hoogte)
    '---
    Call Unload_ThumbInnerframes
    
    ThumbInnerframe(0).Height = ThumbOuterFrame.Height - 1: V_Scroll.Enabled = False
    '---
    ThumbPage(0).Visible = True
    ThumbPage(0).Width = CLng(Pagina_Breedte * (m_Print_Zoomfactor / 100))
    ThumbPage(0).Height = CLng(Pagina_Hoogte * (m_Print_Zoomfactor / 100))
    ThumbPage(0).Top = Pagina_Seperator
    ThumbPage(0).Left = (ThumbOuterFrame.Width / 2) - (ThumbPage(0).Width / 2)
    ThumbPage(0).BackColor = QBColor(15)
    '---
    Print_Thumb_Size = CInt((ThumbPage(0).Width / m_Print_NrThumbCols)) - 11
    '---
    'Thumbs Per Rij wordt bepaald door de breedte van de control /de afmeting ng van de thumb en de groote
    Print_Thumbs_PerRij = ThumbPage(0).Width \ Print_Thumb_Size
    Print_Thumbs_PerKolom = ThumbPage(0).Height \ Print_Thumb_Size
    
    X_Marge = CInt((ThumbPage(0).Width / 2) - ((Print_Thumbs_PerRij * Print_Thumb_Size) / 2))
    Y_Marge = CInt((ThumbPage(0).Height / 2) - ((Print_Thumbs_PerKolom * Print_Thumb_Size) / 2))
        
    ThumbPage(0).Cls
    ThumbPage(0).AutoRedraw = True
    '================[ Padinstellingen
    If Ambient.UserMode = True Then 'gebruikers niveau
       pad = File1.Path
       If Right$(pad, 1) <> "\" Then pad = pad & "\"
       Aantal = File1.ListCount
       ThumbsPerPage = Print_Thumbs_PerKolom * Print_Thumbs_PerRij
    Else
       Aantal = Print_Thumbs_PerKolom * Print_Thumbs_PerRij
       ThumbsPerPage = Print_Thumbs_PerKolom * Print_Thumbs_PerRij
    End If
   '================[ Thumbs laden
    
    start_time = Timer
    
    For I = 0 To Aantal - 1
        DoEvents
        If m_ResetLoading = True Then Exit For
        If Ambient.UserMode = True Then 'gebruikers niveau
           Fnm = File1.List(I)
           Bron_PadEnFile = pad & Fnm
           If Fnm = "" Then GoTo SubEinde
           P_load.Picture = LoadPicture(Bron_PadEnFile)
        Else
           P_load.Picture = PreviewPicture.Picture
        End If
           
        If m_Language = EnumLang.Dutch Then LbInfo.Caption = " Pagina " & Bladzijde + 1 & " Afbeelding: " & CStr(I + 1) & " van: " & CStr(Aantal)
        If m_Language = EnumLang.English Then LbInfo.Caption = " Page " & Bladzijde + 1 & " picture: " & CStr(I + 1) & " from: " & CStr(Aantal)
                  
        If I >= (ThumbsPerPage) * (Bladzijde + 1) Then
           Bladzijde = Bladzijde + 1
           '-------------------- Pagina_Seperator
           Tp = Pagina_Seperator + (ThumbPage(0).Height * Bladzijde) + (Pagina_Seperator * Bladzijde)
           lft = ThumbPage(0).Left
           Wid = ThumbPage(0).Width
           Hgt = ThumbPage(0).Height
           
            
           Call Load_ThumbPage(Bladzijde, lft, Tp, Wid, Hgt)
           
           If Can_Print = True Then Printer.NewPage
           '-------------------- verdelen over pagina
           kolom = 0
           Rij = 0
           '-------------------- Frame vergroten
           Hgt = ThumbPage(Bladzijde).Top + ThumbPage(Bladzijde).Height + Pagina_Seperator
           '---
           If Hgt > ThumbOuterFrame.Height Then
              ThumbInnerframe(0).Height = Hgt
              V_Scroll.Enabled = True: V_Scroll.Max = (ThumbInnerframe(0).Height - ThumbOuterFrame.Height) - BinnenMarge
              Top_Oud = ThumbPage(Bladzijde - 1).Top
              Top_Nieuw = ThumbPage(Bladzijde).Top
              For Tp = Top_Oud To Top_Nieuw Step 2
                  ThumbInnerframe(0).Top = (0 - Tp) + BinnenMarge
                  If Tp <= V_Scroll.Max Then V_Scroll.Value = Tp
              Next Tp
           Else
              ThumbInnerframe(0).Height = ThumbOuterFrame.Height - 1
              V_Scroll.Enabled = False: V_Scroll.Max = 0
           End If
        End If
        
        '=================================================================
        '   Bereken de juiste verhouding van de afbeelding binnen het vierkant van de ThumbPage
        Call Verschaal_Picture(Print_Thumb_Size, Print_Thumb_Size, Verschaalde_Breedte, Verschaalde_Hoogte, 10)
        
        X1 = (kolom * Print_Thumb_Size) + X_Marge + ((Print_Thumb_Size - Verschaalde_Breedte) / 2)
        Y1 = ((Rij * Print_Thumb_Size) + Y_Marge) + ((Print_Thumb_Size - Verschaalde_Hoogte) / 2)
        
        '=================================================================
        
        If Verschaalde_Breedte And Verschaalde_Hoogte > 0 Then
          ThumbPage(Bladzijde).PaintPicture P_load.Picture, X1, Y1, Verschaalde_Breedte, Verschaalde_Hoogte
          
          Extender.Parent.Caption = "Can_Print = " & Can_Print
          If Can_Print = True Then
             Printer.PaintPicture P_load.Picture, X1, Y1, Verschaalde_Breedte, Verschaalde_Hoogte
          End If
        End If
        
        '-----
        kolom = kolom + 1: If kolom > Print_Thumbs_PerRij - 1 Then Rij = Rij + 1: kolom = 0
        '-----
        Call Make_Gradient(Gradient, I, Aantal - 1, m_FloodFillColor, LbInfo.BackColor)
     Next
     '===========================================================
     
     If Can_Print = True Then Printer.EndDoc
     
     
     If m_ResetLoading = True Then GoTo SubEinde
     ' Call SavePicture(Pagina(Bladzijde).Image, App.Path & "\" & Bladzijde & "_Prnt.bmp")
     If V_Scroll.Max > 0 Then
        V_Scroll.Value = 0
        ThumbInnerframe(0).Top = 0
     End If
SubEinde:
    TimeUsed = Format$(Timer - start_time, "00.00") & " sec"
    If m_Language = EnumLang.Dutch Then LbInfo.Caption = " Pagina: " & Bladzijde + 1 & " Afbeelding: " & CStr(I + 1) & " van: " & CStr(Aantal) & "  Tijd:" & TimeUsed
    If m_Language = EnumLang.English Then LbInfo.Caption = " Page: " & Bladzijde + 1 & " Picture: " & CStr(I + 1) & " from: " & CStr(Aantal) & "  Tijd:" & TimeUsed
    
    Bezig_met_Laden = False
    m_ResetLoading = False
    Can_Print = False
    
    Gradient.Visible = False

Exit Sub
fout:
Select Case Err
Case 480 ' can`t AutoRedraw
     Resume
Case Else
     MsgBox Error(Err)
     Bezig_met_Laden = False
     Can_Print = False
     m_ResetLoading = False
End Select
Gradient.Visible = False

End Sub




' ############### H E L P E R - R O U T I N E S ###############



Private Sub Verschaal_Picture(ThumbnailBuitenBreedte, ThumbnailBuitenHoogte, Verschaalde_Breedte, Verschaalde_Hoogte, BinnenMarge)
'Thumb_Height
  Verschaalde_Breedte = P_load.Width
  Verschaalde_Hoogte = P_load.Height
  If Verschaalde_Breedte > (ThumbnailBuitenBreedte - BinnenMarge) Then
     Verschaalde_Hoogte = CInt(Verschaalde_Hoogte * (ThumbnailBuitenBreedte - BinnenMarge) / Verschaalde_Breedte)
     Verschaalde_Breedte = CInt((ThumbnailBuitenBreedte - BinnenMarge))
  End If
  If Verschaalde_Hoogte > (ThumbnailBuitenHoogte - BinnenMarge) Then
     Verschaalde_Breedte = CInt(Verschaalde_Breedte * (ThumbnailBuitenHoogte - BinnenMarge) / Verschaalde_Hoogte)
     Verschaalde_Hoogte = CInt((ThumbnailBuitenHoogte - BinnenMarge))
  End If
End Sub


Private Sub Load_ThumbPage(Index, X, y, Breedte, Hoogte)
 On Error Resume Next   'voor het geval deze index al gelanden is
 Load ThumbPage(Index)
 ThumbPage(Index).Left = X
 ThumbPage(Index).Top = y
 ThumbPage(Index).Width = Breedte
 ThumbPage(Index).Height = Hoogte
 ThumbPage(Index).Visible = True
 ThumbPagesLoaded = ThumbPagesLoaded + 1
End Sub
Private Sub Unload_ThumbPages()
 On Local Error GoTo exit_Sub
 For I = 1 To ThumbPagesLoaded
     Unload ThumbPage(I)
     ThumbPagesLoaded = I
 Next
 Exit Sub
exit_Sub:
End Sub

Private Sub Load_ThumbInnerframe(Index, X, y, Breedte, Hoogte)
 On Error GoTo exit_Sub
 Load ThumbInnerframe(Index)
 ThumbInnerframe(Index).Left = X
 ThumbInnerframe(Index).Top = y
 ThumbInnerframe(Index).Width = Breedte
 ThumbInnerframe(Index).Height = Hoogte
 ThumbInnerframe(Index).Visible = True
 '============================
 ThumbFramesLoaded = ThumbFramesLoaded + 1
Exit Sub
exit_Sub:
End Sub

Private Sub Unload_ThumbInnerframes()
 On Local Error GoTo exit_Sub
 For I = 1 To ThumbFramesLoaded
     Unload ThumbInnerframe(I)
     ThumbFramesLoaded = I
 Next
Exit Sub
exit_Sub:
End Sub
Public Sub ResetLoading()
  PreviousPage = 0
  CurrentThumbnail = 0
  PreviousThumbnail = 0
 
  If Bezig_met_Laden = True Then m_ResetLoading = True ' een voorgaande lus eerst afbreken
  Call Unload_ThumbPages
  Call Unload_ThumbInnerframes
  Bezig_met_Laden = True
  m_ResetLoading = False
End Sub


Private Function Make_File_info_string(PathOnly, FilenameOnly)
  On Local Error GoTo fout:
  Dim t As String
  If Right$(PathOnly, 1) <> "\" Then PathOnly = PathOnly & "\"
  Fname = PathOnly & FilenameOnly
  Select Case m_Language
  Case EnumLang.Dutch
       t = t & "<dt>" & "Bestand  : " & FilenameOnly & "</dt>" & vbCrLf
       t = t & "<dt>" & "Groote   : " & Format(FileLen(Fname) \ 1024, "##,##0 Kb") & "</dt>" & vbCrLf
       t = t & "<dt>" & "Datum    : " & FileDateTime(Fname) & vbCrLf
       t = t & "<dt>" & "Breedte  : " & Format(P_load.ScaleWidth, "##,##") & " pixels" & "</dt>" & vbCrLf
       t = t & "<dt>" & "Hoogte   : " & Format(P_load.ScaleHeight, "##,##") & " pixels" & "</dt>" & vbCrLf
  Case Else
       t = t & "<dt>" & "Filename : " & FilenameOnly & "</dt>" & vbCrLf
       t = t & "<dt>" & "Size     : " & Format(FileLen(Fname) \ 1024, "##,##0 Kb") & "</dt>" & vbCrLf
       t = t & "<dt>" & "Date     : " & FileDateTime(Fname) & "</dt>" & vbCrLf
       t = t & "<dt>" & "Width    : " & Format(P_load.ScaleWidth, "##,##") & " pixels" & "</dt>" & vbCrLf
       t = t & "<dt>" & "Height   : " & Format(P_load.ScaleHeight, "##,##") & " pixels" & "</dt>" & vbCrLf
  End Select
  Make_File_info_string = t
 Exit Function
fout:
 Make_File_info_string = ""
End Function
Private Sub SetPresetcolors(Colorset)
Select Case Colorset
Case EnumPresetColor.Black
     m_ThumbPage_FrameColor = &H0&
     m_ThumbPageColor = &H404040
     m_ThumbNailColor = &H0&
     m_ThumbNailBorderColor = &H3A4228
     m_ThumbNailSelectColor = &HFFFF&
Case EnumPresetColor.Antraciet
     m_ThumbPage_FrameColor = 0
     m_ThumbPageColor = &H353535
     m_ThumbNailColor = &H575757
     m_ThumbNailBorderColor = &H999999
     m_ThumbNailSelectColor = &HFF00&
Case EnumPresetColor.Dark_Grey
     m_ThumbPage_FrameColor = &H404040
     m_ThumbPageColor = &H808080
     m_ThumbNailColor = &HC0C0C0
     m_ThumbNailBorderColor = &H404040
     m_ThumbNailSelectColor = &HFF00&
Case EnumPresetColor.Middle_Grey
     m_ThumbPage_FrameColor = &H515151
     m_ThumbPageColor = &HC0C0C0
     m_ThumbNailColor = &HE0E0E0
     m_ThumbNailBorderColor = &HFFFFFF
     m_ThumbNailSelectColor = &HFF&
Case EnumPresetColor.Light_Grey
     m_ThumbPage_FrameColor = &H515151
     m_ThumbPageColor = &HE0E0E0
     m_ThumbNailColor = &HFFFFFF
     m_ThumbNailBorderColor = &HC0C0C0
     m_ThumbNailSelectColor = &HFF&
Case EnumPresetColor.White_Grey
     m_ThumbPage_FrameColor = &H515151
     m_ThumbPageColor = &HFFFFFF
     m_ThumbNailColor = &HE0E0E0
     m_ThumbNailBorderColor = &HC0C0C0
     m_ThumbNailSelectColor = &HFF&
Case Else
End Select
If Colorset > 0 Then
   If Ambient.UserMode = False Then Makegallery
End If
End Sub


Public Sub Makegallery()
    
    On Error GoTo fout:
    
    Call ResetLoading
      
    PrintPreviewMode = False
    HtmlPreviewmode = False
    
    Dim Aantal, R As Integer, c As Integer, I As Integer, Bladzijde As Integer, Klr(30) As Long
    '---
    Bezig_met_Laden = True
    '---
    Pagina_Seperator = BinnenMarge * 2
    '---
    Pagina_Hoogte = ThumbOuterFrame.Height - (Pagina_Seperator * 2)
    Pagina_Breedte = ThumbOuterFrame.Width - (BinnenMarge * 2)
    '---
    ThumbOuterFrame.BackColor = m_ThumbPage_FrameColor
    ThumbInnerframe(0).BackColor = m_ThumbPage_FrameColor
    '---
    ThumbInnerframe(0).Visible = True
    ThumbInnerframe(0).Height = ThumbOuterFrame.Height - 1: V_Scroll.Enabled = False
    ThumbInnerframe(0).Width = Pagina_Breedte
    ThumbInnerframe(0).Top = 0
    '---
    ThumbPage(0).Visible = True
    ThumbPage(0).Width = Pagina_Breedte
    ThumbPage(0).Height = Pagina_Hoogte
    ThumbPage(0).Top = Pagina_Seperator
    ThumbPage(0).Left = (ThumbInnerframe(0).Width / 2) - (ThumbPage(0).Width / 2)
    ThumbPage(0).BackColor = m_ThumbPageColor
    
    '---
    marge = 10 ' De marge tussen het kader van, en de ThumbPage
    Tmarge = 5 ' De ruimte tussen de thumbnailkaders
    
    '=============================================================
    Thumb_Height = m_ThumbNailSize + marge + m_ThumbnailExtraHeight
    Thumb_Width = m_ThumbNailSize + marge + m_ThumbnailExtraWidth
    
    '---
    'Thumbs Per Rij wordt bepaald door de breedte van de control /de afmeting ng van de thumb en de groote
    Thumbs_PerRij = ThumbPage(0).Width \ (Thumb_Width + Tmarge)
    Thumbs_PerKolom = ThumbPage(0).Height \ (Thumb_Height + Tmarge)
    
    'ThumbPageStatus(m_ThumbNailSize,ThumbsPerPage,PagesNeeded
    
        
    X_Marge = CInt((ThumbPage(0).Width / 2) - ((Thumbs_PerRij * (Thumb_Width + Tmarge)) / 2)) ' - Tmarge
    Y_Marge = CInt((ThumbPage(0).Height / 2) - ((Thumbs_PerKolom * (Thumb_Height + Tmarge)) / 2)) ' - Tmarge
        
    ThumbPage(0).Cls
    ThumbPage(0).AutoRedraw = True
    '================[ Padinstellingen
    If Ambient.UserMode = True Then 'gebruikers niveau
       pad = File1.Path
       If Right$(pad, 1) <> "\" Then pad = pad & "\"
       Aantal = File1.ListCount
       ThumbsPerPage = Thumbs_PerKolom * Thumbs_PerRij
       '---------------------------------
       PagesNeeded = Aantal \ ThumbsPerPage
       '--------------------------------
       If PagesNeeded > m_MaxThumbnailpages Then
          PageResolution = CStr(ThumbPage(0).Width) & "x" & CStr(ThumbPage(0).Height)
          Call ToMuchPages(PagesNeeded, m_MaxThumbnailpages, ThumbsPerPage, m_ThumbNailSize, PageResolution)
          Bezig_met_Laden = False
          m_ResetLoading = False
          Can_Print = False
          Gradient.Visible = False
          Call ResetLoading
          Exit Sub
       End If
       RaiseEvent ThumbPageStatus(m_ThumbNailSize, ThumbsPerPage, PagesNeeded, Aantal)
    Else
       Aantal = Thumbs_PerKolom * Thumbs_PerRij
       ThumbsPerPage = Thumbs_PerKolom * Thumbs_PerRij
    
    End If
    ReDim Thumb_Index(Aantal, 7)
   '================[ Thumbs laden
    start_time = Timer
    
    Gradient.Visible = True
    
    For I = 0 To Aantal - 1
        DoEvents
        If m_ResetLoading = True Then Exit For
        If Ambient.UserMode = True Then 'gebruikers niveau
           Fnm = File1.List(I)
           Bron_PadEnFile = pad & Fnm
           If Fnm = "" Then GoTo SubEinde
           P_load.Picture = LoadPicture(Bron_PadEnFile)
        Else
           P_load.Picture = PreviewPicture.Picture
        End If
           
        If m_Language = EnumLang.Dutch Then LbInfo.Caption = " Pagina " & Bladzijde + 1 & " Afbeelding: " & CStr(I + 1) & " van: " & CStr(Aantal) & "  ThumbSize:" & m_ThumbNailSize & " x " & m_ThumbNailSize
        If m_Language = EnumLang.English Then LbInfo.Caption = " Page " & Bladzijde + 1 & " picture: " & CStr(I + 1) & " from: " & CStr(Aantal) & "  ThumbSize:" & m_ThumbNailSize & " x " & m_ThumbNailSize
                  
        If I >= (ThumbsPerPage) * (Bladzijde + 1) Then
           Bladzijde = Bladzijde + 1
           Tp = Pagina_Seperator + (ThumbPage(0).Height * Bladzijde) + (Pagina_Seperator * Bladzijde)
           kolom = 0: Rij = 0
           Call Load_ThumbPage(Bladzijde, ThumbPage(0).Left, Tp, ThumbPage(0).Width, ThumbPage(0).Height)
           
           
           Hgt = ThumbPage(Bladzijde).Top + ThumbPage(Bladzijde).Height + Pagina_Seperator
           If Hgt > ThumbOuterFrame.Height Then
              ThumbInnerframe(0).Height = Hgt
              V_Scroll.Enabled = True: V_Scroll.Max = (ThumbInnerframe(0).Height - ThumbOuterFrame.Height) - BinnenMarge
              Top_Oud = ThumbPage(Bladzijde - 1).Top
              Top_Nieuw = ThumbPage(Bladzijde).Top
              '-----------------------------------
              waarde = 0
              waarde = CInt(ThumbPage(Bladzijde).Width \ 100)
              waarde = waarde + waarde
              '-------
              If waarde <= 1 Then waarde = 2
              '-------
              For Tp = Top_Oud To Top_Nieuw Step waarde
                  ThumbInnerframe(0).Top = (0 - Tp) + BinnenMarge
                  If Tp <= V_Scroll.Max Then V_Scroll.Value = Tp
              Next Tp
           Else
              ThumbInnerframe(0).Height = ThumbOuterFrame.Height - 1
              V_Scroll.Enabled = False: V_Scroll.Max = 0
           End If
        
        End If
        '=================================================================
        Call Verschaal_Picture(Thumb_Width, Thumb_Height, Verschaalde_Breedte, Verschaalde_Hoogte, marge)
                
        If Verschaalde_Breedte And Verschaalde_Hoogte > 0 Then
           X1 = (kolom * (Thumb_Width + Tmarge)) + X_Marge
           Y1 = (Rij * (Thumb_Height + Tmarge)) + Y_Marge
           X2 = (X1 + Thumb_Width)
           Y2 = (Y1 + Thumb_Height)
           '---
           ThumbPage(Bladzijde).Line (X1 - 1, Y1 - 1)-(X2 + 1, Y2 + 1), m_ThumbNailBorderColor, BF
           ThumbPage(Bladzijde).Line (X1, Y1)-(X2, Y2), m_ThumbNailColor, BF
           '---
           Xx1 = (X1 + (Thumb_Width - Verschaalde_Breedte) / 2)
           Yy1 = (Y1 + (Thumb_Height - Verschaalde_Hoogte) / 2)
           Xx2 = (Xx1 - Thumb_Width)
           Yy2 = (Yy1 - Thumb_Height)
           '---
           ThumbPage(Bladzijde).PaintPicture P_load.Picture, Xx1, Yy1, Verschaalde_Breedte, Verschaalde_Hoogte
        
           Thumb_Index(I, ThumbID.X_1) = X1
           Thumb_Index(I, ThumbID.y_1) = Y1
           Thumb_Index(I, ThumbID.X_2) = X2
           Thumb_Index(I, ThumbID.y_2) = Y2
           Thumb_Index(I, ThumbID.File_name) = Bron_PadEnFile
           Thumb_Index(I, ThumbID.pagina) = Bladzijde
        
        End If
        kolom = kolom + 1: If kolom > Thumbs_PerRij - 1 Then Rij = Rij + 1: kolom = 0
        '=================================================================
        Call Make_Gradient(Gradient, I, Aantal - 1, m_FloodFillColor, LbInfo.BackColor)
        
     Next
     If V_Scroll.Max > 0 Then
        V_Scroll.Value = 0
        ThumbInnerframe(0).Top = 0
     End If
SubEinde:
    TimeUsed = Format$(Timer - start_time, "00.00") & " sec"
    If m_Language = EnumLang.Dutch Then LbInfo.Caption = " Pagina: " & Bladzijde & " Afbeelding: " & CStr(I + 1) & " van: " & CStr(Aantal) & "  Tijd:" & TimeUsed & "  ThumbSize:" & m_ThumbNailSize & " x " & m_ThumbNailSize
    If m_Language = EnumLang.English Then LbInfo.Caption = " Page: " & Bladzijde & " Picture: " & CStr(I + 1) & " from: " & CStr(Aantal) & "  Tijd:" & TimeUsed & "  ThumbSize:" & m_ThumbNailSize & " x " & m_ThumbNailSize
    
    Bezig_met_Laden = False
    m_ResetLoading = False
    Can_Print = False
    Gradient.Visible = False

Exit Sub
fout:
     ff = FreeFile
             Datum = Day(Now) & "-"
     Datum = Datum & Month(Now) & "-"
     Datum = Datum & Year(Now)
            Tijd = Hour(Now) & "-"
     Tijd = Tijd & Minute(Now) & "-"
     Tijd = Tijd & Second(Now)
     Datumtijd = Datum & "__" & Tijd & "_"
     
     '====================================
          
          st = "========================================" & vbCrLf
     st = st & "datum.......: " & Datum & vbCrLf
     st = st & "Tijd........: " & Tijd & vbCrLf
     st = st & "========================================" & vbCrLf
     st = st & "Foutmelding : " & Error(Err) & vbCrLf
     st = st & "========================================" & vbCrLf
     st = st & "Aantal afbeeldingen : " & Aantal & vbCrLf
     st = st & "Gestopt bij.........: " & I & " pagina....:" & Bladzijde & vbCrLf
     st = st & "ThumbNailSize.......: " & m_ThumbNailSize & vbCrLf
     st = st & "ThumbsPerPage ......:" & ThumbsPerPage & vbCrLf
     st = st & "PagesNeeded.........: " & PagesNeeded & vbCrLf
     st = st & "========================================" & vbCrLf
     st = st & "Paginagrootte : " & vbCrLf
     st = st & "      Breedte :" & ThumbPage(0).Width & vbCrLf
     st = st & "      Hoogte  :" & ThumbPage(0).Height & vbCrLf & vbCrLf & vbCrLf
     '====================================
     Open App.Path & "\Error_" & Datumtijd & ".txt" For Output As #ff
     ' ThumbPageStatus(m_ThumbNailSize, ThumbsPerPage, PagesNeeded, Aantal)
     Print #ff, st
     Close #ff
     st = st & "E X I T  -  T H U M B N A I L  -   F U N C T I E "
          
     MsgBox st
     
     Bezig_met_Laden = False
     Can_Print = False
     m_ResetLoading = False
     Gradient.Visible = False
     Call ResetLoading
     '====================================
End Sub

Private Sub Make_Gradient(Pb As PictureBox, Currentindex, Maxindex, Fore_Color, Back_Color)
   On Error Resume Next
   If Pb.BackColor <> Back_Color Then Pb.BackColor = Back_Color
   Prcnt = Currentindex / Maxindex
   m_CurrentIndex = (Pb.Width) * Prcnt
   Pb.Cls
   Pb.Line (0, 0)-(m_CurrentIndex, Pb.Height), Fore_Color, BF
End Sub

Public Function ThumbnailPagesNeeded()
    marge = 10 ' De marge tussen het kader van, en de ThumbPage
    Tmarge = 5 ' De ruimte tussen de thumbnailkaders
    Thumb_Size = m_ThumbNailSize + marge
    Thumbs_PerRij = ThumbOuterFrame.Width \ (Thumb_Size + Tmarge)
    Thumbs_PerKolom = ThumbOuterFrame.Height \ (Thumb_Size + Tmarge)
    ThumbsPerPage = Thumbs_PerKolom * Thumbs_PerRij
    Aantal = File1.ListCount
    ThumbnailPagesNeeded = Aantal \ ThumbsPerPage
End Function
Public Function ThumbnailsPerPage()
    marge = 10 ' De marge tussen het kader van, en de ThumbPage
    Tmarge = 5 ' De ruimte tussen de thumbnailkaders
    Thumb_Size = m_ThumbNailSize + marge
    Thumbs_PerRij = ThumbOuterFrame.Width \ (Thumb_Size + Tmarge)
    Thumbs_PerKolom = ThumbOuterFrame.Height \ (Thumb_Size + Tmarge)
    ThumbnailsPerPage = Thumbs_PerKolom * Thumbs_PerRij
End Function

Private Sub ToMuchPages(Pages, Maxpages, TmbsPerPage, CurrentThumbsize, PageResolution)
    maxfiles = Maxpages * TmbsPerPage
    marge = 10 ' De marge tussen het kader van, en de ThumbPage
    Tmarge = 5 ' De ruimte tussen de thumbnailkaders
    For I = 1 To CurrentThumbsize
        New_ThumbNailSize = CurrentThumbsize - I
    
        Thmb_Size = New_ThumbNailSize + marge
        Thmbs_PerRij = ThumbPage(0).Width \ (Thmb_Size + Tmarge)
        Thmbs_PerKolom = ThumbPage(0).Height \ (Thmb_Size + Tmarge)
        ThmbnailsPerPage = Thmbs_PerKolom * Thmbs_PerRij
        Aantal = File1.ListCount
        PagesNeeded = Aantal \ ThmbnailsPerPage
        If PagesNeeded <= Maxpages Then
           RecommendedThumbsize = New_ThumbNailSize
           Exit For
        End If
    Next
    If m_Language = EnumLang.Dutch Then
       st = st & "Er moeten te veel thumbnail pagina's gegenereerd worden!" & vbCrLf & vbCrLf
       st = st & "Berekend aantal pagina's = " & Pages & vbCrLf
       st = st & "Aanbevolen maximum = " & Maxpages & " Pagina's" & vbCrLf & vbCrLf
       st = st & "Probeer het volgende: " & vbCrLf & vbCrLf
       st = st & "     1 > Beperk het aantal afbeeldingen per map in deze thumbnail en paginagrootte en tot " & maxfiles & vbCrLf & vbCrLf
       st = st & "     2 > Of verklein de thumbnailgrootte van  " & CurrentThumbsize & " x " & CurrentThumbsize & "  naar  " & RecommendedThumbsize & " x " & RecommendedThumbsize & "  of kleiner" & vbCrLf & vbCrLf
       st = st & "     3 > Of vergroot het thumbnailwindow meer dan " & PageResolution & vbCrLf & vbCrLf
       st = st & "     4 > Een combinatie van [1] en [2] kan nog effectiever zijn  " & vbCrLf & vbCrLf
       
       MsgBox st, vbExclamation, "Te veel Thumbnailpagina's"
 
    ElseIf m_Language = EnumLang.English Then
       st = st & "There are to mutch pages to generate !" & vbCrLf & vbCrLf
       st = st & "Calculated pages = " & Pages & vbCrLf
       st = st & "Recommended as maximum = " & Maxpages & " Pages" & vbCrLf & vbCrLf
       st = st & "Tri to: " & vbCrLf & vbCrLf
       st = st & "     1 > Reduce the number of files in a directory used with this thumbnail and page resolution to " & maxfiles & vbCrLf & vbCrLf
       st = st & "     2 > Or decrease the current thumbsize from " & CurrentThumbsize & " x " & CurrentThumbsize & " to " & RecommendedThumbsize & " x " & RecommendedThumbsize & " or even smaller & vbCrLf" & vbCrLf
       st = st & "     3 > Or Ingrease the size of the thumbnailwindow more then " & PageResolution & vbCrLf & vbCrLf
       st = st & "     4 > A combination of [1] and [2] may be even more effectieve " & vbCrLf & vbCrLf
    
       MsgBox st, vbExclamation, "To mutch pages to generate"
    End If
   
End Sub

Public Property Get ThumbnailExtraWidth() As ThumbSize
    ThumbnailExtraWidth = m_ThumbnailExtraWidth
End Property
        Public Property Let ThumbnailExtraWidth(ByVal New_ThumbnailExtraWidth As ThumbSize)
            m_ThumbnailExtraWidth = New_ThumbnailExtraWidth
            PropertyChanged "ThumbnailExtraWidth"
            If Ambient.UserMode = False Then Call Makegallery
        End Property
Public Property Get ThumbnailExtraHeight() As ThumbSize
    ThumbnailExtraHeight = m_ThumbnailExtraHeight
End Property
        Public Property Let ThumbnailExtraHeight(ByVal New_ThumbnailExtraHeight As ThumbSize)
            m_ThumbnailExtraHeight = New_ThumbnailExtraHeight
            PropertyChanged "ThumbnailExtraHeight"
            If Ambient.UserMode = False Then Call Makegallery
        End Property

