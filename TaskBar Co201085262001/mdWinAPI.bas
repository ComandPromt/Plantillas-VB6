Attribute VB_Name = "mdWinAPI"
Option Explicit

'=========================================================================
' API
'=========================================================================

Public Const WM_NCACTIVATE = &H86
Public Const WM_ACTIVATE = &H6
Public Const WM_MDISETMENU = &H230
Public Const WM_SYSCOMMAND = &H112
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_ERASEBKGND = &H14
Public Const WM_PAINT = &HF
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CANCELMODE = &H1F
Public Const WM_TIMER = &H113
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201

Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2
Public Const WA_INACTIVE = 0

Public Const SC_MINIMIZE = &HF020&
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_RESTORE = &HF120&

Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_CAPTION = &HC00000 '  WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_HSCROLL = &H100000
Public Const WS_VSCROLL = &H200000
Public Const WS_TABSTOP = &H10000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_BORDER = &H800000
Public Const WS_SYSMENU = &H80000
Public Const WS_EX_DLGMODALFRAME = &H1
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_WINDOWEDGE = &H100

Public Const SPI_GETWORKAREA = 48

Public Const STRETCH_HALFTONE = 4
Public Const STRETCH_DELETESCANS = 3

Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8

Public Const GWL_WNDPROC = (-4)
Public Const GWL_STYLE = (-16)
Public Const GWL_USERDATA = (-21)
Public Const GWL_EXSTYLE = (-20)

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNORMAL = 1

Public Const SWP_FRAMECHANGED = &H20 '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOOWNERZORDER = &H200 '  Don't do owner Z ordering
Public Const SWP_DRAWFRAME = &H20
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Public Const DFC_BUTTON = 4
Public Const DFC_CAPTION = 1
Public Const DFC_MENU = 2
Public Const DFC_SCROLL = 3
Public Const DFCS_ADJUSTRECT = &H2000
Public Const DFCS_BUTTON3STATE = &H8
Public Const DFCS_BUTTONCHECK = &H0
Public Const DFCS_BUTTONPUSH = &H10
Public Const DFCS_BUTTONRADIO = &H4
Public Const DFCS_BUTTONRADIOIMAGE = &H1
Public Const DFCS_BUTTONRADIOMASK = &H2
Public Const DFCS_CAPTIONCLOSE = &H0
Public Const DFCS_CAPTIONHELP = &H4
Public Const DFCS_CAPTIONMAX = &H2
Public Const DFCS_CAPTIONMIN = &H1
Public Const DFCS_CAPTIONRESTORE = &H3
Public Const DFCS_CHECKED = &H400
Public Const DFCS_FLAT = &H4000
Public Const DFCS_INACTIVE = &H100
Public Const DFCS_MENUARROW = &H0
Public Const DFCS_MENUARROWRIGHT = &H4
Public Const DFCS_MENUBULLET = &H2
Public Const DFCS_MENUCHECK = &H1
Public Const DFCS_MONO = &H8000
Public Const DFCS_PUSHED = &H200
Public Const DFCS_SCROLLCOMBOBOX = &H5
Public Const DFCS_SCROLLDOWN = &H1
Public Const DFCS_SCROLLLEFT = &H2
Public Const DFCS_SCROLLRIGHT = &H3
Public Const DFCS_SCROLLSIZEGRIP = &H8
Public Const DFCS_SCROLLSIZEGRIPRIGHT = &H10
Public Const DFCS_SCROLLUP = &H0

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKEN = &HA

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const RGN_DIFF = 4

Public Const ETO_CLIPPED = 4
Public Const ETO_OPAQUE = 2

Public Const TA_CENTER = 6
Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2

Public Const LOGPIXELSY = 90 '  Logical pixels/inch in Y
Public Const OPAQUE = 2
Public Const TRANSPARENT = 1

Public Const FW_BOLD = 700
Public Const FW_NORMAL = 400

Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8

Public Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function IntersectClipRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, ByVal lpDx As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function GetDialogBaseUnits Lib "user32" () As Long
Public Declare Function PathCompactPath Lib "shlwapi.dll" Alias "PathCompactPathA" (ByVal hDC As Long, ByVal Path As String, ByVal dx As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function PlayMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hMF As Long) As Long
Public Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hEmf As Long, lpRect As RECT) As Long
Public Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As Long
Public Declare Function CreateEnhMetaFileLong Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, ByVal lpRect As Long, ByVal lpDescription As String) As Long
Public Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Any) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type PAINTSTRUCT
        hDC As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved(32) As Byte
End Type
Public Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type
Public Const LF_FACESIZE = 32
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

'=========================================================================
' Functions
'=========================================================================

Public Sub SetWindowStyle(hWnd As Long, Style As Long, OnOff As Boolean)
    Dim lStyle As Long
    lStyle = GetWindowLong(hWnd, GWL_STYLE)
    If OnOff Then
        lStyle = lStyle Or Style
    Else
        lStyle = lStyle And (-1 - Style)
    End If
    SetWindowLong hWnd, GWL_STYLE, lStyle
End Sub

Public Sub SetWindowExStyle(hWnd As Long, ExStyle As Long, OnOff As Boolean)
    Dim lStyle As Long
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    If OnOff Then
        lStyle = lStyle Or ExStyle
    Else
        lStyle = lStyle And (-1 - ExStyle)
    End If
    SetWindowLong hWnd, GWL_EXSTYLE, lStyle
End Sub

Public Function TranslateColor(cOle As OLE_COLOR) As Long
    If CLng(cOle) < 0 Then
        TranslateColor = GetSysColor(Loword(CLng(cOle)))
    Else
        TranslateColor = CLng(cOle)
    End If
End Function

Public Function Hiword(l As Long) As Long
    Hiword = l / &H10000
End Function

Public Function Loword(l As Long) As Long
    Loword = l And &HFFFF&
    If Loword > &H7FFF Then
        Loword = Loword - &H10000
    End If
End Function

Function PtInRect(rc As RECT, pt As POINTAPI) As Boolean
    PtInRect = rc.Left <= pt.X And pt.X <= rc.Right And rc.Top <= pt.Y And pt.Y <= rc.Bottom
End Function

Public Function RefreshDC(ByVal hWnd As Long)
    Dim hDC         As Long
    Dim rcClient    As RECT
    
    On Error Resume Next
    If hWnd <> 0 Then
        hDC = GetDC(hWnd)
        GetClientRect hWnd, rcClient
        InvalidateRect hWnd, rcClient, 0
        ReleaseDC hWnd, hDC
    End If
End Function

Public Function CreateLogFont(ByVal hDC As Long, oFont As StdFont) As Long
    Dim lf          As LOGFONT
    
    On Error Resume Next
    lf.lfHeight = -CLng(oFont.Size * GetDeviceCaps(hDC, LOGPIXELSY) / 72)
    lf.lfWidth = 0
    lf.lfEscapement = 0
    lf.lfOrientation = 10
    lf.lfWeight = IIf(oFont.Bold, FW_BOLD, FW_NORMAL)
    lf.lfItalic = -oFont.Italic
    lf.lfUnderline = -oFont.Underline
    lf.lfStrikeOut = -oFont.Strikethrough
    lf.lfCharSet = oFont.Charset ' ANSI_CHARSET
    lf.lfOutPrecision = 0 ' OUT_DEFAULT_PRECIS
    lf.lfClipPrecision = 0 ' CLIP_DEFAULT_PRECIS
    lf.lfQuality = 0 ' DEFAULT_QUALITY
    lf.lfPitchAndFamily = 0 + 0 ' DEFAULT_PITCH | FF_DONTCARE
    CopyMemory lf.lfFaceName(1), ByVal (oFont.Name), Len(oFont.Name) + 1
    CreateLogFont = CreateFontIndirect(lf)
End Function

Public Function CalcHeight(oFont As StdFont) As Long
    Dim hDC             As Long
    Dim hFont           As Long
    Dim hPrevFont       As Long
    Dim tm              As TEXTMETRIC
    
    On Error Resume Next
    hDC = GetDC(0)
    hFont = CreateLogFont(hDC, oFont)
    hPrevFont = SelectObject(hDC, hFont)
    GetTextMetrics hDC, tm
    DeleteObject hFont
    SelectObject hDC, hPrevFont
    ReleaseDC 0, hDC
    CalcHeight = tm.tmHeight
End Function

Public Function IsIntersectRect(rcFirst As RECT, rcSecond As RECT) As Boolean
    Dim rc              As RECT
    
    On Error Resume Next
    IsIntersectRect = IntersectRect(rc, rcFirst, rcSecond) <> 0
End Function
