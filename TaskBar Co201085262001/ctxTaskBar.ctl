VERSION 5.00
Begin VB.UserControl ctxTaskBar 
   Alignable       =   -1  'True
   BackColor       =   &H80000018&
   CanGetFocus     =   0   'False
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   DrawWidth       =   1768
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   Begin VB.Image imgDefIcon 
      Height          =   192
      Left            =   2940
      Picture         =   "ctxTaskBar.ctx":0000
      Top             =   420
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgCapture 
      Height          =   1104
      Left            =   1008
      Top             =   1008
      Width           =   1440
   End
End
Attribute VB_Name = "ctxTaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event StartMenu()
Event BeforeTaskSwitch(ByVal NewTask As Long, Cancel As Boolean)
Attribute BeforeTaskSwitch.VB_MemberFlags = "200"
Event TrayMouseDown(ByVal Idx As Long, ByVal Button As Long, ByVal X As Long, ByVal Y As Long)
Event TrayMouseUp(ByVal Idx As Long, ByVal Button As Long, ByVal X As Long, ByVal Y As Long)
Event TaskMouseDown(ByVal Idx As Long, ByVal Button As Long, ByVal X As Long, ByVal Y As Long)
Event TaskMouseUp(ByVal Idx As Long, ByVal Button As Long, ByVal X As Long, ByVal Y As Long)

'=========================================================================
' Constants and variables
'=========================================================================

Private Const CONTROL_HEIGHT    As Long = 30
Private Const SEP_WIDTH         As Long = 14
Private Const BRD_WIDTH         As Long = 3
Private Const MIN_BUTTON_WIDTH  As Long = 16 + 2 * BRD_WIDTH
Private Const MAX_BUTTON_WIDTH  As Long = 165
Private Const SNG_HIMETRIC      As Single = 21.66667
Private Const TRAY_PADDING      As Long = 4
Private Const DEF_STARTMENUCAPTION As String = "Start"
Private Const DEF_STARTMENUTOOLTIPCAPTION As String = "Click here to begin"
Private Const DEF_BUFFERDRAW    As Boolean = False

Private m_lActive           As Long
Private m_lPressed          As Long
Private m_oStartMenuIcon    As StdPicture
Private m_sStartMenuCaption As String
Private m_sStartMenuTooltipText As String
Private m_bBufferDraw       As Boolean

Private m_cTasks            As cTaskBarInfos
Private m_cTrayIcons        As cTaskBarInfos
Private m_lControlLines     As Long
Private m_lButtonWidth      As Long
Private m_lButtonsPerLine   As Long
Private m_bResizing         As Boolean

'=========================================================================
' Properties
'=========================================================================

Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Property Set Font(oValue As StdFont)
    Set UserControl.Font = oValue
    UserControl.Font.Bold = False
    UserControl_Resize
    PropertyChanged "Font"
End Property

Property Get StartMenuIcon() As StdPicture
    Set StartMenuIcon = m_oStartMenuIcon
End Property

Property Set StartMenuIcon(oValue As StdPicture)
    On Error Resume Next
    Set m_oStartMenuIcon = oValue
    Refresh
    PropertyChanged "StartMenuIcon"
End Property

Property Get StartMenuCaption() As String
Attribute StartMenuCaption.VB_UserMemId = -518
    StartMenuCaption = m_sStartMenuCaption
End Property

Property Let StartMenuCaption(sValue As String)
    m_sStartMenuCaption = sValue
    Refresh
    PropertyChanged "StartMenuCaption"
End Property

Property Get StartMenuTooltipText() As String
    StartMenuTooltipText = m_sStartMenuTooltipText
End Property

Property Let StartMenuTooltipText(sValue As String)
    m_sStartMenuTooltipText = sValue
    PropertyChanged "StartMenuTooltipText"
End Property

Property Get ActiveTask() As Long
Attribute ActiveTask.VB_MemberFlags = "400"
    ActiveTask = m_lActive
End Property

Property Let ActiveTask(ByVal lValue As Long)
    Dim rcClient As RECT
    
    On Error Resume Next
    GetClientRect hWnd, rcClient
    If m_lActive >= 0 Then
        InvalidateRect hWnd, GetButtonRect(rcClient, m_lActive), 0
    End If
    If lValue >= 0 And lValue <= Tasks.Count Then
        m_lActive = lValue
    Else
        m_lActive = -1
    End If
    If m_lActive >= 0 Then
        InvalidateRect hWnd, GetButtonRect(rcClient, m_lActive), 0
    End If
    '--- fix redraw in IDE
    Debug.Print DebugRefresh;
End Property

Property Get PressedTask() As Long
    PressedTask = m_lPressed
End Property

Property Let PressedTask(ByVal lValue As Long)
    Dim rcClient As RECT
    
    On Error Resume Next
    GetClientRect hWnd, rcClient
    If m_lPressed <> 0 Then
        InvalidateRect hWnd, GetButtonRect(rcClient, Abs(m_lPressed)), 0
    End If
    If lValue >= -Tasks.Count And lValue <= Tasks.Count Then
        m_lPressed = lValue
    Else
        m_lPressed = 0
    End If
    If m_lPressed <> 0 Then
        InvalidateRect hWnd, GetButtonRect(rcClient, Abs(m_lPressed)), 0
    End If
    '--- fix redraw in IDE
    Debug.Print DebugRefresh;
End Property

Property Get BufferDraw() As Boolean
    BufferDraw = m_bBufferDraw
End Property

Property Let BufferDraw(ByVal bValue As Boolean)
    m_bBufferDraw = bValue
    PropertyChanged "BufferDraw"
End Property

'= Collections ===========================================================

Property Get Tasks() As cTaskBarInfos
Attribute Tasks.VB_MemberFlags = "400"
    Set Tasks = m_cTasks
End Property

Property Get TrayIcons() As cTaskBarInfos
Attribute TrayIcons.VB_MemberFlags = "400"
    Set TrayIcons = m_cTrayIcons
End Property

'= Private ===============================================================

Private Property Get TrayWidth() As Long
    Dim lTrayLines  As Long
    On Error Resume Next
    If m_lControlLines = 1 Then
        TrayWidth = SEP_WIDTH + ClockWidth + TrayIcons.Count * (16 + TRAY_PADDING) + 2 * BRD_WIDTH + 2 + TRAY_PADDING
    Else
        lTrayLines = (GetControlHeight() - GetButtonHeight() - 2 * BRD_WIDTH) \ (16 + TRAY_PADDING)
        TrayWidth = (16 + TRAY_PADDING) * ((TrayIcons.Count + lTrayLines - 1) \ lTrayLines)
        If TrayWidth < ClockWidth Then
            TrayWidth = ClockWidth
        End If
        TrayWidth = TrayWidth + SEP_WIDTH + 2 * BRD_WIDTH + 2 + TRAY_PADDING
    End If
End Property

Private Property Get StartMenuWidth() As Long
    On Error Resume Next
    UserControl.Font.Bold = True
    StartMenuWidth = TextWidth(StartMenuCaption) / Screen.TwipsPerPixelX + 16 + 4 * BRD_WIDTH + 3
    UserControl.Font.Bold = False
End Property

Private Property Get ClockWidth() As Long
    On Error Resume Next
    ClockWidth = TextWidth(Format(Now, STR_SHORTTIME)) / Screen.TwipsPerPixelX + 6 * BRD_WIDTH
End Property

Private Property Get DEF_STARTMENUICON() As StdPicture
    Set DEF_STARTMENUICON = imgDefIcon.Picture
End Property

Private Property Get DEF_FONT() As StdFont
    Set DEF_FONT = New StdFont
    DEF_FONT.Name = "Tahoma"
    DEF_FONT.Size = 8
End Property

'=========================================================================
' Methods
'=========================================================================

Private Function DebugRefresh() As String
    Refresh
End Function

Private Function GetStartMenuRect(rcClient As RECT) As RECT
    On Error Resume Next
    GetStartMenuRect.Left = rcClient.Left
    GetStartMenuRect.Top = rcClient.Top
    GetStartMenuRect.Right = GetStartMenuRect.Left + StartMenuWidth + BRD_WIDTH
    GetStartMenuRect.Bottom = GetStartMenuRect.Top + GetLineHeight() + BRD_WIDTH
    InflateRect GetStartMenuRect, -BRD_WIDTH, -BRD_WIDTH
    If GetStartMenuRect.Right > rcClient.Right - TrayWidth - SEP_WIDTH Then
        GetStartMenuRect.Left = -1
        GetStartMenuRect.Top = -1
        GetStartMenuRect.Right = -1
        GetStartMenuRect.Bottom = -1
    End If
End Function

Private Function GetButtonRect(rcClient As RECT, ByVal lIdx As Long) As RECT
    Dim rcStart             As RECT
    Dim lButtonLine         As Long
    
    On Error Resume Next
    If lIdx = 0 Then
        GetButtonRect = GetStartMenuRect(rcClient)
    Else
        lButtonLine = (lIdx - 1) \ m_lButtonsPerLine
        GetButtonRect.Left = GetStartMenuRect(rcClient).Right + SEP_WIDTH + (lIdx - 1 - lButtonLine * m_lButtonsPerLine) * m_lButtonWidth - BRD_WIDTH
        GetButtonRect.Top = rcClient.Top + lButtonLine * GetLineHeight()
        GetButtonRect.Right = GetButtonRect.Left + m_lButtonWidth + BRD_WIDTH
        GetButtonRect.Bottom = GetButtonRect.Top + GetLineHeight() + BRD_WIDTH
        InflateRect GetButtonRect, -BRD_WIDTH, -BRD_WIDTH
        If GetButtonRect.Right < GetButtonRect.Left + MIN_BUTTON_WIDTH Then
            GetButtonRect.Left = -1
            GetButtonRect.Top = -1
            GetButtonRect.Right = -1
            GetButtonRect.Bottom = -1
        End If
    End If
End Function

Private Function GetTrayRect(rcClient As RECT) As RECT
    On Error Resume Next
    GetTrayRect.Left = rcClient.Right - TrayWidth + SEP_WIDTH
    GetTrayRect.Top = rcClient.Top
    GetTrayRect.Right = rcClient.Right
    GetTrayRect.Bottom = rcClient.Bottom
    InflateRect GetTrayRect, -BRD_WIDTH, -BRD_WIDTH
End Function

Private Function GetClockRect(rcClient As RECT) As RECT
    On Error Resume Next
    GetClockRect = GetTrayRect(rcClient)
    InflateRect GetClockRect, -1, -1
    If m_lControlLines = 1 Then
        GetClockRect.Left = GetClockRect.Right - ClockWidth
    End If
    GetClockRect.Bottom = GetClockRect.Top + GetButtonHeight() - 2
End Function

Private Function GetTrayIconRect(rcClient As RECT, ByVal lIdx As Long) As RECT
    Dim rcClock         As RECT
    Dim rcTray          As RECT
    Dim lIconsPerLine   As Long
    Dim lIconRow        As Long
    
    On Error Resume Next
    rcClock = GetClockRect(rcClient)
    If m_lControlLines = 1 Then
        GetTrayIconRect.Left = rcClock.Left - (16 + TRAY_PADDING) * (TrayIcons.Count - lIdx + 1) - TRAY_PADDING \ 2
        GetTrayIconRect.Top = rcClock.Top + (rcClock.Bottom - rcClock.Top - 16 - TRAY_PADDING) \ 2
    Else
        lIconsPerLine = (TrayWidth - SEP_WIDTH - 2 * BRD_WIDTH - TRAY_PADDING) \ (16 + TRAY_PADDING)
        lIconRow = (lIdx - 1) \ lIconsPerLine
        GetTrayIconRect.Left = rcClock.Left + (16 + TRAY_PADDING) * (lIdx - lIconRow * lIconsPerLine - 1) + TRAY_PADDING \ 2
        GetTrayIconRect.Top = rcClock.Bottom + (16 + TRAY_PADDING) * lIconRow
    End If
    GetTrayIconRect.Right = GetTrayIconRect.Left + 16 + TRAY_PADDING - 1
    GetTrayIconRect.Bottom = GetTrayIconRect.Top + 16 + TRAY_PADDING - 1
End Function

Private Function GetButtonHeight() As Long
    On Error Resume Next
    GetButtonHeight = CalcHeight(Font)
    If GetButtonHeight < 19 Then
        GetButtonHeight = 19
    End If
    GetButtonHeight = GetButtonHeight + 2 * BRD_WIDTH
End Function

Private Function GetLineHeight() As Long
    On Error Resume Next
    GetLineHeight = GetButtonHeight() + BRD_WIDTH
End Function

Private Function GetControlHeight() As Long
    GetControlHeight = GetLineHeight() * m_lControlLines + BRD_WIDTH
End Function

Private Function GetMaxControlLines() As Long
    Dim rcClient        As RECT
    
    On Error Resume Next
    GetClientRect GetParent(UserControl.hWnd), rcClient
    GetMaxControlLines = (rcClient.Bottom - rcClient.Top) \ 2 \ GetLineHeight()
End Function

Private Sub DrawSeparator(ByVal hdc As Long, rc As RECT)
    Dim hDarkPen        As Long
    Dim hLightPen       As Long
    Dim hPrevPen        As Long
    
    On Error Resume Next
    hDarkPen = CreatePen(0, 1, TranslateColor(vb3DShadow))
    hLightPen = CreatePen(0, 1, TranslateColor(vb3DHighlight))
    InflateRect rc, 0, -1
    ExtTextOut hdc, 0, 0, ETO_OPAQUE, rc, "", 0, 0
    InflateRect rc, 0, -2
    '--- dark
    hPrevPen = SelectObject(hdc, hDarkPen)
    MoveToEx hdc, rc.Left + BRD_WIDTH, rc.Top, ByVal 0
    LineTo hdc, rc.Left + BRD_WIDTH, rc.Bottom
    MoveToEx hdc, rc.Left + BRD_WIDTH + 6, rc.Top + BRD_WIDTH, ByVal 0
    LineTo hdc, rc.Left + BRD_WIDTH + 6, rc.Bottom - BRD_WIDTH
    LineTo hdc, rc.Left + BRD_WIDTH + 4, rc.Bottom - BRD_WIDTH
    '--- light
    SelectObject hdc, hLightPen
    MoveToEx hdc, rc.Left + BRD_WIDTH + 1, rc.Top, ByVal 0
    LineTo hdc, rc.Left + BRD_WIDTH + 1, rc.Bottom
    MoveToEx hdc, rc.Left + BRD_WIDTH + 5, rc.Top + BRD_WIDTH, ByVal 0
    LineTo hdc, rc.Left + BRD_WIDTH + 4, rc.Top + BRD_WIDTH
    LineTo hdc, rc.Left + BRD_WIDTH + 4, rc.Bottom - BRD_WIDTH - 1
    '--- cleanup
    SelectObject hdc, hPrevPen
    DeleteObject hDarkPen
    DeleteObject hLightPen
End Sub

Private Sub DrawPicture( _
        ByVal hdc As Long, _
        oPic As StdPicture, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cxWidth As Long, _
        ByVal cyHeight As Long)
    Dim hMemDC          As Long
    Dim hPrevBmp        As Long
    Dim rc              As RECT
    Dim hEmf            As Long
    
    On Error Resume Next
    Select Case oPic.Type
    Case vbPicTypeIcon
        DrawIconEx hdc, X, Y, oPic.Handle, cxWidth, cyHeight, 0, 0, DI_NORMAL
    Case vbPicTypeBitmap
        hMemDC = CreateCompatibleDC(hdc)
        hPrevBmp = SelectObject(hMemDC, oPic.Handle)
        StretchBlt hdc, X, Y, cxWidth, cyHeight, hMemDC, 0, 0, oPic.Width / SNG_HIMETRIC, oPic.Height / SNG_HIMETRIC, SRCCOPY
        SelectObject hMemDC, hPrevBmp
        DeleteDC hMemDC
    Case vbPicTypeEMetafile, vbPicTypeMetafile
        rc.Left = X: rc.Top = Y
        rc.Right = X + cxWidth: rc.Bottom = Y + cyHeight
        If oPic.Type = vbPicTypeMetafile Then
            hMemDC = CreateEnhMetaFileLong(hdc, vbNullString, 0, vbNullString)
            PlayMetaFile hMemDC, oPic.Handle
            hEmf = CloseEnhMetaFile(hMemDC)
        Else
            hEmf = oPic.Handle
        End If
        PlayEnhMetaFile hdc, hEmf, rc
        If oPic.Type = vbPicTypeMetafile Then
            DeleteEnhMetaFile hEmf
        End If
    End Select
End Sub

Private Sub DrawControl()
    Dim rcClient        As RECT
    Dim rcButton        As RECT
    Dim rcTray          As RECT
    Dim hRgn            As Long
    Dim hRgnButton      As Long
    Dim lI              As Long
    Dim lPrevBkColor    As Long
    Dim sText           As String
    Dim oTextMetric     As TEXTMETRIC
    Dim lX              As Long
    Dim lY              As Long
    Dim oPic            As StdPicture
    Dim hWinDC          As Long
    Dim hdc             As Long
    Dim hBmp            As Long
    Dim hPrevBmp        As Long
    Dim hFont           As Long
    Dim hBoldFont       As Long
    Dim hPrevFont       As Long
    
    
    On Error Resume Next
    '--- if anything to paint
    '--- prepare
    GetClientRect hWnd, rcClient
    hWinDC = GetDC(hWnd)
If BufferDraw Then
    hdc = CreateCompatibleDC(hWinDC)
    hBmp = CreateCompatibleBitmap(hWinDC, UpdateRect.Right - UpdateRect.Left, UpdateRect.Bottom - UpdateRect.Top)
    hPrevBmp = SelectObject(hdc, hBmp)
    lPrevBkColor = SetBkColor(hdc, TranslateColor(vbButtonFace))
    ExtTextOut hdc, 0, 0, ETO_OPAQUE, rcClient, "", 0, 0
    SetViewportOrgEx hdc, -UpdateRect.Left, -UpdateRect.Top, ByVal 0
Else
    hdc = hWinDC
    lPrevBkColor = SetBkColor(hdc, TranslateColor(vbButtonFace))
End If
    hFont = CreateLogFont(hdc, Font)
    Font.Bold = True
    hBoldFont = CreateLogFont(hdc, Font)
    Font.Bold = False
    hPrevFont = SelectObject(hdc, hFont)
    hRgn = CreateRectRgnIndirect(rcClient)
    SetBkMode hdc, TRANSPARENT
    GetTextMetrics hdc, oTextMetric
    lX = Loword(GetDialogBaseUnits) / 4 + 1
    SetStretchBltMode hdc, STRETCH_HALFTONE
    
    '--- calc button properties
    m_lButtonWidth = (rcClient.Right - rcClient.Left - TrayWidth - GetStartMenuRect(rcClient).Right - SEP_WIDTH) \ ((Tasks.Count + m_lControlLines - 1) \ m_lControlLines)
    If m_lButtonWidth > MAX_BUTTON_WIDTH Then
        m_lButtonWidth = MAX_BUTTON_WIDTH
    End If
    m_lButtonsPerLine = (rcClient.Right - rcClient.Left - TrayWidth - GetStartMenuRect(rcClient).Right - SEP_WIDTH) \ m_lButtonWidth
    
    '--- start button
    rcButton = GetStartMenuRect(rcClient)
    If IsIntersectRect(rcButton, UpdateRect) Then
        DrawFrameControl hdc, rcButton, DFC_BUTTON, IIf(m_lActive = 0, DFCS_BUTTONPUSH Or DFCS_PUSHED Or DFCS_CHECKED, DFCS_BUTTONPUSH)
        hRgnButton = CreateRectRgnIndirect(rcButton)
        CombineRgn hRgn, hRgn, hRgnButton, RGN_DIFF
        DeleteObject hRgnButton
        InflateRect rcButton, -BRD_WIDTH, -BRD_WIDTH
        If 0 = m_lActive Then
            OffsetRect rcButton, 1, 1
            rcButton.Right = rcButton.Right - 1
        End If
        Set oPic = StartMenuIcon
        If Not oPic Is Nothing Then
            DrawPicture hdc, oPic, rcButton.Left, rcButton.Top + (rcButton.Bottom - rcButton.Top - 16) \ 2, 16, 16
            rcButton.Left = rcButton.Left + 17
        End If
        sText = StartMenuCaption
        lY = (rcButton.Bottom + rcButton.Top - oTextMetric.tmHeight) \ 2
        SelectObject hdc, hBoldFont
        SetTextAlign hdc, TA_LEFT
        ExtTextOut hdc, rcButton.Left + lX, lY, ETO_CLIPPED, rcButton, sText, Len(sText), 0
        SelectObject hdc, hFont
        rcButton.Right = rcButton.Right + BRD_WIDTH
    End If
    
    '--- taskbar
    '--- separator
    rcButton.Left = rcButton.Right
    rcButton.Top = rcClient.Top
    rcButton.Right = rcButton.Left + SEP_WIDTH
    rcButton.Bottom = rcClient.Bottom
    DrawSeparator hdc, rcButton
    hRgnButton = CreateRectRgnIndirect(rcButton)
    CombineRgn hRgn, hRgn, hRgnButton, RGN_DIFF
    DeleteObject hRgnButton
    '--- buttons
    For lI = 1 To Tasks.Count
        rcButton = GetButtonRect(rcClient, lI)
        If IsIntersectRect(rcButton, UpdateRect) Then
            DrawFrameControl hdc, rcButton, DFC_BUTTON, DFCS_BUTTONPUSH Or _
                Switch(lI = m_lActive, DFCS_PUSHED Or DFCS_CHECKED, _
                    lI = m_lPressed, DFCS_PUSHED)
            hRgnButton = CreateRectRgnIndirect(rcButton)
            CombineRgn hRgn, hRgn, hRgnButton, RGN_DIFF
            DeleteObject hRgnButton
            InflateRect rcButton, -BRD_WIDTH, -BRD_WIDTH
            If lI = m_lActive Then
                OffsetRect rcButton, 0, 1
                SelectObject hdc, hBoldFont
            End If
            Set oPic = Nothing
            Set oPic = m_cTasks(lI).Icon
            If Not oPic Is Nothing Then
                DrawPicture hdc, oPic, rcButton.Left, rcButton.Top + (rcButton.Bottom - rcButton.Top - 16) \ 2, 16, 16
                rcButton.Left = rcButton.Left + 17
            End If
            sText = ""
            sText = m_cTasks(lI).Caption
            PathCompactPath hdc, sText, rcButton.Right - rcButton.Left - lX - 1
            If InStr(sText, Chr(0)) > 0 Then
                sText = Left(sText, InStr(sText, Chr(0)) - 1)
            End If
            lY = (rcButton.Bottom + rcButton.Top - oTextMetric.tmHeight) \ 2
            SetTextAlign hdc, TA_LEFT
            ExtTextOut hdc, rcButton.Left + lX, lY, ETO_CLIPPED, rcButton, sText, Len(sText), 0
            SelectObject hdc, hFont
        End If
    Next
    '--- tray
    rcTray = GetTrayRect(rcClient)
    rcButton.Left = rcTray.Left - SEP_WIDTH
    rcButton.Top = rcClient.Top
    rcButton.Right = rcButton.Left + SEP_WIDTH
    rcButton.Bottom = rcClient.Bottom
    '--- separator
    DrawSeparator hdc, rcButton
    hRgnButton = CreateRectRgnIndirect(rcButton)
    CombineRgn hRgn, hRgn, hRgnButton, RGN_DIFF
    DeleteObject hRgnButton
    If IsIntersectRect(rcTray, UpdateRect) Then
        '--- edge
        rcButton = rcTray
        ExtTextOut hdc, 0, 0, ETO_OPAQUE, rcButton, "", 0, 0
        DrawEdge hdc, rcButton, BDR_SUNKENOUTER, BF_RECT
        hRgnButton = CreateRectRgnIndirect(rcButton)
        CombineRgn hRgn, hRgn, hRgnButton, RGN_DIFF
        DeleteObject hRgnButton
        '--- clock
        rcButton = GetClockRect(rcClient)
        sText = Format(Now, STR_SHORTTIME)
        lY = (rcButton.Bottom + rcButton.Top - oTextMetric.tmHeight) \ 2
        SetTextAlign hdc, TA_CENTER
        ExtTextOut hdc, (rcButton.Left + rcButton.Right) \ 2, lY, ETO_OPAQUE Or ETO_CLIPPED, rcButton, sText, Len(sText), 0
        '--- tray icons
        For lI = 1 To TrayIcons.Count
            rcButton = GetTrayIconRect(rcClient, lI)
            Set oPic = Nothing
            Set oPic = m_cTrayIcons(lI).Icon
            If Not oPic Is Nothing Then
                DrawPicture hdc, oPic, rcButton.Left + TRAY_PADDING \ 2, rcButton.Top + (rcButton.Bottom - rcButton.Top - 16) \ 2, 16, 16
            End If
        Next
    End If
    '--- background and outer edge
    DrawEdge hdc, rcClient, BDR_RAISEDINNER, BF_RECT
If BufferDraw Then
    SetViewportOrgEx hdc, 0, 0, ByVal 0
    '-- bit-blit
    BitBlt hWinDC, UpdateRect.Left, UpdateRect.Top, UpdateRect.Right - UpdateRect.Left, UpdateRect.Bottom - UpdateRect.Top, hdc, 0, 0, SRCCOPY
Else
    SelectClipRgn hdc, hRgn
    InflateRect rcClient, -1, -1
    ExtTextOut hdc, 0, 0, ETO_OPAQUE, rcClient, "", 0, 0
    InflateRect rcClient, 1, 1
End If
    '--- clean up
    SetBkColor hdc, lPrevBkColor
    SelectClipRgn hdc, 0
    DeleteObject hRgn
    SelectObject hdc, hPrevBmp
    DeleteObject hBmp
    SelectObject hdc, hPrevFont
    DeleteObject hFont
    DeleteObject hBoldFont
If BufferDraw Then
    DeleteDC hdc
Else
    ReleaseDC hWnd, hdc
End If
    ReleaseDC hWnd, hWinDC
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub UserControl_Initialize()
    On Error Resume Next
    m_lActive = -1
    m_lControlLines = 1
    Set m_cTasks = New cTaskBarInfos
    m_cTasks.hWnd = UserControl.hWnd
    Set m_cTrayIcons = New cTaskBarInfos
    m_cTrayIcons.hWnd = UserControl.hWnd
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    PaintUnsubclass hWnd
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'--- received upon WM_CANCELMODE
    If PressedTask <> 0 Then
        PressedTask = 0
    End If
    '--- stop sizing
    m_bResizing = False
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
'--- received upon WM_PAINT
    DrawControl
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    Set Font = Ambient.Font
    StartMenuCaption = DEF_STARTMENUCAPTION
    StartMenuTooltipText = DEF_STARTMENUTOOLTIPCAPTION
    Set StartMenuIcon = DEF_STARTMENUICON
    BufferDraw = DEF_BUFFERDRAW
    PaintSubclass hWnd, IIf(Ambient.UserMode, 1000, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        Set Font = .ReadProperty("Font", DEF_FONT)
        StartMenuCaption = .ReadProperty("StartMenuCaption", DEF_STARTMENUCAPTION)
        StartMenuTooltipText = .ReadProperty("StartMenuTooltipText", DEF_STARTMENUTOOLTIPCAPTION)
        Set StartMenuIcon = .ReadProperty("StartMenuIcon", DEF_STARTMENUICON)
        BufferDraw = .ReadProperty("BufferDraw", DEF_BUFFERDRAW)
    End With
    PaintSubclass hWnd, IIf(Ambient.UserMode, 1000, 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        .WriteProperty "Font", Font, DEF_FONT
        .WriteProperty "StartMenuCaption", StartMenuCaption, DEF_STARTMENUCAPTION
        .WriteProperty "StartMenuTooltipText", StartMenuTooltipText, DEF_STARTMENUTOOLTIPCAPTION
        .WriteProperty "StartMenuIcon", StartMenuIcon, DEF_STARTMENUICON
        .WriteProperty "BufferDraw", BufferDraw, DEF_BUFFERDRAW
    End With
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Height = ScaleY(GetControlHeight(), vbPixels)
    imgCapture.Move 0, 0, ScaleWidth, ScaleHeight
    Refresh
End Sub

Private Sub imgCapture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pt              As POINTAPI
    Dim rcClient        As RECT
    Dim lI              As Long
    Dim rc              As RECT
    
    On Error Resume Next
    GetClientRect hWnd, rcClient
    pt.X = ScaleX(X, vbTwips, vbPixels)
    pt.Y = ScaleY(Y, vbTwips, vbPixels)
    '--- check if sizing
    If imgCapture.MousePointer = vbSizeNS And Button = vbLeftButton Then
        m_bResizing = True
        Exit Sub
    End If
    '--- ignore event if sizing
    If m_bResizing Then
        Exit Sub
    End If
    '--- hittest start menu
    rc = GetStartMenuRect(rcClient)
    If PtInRect(rc, pt) Then
        RaiseEvent TaskMouseDown(0, Button, pt.X - rc.Left, pt.Y - rc.Top)
        If Button = vbLeftButton Then
            ActiveTask = 0
            RaiseEvent StartMenu
            ActiveTask = -1
        End If
        Exit Sub
    End If
    '--- hittest tasks
    For lI = 1 To Tasks.Count
        rc = GetButtonRect(rcClient, lI)
        If PtInRect(rc, pt) Then
            RaiseEvent TaskMouseDown(lI, Button, pt.X - rc.Left, pt.Y - rc.Top)
            If Button = vbLeftButton Then
                PressedTask = lI
            End If
            Exit Sub
        End If
    Next
    '--- hittest clock
    rc = GetClockRect(rcClient)
    If PtInRect(rc, pt) Then
        RaiseEvent TrayMouseDown(0, Button, pt.X - rc.Left, pt.Y - rc.Top)
        Exit Sub
    End If
    '--- hittest tray icons
    For lI = 1 To TrayIcons.Count
        rc = GetTrayIconRect(rcClient, lI)
        If PtInRect(rc, pt) Then
            RaiseEvent TrayMouseDown(lI, Button, pt.X - rc.Left, pt.Y - rc.Top)
            Exit Sub
        End If
    Next
End Sub

Private Sub imgCapture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pt              As POINTAPI
    Dim rcClient        As RECT
    Dim lI              As Long
    Dim rc              As RECT
    Dim lControlLines   As Long
    
    On Error Resume Next
    GetClientRect hWnd, rcClient
    pt.X = ScaleX(X, vbTwips, vbPixels)
    pt.Y = ScaleY(Y, vbTwips, vbPixels)
    If m_bResizing Then
        '--- figure out new control size
        If Extender.Align = vbAlignBottom Then
            lControlLines = m_lControlLines - Round(ScaleY(Y, vbTwips, vbPixels) / GetLineHeight())
        Else
            lControlLines = m_lControlLines + Round((ScaleY(Y, vbTwips, vbPixels) - rcClient.Bottom) / GetLineHeight())
        End If
        If lControlLines < 1 Then
            lControlLines = 1
        ElseIf lControlLines > GetMaxControlLines() Then
            lControlLines = GetMaxControlLines()
        End If
        If lControlLines <> m_lControlLines Then
            m_lControlLines = lControlLines
            UserControl_Resize
        End If
        Exit Sub
    End If
    '--- change mouse pointer
    If Extender.Align = vbAlignBottom Then
        imgCapture.MousePointer = IIf(pt.Y >= 0 And pt.Y <= BRD_WIDTH, vbSizeNS, vbDefault)
    Else
        imgCapture.MousePointer = IIf(pt.Y >= rcClient.Bottom - BRD_WIDTH And pt.Y <= rcClient.Bottom, vbSizeNS, vbDefault)
    End If
    '--- if task pressed
    If (Button And vbLeftButton) <> 0 And m_lPressed <> 0 Then
        If m_lPressed > 0 Then
            If Not PtInRect(GetButtonRect(rcClient, m_lPressed), pt) Then
                PressedTask = -PressedTask
            End If
        Else
            If PtInRect(GetButtonRect(rcClient, -m_lPressed), pt) Then
                PressedTask = -PressedTask
            End If
        End If
    End If
    '--- hittest start menu
    rc = GetStartMenuRect(rcClient)
    If PtInRect(rc, pt) Then
        If StartMenuTooltipText <> "" Then
            imgCapture.TooltipText = StartMenuTooltipText
        Else
            imgCapture.TooltipText = StartMenuCaption
        End If
        Exit Sub
    End If
    '--- hittest tasks
    For lI = 1 To Tasks.Count
        rc = GetButtonRect(rcClient, lI)
        If PtInRect(rc, pt) Then
            If Tasks(lI).TooltipText <> "" Then
                imgCapture.TooltipText = Tasks(lI).TooltipText
            Else
                imgCapture.TooltipText = Tasks(lI).Caption
            End If
            Exit Sub
        End If
    Next
    '--- hittest tray icons
    For lI = 1 To TrayIcons.Count
        rc = GetTrayIconRect(rcClient, lI)
        If PtInRect(rc, pt) Then
            If TrayIcons(lI).TooltipText <> "" Then
                imgCapture.TooltipText = TrayIcons(lI).TooltipText
            Else
                imgCapture.TooltipText = TrayIcons(lI).Caption
            End If
            Exit Sub
        End If
    Next
    '--- hittest clock
    rc = GetClockRect(rcClient)
    If PtInRect(rc, pt) Then
        imgCapture.TooltipText = Format(Date, STR_LONGDATE)
        Exit Sub
    End If
    '--- no tooltip
    imgCapture.TooltipText = ""
End Sub

Private Sub imgCapture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bCancel         As Boolean
    Dim pt              As POINTAPI
    Dim rcClient        As RECT
    Dim rc              As RECT
    Dim lI              As Long
    
    On Error Resume Next
    GetClientRect hWnd, rcClient
    pt.X = ScaleX(X, vbTwips, vbPixels)
    pt.Y = ScaleY(Y, vbTwips, vbPixels)
    '--- stop sizing if left button
    If m_bResizing Then
        If Button = vbLeftButton Then
            m_bResizing = False
        End If
        Exit Sub
    End If
    '--- hittest start menu
    rc = GetStartMenuRect(rcClient)
    If PtInRect(rc, pt) Then
        RaiseEvent TaskMouseUp(0, Button, pt.X - rc.Left, pt.Y - rc.Top)
        Exit Sub
    End If
    '--- hittest tasks
    For lI = 1 To Tasks.Count
        rc = GetButtonRect(rcClient, lI)
        If PtInRect(rc, pt) Then
            RaiseEvent TaskMouseUp(lI, Button, pt.X - rc.Left, pt.Y - rc.Top)
            If Button = vbLeftButton Then
                If m_lPressed <> 0 Then
                    If m_lPressed > 0 Then
                        RaiseEvent BeforeTaskSwitch(m_lPressed, bCancel)
                        If Not bCancel Then
                            ActiveTask = m_lPressed
                        End If
                    End If
                    PressedTask = 0
                End If
            End If
            Exit Sub
        End If
    Next
    '--- hittest clock
    rc = GetClockRect(rcClient)
    If PtInRect(rc, pt) Then
        RaiseEvent TrayMouseUp(0, Button, pt.X - rc.Left, pt.Y - rc.Top)
        Exit Sub
    End If
    '--- hittest tray icons
    For lI = 1 To TrayIcons.Count
        rc = GetTrayIconRect(rcClient, lI)
        If PtInRect(rc, pt) Then
            RaiseEvent TrayMouseUp(lI, Button, pt.X - rc.Left, pt.Y - rc.Top)
        End If
    Next
End Sub

