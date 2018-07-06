Attribute VB_Name = "Module1"
Option Explicit

Public mlWndProc  As Long
Dim mlSetStyle    As Long
Dim mlHook        As Long
Dim mlHookWndProc As Long
Dim mhWndLast     As Long

Type CWPSTRUCT
    lParam  As Long
    wParam  As Long
    message As Long
    hwnd    As Long
End Type

Type CREATESTRUCT
    lpCreateParams As Long
    hInstance      As Long
    hMenu          As Long
    hWndParent     As Long
    cy             As Long
    cx             As Long
    Y              As Long
    X              As Long
    style          As Long
    lpszName       As Long 'String
    lpszClass      As Long 'String
    ExStyle        As Long
End Type

Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Type DRAWITEMSTRUCT
    CtlType    As Long
    CtlID      As Long
    itemID     As Long
    itemAction As Long
    itemState  As Long
    hwndItem   As Long
    hdc        As Long
    rcItem     As RECT
    itemData   As Long
End Type

Type MEASUREITEMSTRUCT
    CtlType    As Long
    CtlID      As Long
    itemID     As Long
    itemWidth  As Long
    itemHeight As Long
    itemData   As Long
End Type

Type SIZE
    cx As Long
    cy As Long
End Type
'
' Copies a block of memory from one location to another
' provided the blocks do not overlap.
' hpvDest   - pointer to address of copy destination
' hpvSource - pointer to address of block to copy
' cbCopy    - size, in bytes, of block to copy
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'
' Performs a bit-block transfer of color data corresponding to a
' rectangle of pixels from a source device context into a destination
' device context.
' hdcDest  - handle to destination device context
' x        - x-coordinate of destination rectangle's upper-left corner
' y        - y-coordinate of destination rectangle's upper-left corner
' nWidth   - width of destination rectangle
' nHeight  - height of destination rectangle
' hSrcDC   - handle to source device context
' xSrc     - x-coordinate of source rectangle's upper-left corner
' ySrc     - y-coordinate of source rectangle's upper-left corner
' dwRop    - raster operation code
Declare Function BitBlt Lib "gdi32" _
    (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'
' Creates a memory device context (DC) compatible with the specified device.
' hdc - handle to the device context
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'
' Gets a handle to a display device context for the client area of a
' specified window or for the entire screen.
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'
' The DeleteDC function deletes the specified device context (DC).
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'
' Selects an object into the specified device context.
' hdc     - handle to device context
' hObject - handle to object
Declare Function SelectObject Lib "gdi32" _
    (ByVal hdc As Long, ByVal hObject As Long) As Long
'
' Delete a logical pen, brush, font, bitmap, region, or palette, freeing
' all system resources associated with the object. After the object is
' deleted, the specified handle is no longer valid.
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'
' Gets various widths and heights of display elements and system settings.
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'
' Increase or decrease the width and height of the specified rectangle
' by adding x units to the left and right ends of the rectangle and y units
' to the top and bottom. Positive values increase the width and height,
' and negative values decrease them.
' lpRect - pointer to rectangle
' x      - amount to increase or decrease width
' y      - amount to increase or decrease height
Declare Function InflateRect Lib "user32" _
    (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
'
' Load the specified bitmap resource from a module's executable file.
' This function has been superseded by the LoadImage function.
' hInstance    ' handle to application instance
' lpBitmapName ' address of bitmap resource name
Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" _
    (ByVal hInstance As Long, lpBitmapName As Any) As Long
'
' Sends a message to a window or windows.
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long
'
' Gets the current color of the display element indicated by nIndex.
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'
' Sets the current background color to the specified color value,
' or to the nearest physical color if the device cannot represent
' the specified color value.
' hdc     - handle of device context
' crColor - background color value
Declare Function SetBkColor Lib "gdi32" _
    (ByVal hdc As Long, ByVal crColor As Long) As Long
'
' Create a logical font that has specific characteristics. The logical font
' can subsequently be selected as the font for any device.
' nHeight            - logical height of font
' nWidth             - logical average character width
' nEscapement        - angle of escapement
' nOrientation       - base-line orientation angle
' fnWeight           - font weight
' fdwItalic          - italic attribute flag
' fdwUnderline       - underline attribute flag
' fdwStrikeOut       - strikeout attribute flag
' fdwCharSet         - character set identifier
' fdwOutputPrecision - output precision
' fdwClipPrecision   - clipping precision
' fdwQuality         - output quality
' fdwPitchAndFamily  - pitch and family
' lpszFace           - pointer to typeface name string
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" _
    (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, _
    ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
    ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, _
    ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, _
    ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, _
    ByVal lpszFace As String) As Long
'
' Computes the width and height of the specified string of text.
' hdc       - handle to device context
' lpsz      - pointer to text string
' cbString  - number of characters in string
' lpSize    - pointer to structure for string size
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" _
    (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, _
    lpSize As SIZE) As Long
'
' Sets the text color for the specified device context to the specified color.
' hdc     - handle to device context
' crColor - text color
Declare Function SetTextColor Lib "gdi32" _
    (ByVal hdc As Long, ByVal crColor As Long) As Long
'
' Writes a character string at the specified location, using the currently
' selected font, background color, and text color.
' hdc      - handle to device context
' x        - x-coordinate of starting position
' y        - y-coordinate of starting position
' lpString - pointer to string
' nCount   - number of characters in string
Declare Function TextOut Lib "gdi32" Alias "TextOutA" _
    (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal lpString As String, ByVal nCount As Long) As Long
'
' Installs an application-defined hook procedure into a hook chain. You
' install a hook procedure to monitor the system for certain types of events.
' These events are associated either with a specific thread or with all threads
' in the system (system wide hooks are not available from VB).
' idHook     - type of hook to install
' lpfn       - address of hook procedure
' hMod       - handle to application instance
' dwThreadId - thread to install the hook for
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
    (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, _
    ByVal dwThreadId As Long) As Long
'
' This function passes the hook information to the next hook procedure
' in the current hook chain. A hook procedure can call this function either
' before or after processing the hook information.
' mlHook  - handle to current hook
' nCode  - hook code passed to hook procedure
' wParam - value passed to hook procedure
' lParam - value passed to hook procedure
Declare Function CallNextHookEx Lib "user32" _
    (ByVal mlHook As Long, ByVal ncode As Long, _
    ByVal wParam As Long, lParam As Any) As Long
'
' Removes a hook procedure installed in a hook chain.
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal mlHook As Long) As Long
'
' Get information about the specified windowand the value at the specified
' offset into the extra window memory of a window.
' hWnd   - handle of window
' nIndex - offset of value to retrieve
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'
' Changes an attribute of the specified window and sets a value
' at the specified offset into the extra window memory of a window.
' hWnd      - handle of window
' nIndex    - offset of value to set
' dwNewLong - new value
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'
' Pass message information to the specified window procedure.
' lpPrevWndFunc - pointer to previous procedure
' hWnd          - handle to window
' Msg           - message
' wParam        - first message parameter
' lParam        - second message parameter
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
' Retrieves the name of the class to which the specified window belongs.
' hWnd        - handle of window
' lpClassName - address of buffer for class name
' nMaxCount   - size of buffer, in characters
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As Long, ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long
'
' Gets a handle to the specified child window's parent window.
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Const WH_CALLWNDPROC = 4
Public Const SM_CXMENUSIZE = 54
'
' Misc Windows messages
'
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_KILLFOCUS = &H8
Public Const WM_SETTEXT = &HC
Public Const WM_PAINT = &HF
Public Const WM_ERASEBKGND = &H14
Public Const WM_DRAWITEM = &H2B
Public Const WM_NCPAINT = &H85
Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const WM_PARENTNOTIFY = &H210

Public Const CB_GETLBTEXT = &H148
Public Const CB_SETITEMDATA = &H151
Public Const LB_GETTEXT = &H189
Public Const LB_SETITEMHEIGHT = &H1A0
'
' Window field offsets for GetWindowLong and GetWindowWord APIs.
'
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const GWL_ID = (-12)
'
' OEM Resource Ordinal Numbers.
' Predefined bitmaps used by the Win32 API.
'
Public Const OBM_UPARROW = 32753
Public Const OBM_DNARROW = 32752
Public Const OBM_RGARROW = 32751
Public Const OBM_LFARROW = 32750
Public Const OBM_REDUCE = 32749
Public Const OBM_ZOOM = 32748
Public Const OBM_RESTORE = 32747
Public Const OBM_REDUCED = 32746
Public Const OBM_ZOOMD = 32745
Public Const OBM_RESTORED = 32744
Public Const OBM_UPARROWD = 32743
Public Const OBM_DNARROWD = 32742
Public Const OBM_RGARROWD = 32741
Public Const OBM_LFARROWD = 32740
Public Const OBM_MNARROW = 32739
Public Const OBM_COMBO = 32738
Public Const OBM_UPARROWI = 32737
Public Const OBM_DNARROWI = 32736
Public Const OBM_RGARROWI = 32735
Public Const OBM_LFARROWI = 32734
'
' GetSysColor colors.
'
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
'
' Window Styles.
'
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000  'WS_BORDER Or WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000

Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000

Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
'
' Common Window Styles.
'
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_CHILDWINDOW = (WS_CHILD)
'
' Extended Window Styles.
'
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&
'
' Button Control Styles.
'
Public Const BS_PUSHBUTTON = &H0
Public Const BS_DEFPUSHBUTTON = &H1
Public Const BS_CHECKBOX = &H2
Public Const BS_AUTOCHECKBOX = &H3
Public Const BS_RADIOBUTTON = &H4
Public Const BS_3STATE = &H5
Public Const BS_AUTO3STATE = &H6
Public Const BS_GROUPBOX = &H7
Public Const BS_USERBUTTON = &H8
Public Const BS_AUTORADIOBUTTON = &H9
Public Const BS_OWNERDRAW = &HB
Public Const BS_LEFTTEXT = &H20
Public Const BS_TEXT = &H0
Public Const BS_ICON = &H40
Public Const BS_BITMAP = &H80
Public Const BS_LEFT = &H100
Public Const BS_RIGHT = &H200
Public Const BS_CENTER = &H300
Public Const BS_TOP = &H400
Public Const BS_BOTTOM = &H800
Public Const BS_VCENTER = &HC00
Public Const BS_PUSHLIKE = &H1000
Public Const BS_MULTILINE = &H2000
Public Const BS_NOTIFY = &H4000
Public Const BS_FLAT = &H8000
Public Const BS_RIGHTBUTTON = BS_LEFTTEXT
'
' Combo Box styles.
'
Public Const CBS_SIMPLE = &H1&
Public Const CBS_DROPDOWN = &H2&
Public Const CBS_DROPDOWNLIST = &H3&
Public Const CBS_OWNERDRAWFIXED = &H10&
Public Const CBS_OWNERDRAWVARIABLE = &H20&
Public Const CBS_AUTOHSCROLL = &H40&
Public Const CBS_OEMCONVERT = &H80&
Public Const CBS_SORT = &H100&
Public Const CBS_HASSTRINGS = &H200&
Public Const CBS_NOINTEGRALHEIGHT = &H400&
Public Const CBS_DISABLENOSCROLL = &H800&
'
' Edit Control Styles.
'
Public Const ES_LEFT = &H0&
Public Const ES_CENTER = &H1&
Public Const ES_RIGHT = &H2&
Public Const ES_MULTILINE = &H4&
Public Const ES_UPPERCASE = &H8&
Public Const ES_LOWERCASE = &H10&
Public Const ES_PASSWORD = &H20&
Public Const ES_AUTOVSCROLL = &H40&
Public Const ES_AUTOHSCROLL = &H80&
Public Const ES_NOHIDESEL = &H100&
Public Const ES_OEMCONVERT = &H400&
Public Const ES_READONLY = &H800&
Public Const ES_WANTRETURN = &H1000&
'
' Listbox Styles.
'
Public Const LBS_NOTIFY = &H1&
Public Const LBS_SORT = &H2&
Public Const LBS_NOREDRAW = &H4&
Public Const LBS_MULTIPLESEL = &H8&
Public Const LBS_OWNERDRAWFIXED = &H10&
Public Const LBS_OWNERDRAWVARIABLE = &H20&
Public Const LBS_HASSTRINGS = &H40&
Public Const LBS_USETABSTOPS = &H80&
Public Const LBS_NOINTEGRALHEIGHT = &H100&
Public Const LBS_MULTICOLUMN = &H200&
Public Const LBS_WANTKEYBOARDINPUT = &H400&
Public Const LBS_EXTENDEDSEL = &H800&
Public Const LBS_DISABLENOSCROLL = &H1000&
Public Const LBS_NODATA = &H2000&
Public Const LBS_STANDARD = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)
'
' TabStrip styles.
'
Public Const TCS_SCROLLOPPOSITE = &H1
Public Const TCS_BOTTOM = &H2
Public Const TCS_RIGHT = &H2
Public Const TCS_MULTISELECT = &H4
Public Const TCS_FLATBUTTONS = &H8
Public Const TCS_FORCEICONLEFT = &H10
Public Const TCS_FORCELABELLEFT = &H20
Public Const TCS_HOTTRACK = &H40
Public Const TCS_VERTICAL = &H80
Public Const TCS_TABS = &H0
Public Const TCS_BUTTONS = &H100
Public Const TCS_SINGLELINE = &H0
Public Const TCS_MULTILINE = &H200
Public Const TCS_RIGHTJUSTIFY = &H0
Public Const TCS_FIXEDWIDTH = &H400
Public Const TCS_RAGGEDRIGHT = &H800
Public Const TCS_FOCUSONBUTTONDOWN = &H1000
Public Const TCS_OWNERDRAWFIXED = &H2000
Public Const TCS_TOOLTIPS = &H4000
Public Const TCS_FOCUSNEVER = &H8000
Public Const TCS_EX_FLATSEPARATORS = &H1
Public Const TCS_EX_REGISTERDROP = &H2
'
' ProgressBar styles.
'
Public Const PBS_SMOOTH = 1
Public Const PBS_VERTICAL = 4
'
' Owner draw control types.
'
Public Const ODT_MENU = 1
Public Const ODT_LISTBOX = 2
Public Const ODT_COMBOBOX = 3
Public Const ODT_BUTTON = 4
Public Const ODT_STATIC = 5
Public Const ODT_HEADER = 100
Public Const ODT_TAB = 101
Public Const ODT_LISTVIEW = 102
'
' Owner draw actions.
'
Public Const ODA_DRAWENTIRE = &H1
Public Const ODA_SELECT = &H2
Public Const ODA_FOCUS = &H4
'
' Owner draw state.
'
Public Const ODS_SELECTED = &H1
Public Const ODS_GRAYED = &H2
Public Const ODS_DISABLED = &H4
Public Const ODS_CHECKED = &H8
Public Const ODS_FOCUS = &H10
Public Const ODS_DEFAULT = &H20
Public Const ODS_COMBOBOXEDIT = &H1000
Public Const ODS_HOTLIGHT = &H40
Public Const ODS_INACTIVE = &H80


Public Sub Main()
With App
    '
    ' Create a hook to monitor messages sent to window procedures.
    ' The system calls fAppHook before passing the message to the
    ' receiving window procedure.
    '
    ' Then in fAppHook we can modify the controls to our liking.
    '
    mlHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf fAppHook, .hInstance, .ThreadID)
    '
    ' Show the form.
    '
    Form1.Show
    '
    ' Remove the hook.
    '
    Call UnhookWindowsHookEx(mlHook)
End With
End Sub

Public Function fAppHook(ByVal lHookID As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

Static bCombo As Boolean
Dim CWP       As CWPSTRUCT
Dim k         As Long
Dim sClass    As String
'
' This hook procedure is a callback function invoked by the
' SetWindowsHookEx API issued in Sub Main.  The system calls
' this function whenever a message is sent to this application.
' Before passing the message to the destination window procedure,
' the system passes the message to this procedure. This procedure
' can examine the message but cannot modify it.
'
' The lParam parameter is the address of a CWPSTRUCT structure
' that contains details about the message that was sent.  Using
' the address, the message data is copied into a local copy.
'
Call CopyMemory(CWP, ByVal lParam, Len(CWP))
'
' Now that we have the message structure, see what message was sent.
'
Select Case CWP.message
    '
    ' We are only interested in the Create message.  When this
    ' message is sent we want to modify how the control is
    ' created.
    '
    Case WM_CREATE
        mlSetStyle = 0
        '
        ' Get the name of the class the window belongs to.
        '
        sClass = Space$(128)
        k = GetClassName(CWP.hwnd, ByVal sClass, 128)
        sClass = Left$(sClass, k)
        '
        ' See if the class matches that for one of our
        ' controls.  The best way to determine the class
        ' name is by uncommenting this debug command or
        ' by using Spy++ that comes with VB enterprise.
        '
        ' NOTE:
        '   To use the VB6 version of Microsoft Common Controls,
        '   change the class names to those indicated below.
        '
        'Debug.Print sClass
        Select Case sClass
            Case "ComboLBox"
                mlSetStyle = GetWindowLong(CWP.hwnd, GWL_STYLE)
                If bCombo Then
                    mlSetStyle = mlSetStyle Or LBS_USETABSTOPS
                Else
                    mlSetStyle = mlSetStyle Or LBS_OWNERDRAWFIXED
                    bCombo = True
                End If
            Case "ProgressBarWndClass" 'VB6 = "ProgressBar20WndClass"
                mlSetStyle = GetWindowLong(CWP.hwnd, GWL_STYLE)
                mlSetStyle = mlSetStyle Or PBS_SMOOTH
            Case "TabStripWndClass"    'VB6 = "TabStrip20WndClass"
                mlSetStyle = GetWindowLong(CWP.hwnd, GWL_STYLE)
                mlSetStyle = mlSetStyle Or TCS_OWNERDRAWFIXED
            Case "ThunderCheckBox"
                mlSetStyle = GetWindowLong(CWP.hwnd, GWL_STYLE)
                mlSetStyle = mlSetStyle _
                        And (Not BS_AUTOCHECKBOX) _
                        And (Not BS_AUTO3STATE) _
                         Or BS_CHECKBOX
            Case "ThunderListBox"
                mlSetStyle = GetWindowLong(CWP.hwnd, GWL_STYLE)
                mlSetStyle = mlSetStyle _
                        Or LBS_SORT _
                        Or LBS_OWNERDRAWVARIABLE _
                        Or LBS_HASSTRINGS
        End Select
        '
        ' Subclass the control by setting a new address for
        ' its associated window procedure.  Now when a message is
        ' sent to the control, our callback procedure will be called.
        '
        ' The SetWindowLong returns the address of the original
        ' window procedure.
        '
        If mlSetStyle Then
            mlHookWndProc = SetWindowLong(CWP.hwnd, GWL_WNDPROC, AddressOf fSetStyle)
        End If
End Select
'
' Pass the message information to the original window procedure.
'
fAppHook = CallNextHookEx(mlHook, lHookID, wParam, ByVal lParam)
End Function

Public Function fSetStyle(ByVal hwnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Dim C As CREATESTRUCT
'
' This callback routine is called by Windows whenever a message
' is sent to the control indicated by hwnd. We are only interested
' in the create message.
'
Select Case Msg
    Case WM_CREATE
        '
        ' When a create message is sent, the lParam parameter has the
        ' address of a CreateStruct structure containing style
        ' information for the control being created.  This structure
        ' is copied locally, modified and then copied back so that the
        ' control is created with our desired style.
        '
        Call CopyMemory(C, ByVal lParam, Len(C))
        C.style = mlSetStyle
        Call CopyMemory(ByVal lParam, C, Len(C))
        '
        ' Set the new style.
        '
        Call SetWindowLong(hwnd, GWL_STYLE, C.style)
        '
        ' Unsub-class the control by assigning it the address
        ' of its original window procedure.
        '
        Call SetWindowLong(hwnd, GWL_WNDPROC, mlHookWndProc)
End Select
'
' Call the original window procedure.
'
fSetStyle = CallWindowProc(mlHookWndProc, hwnd, Msg, wParam, lParam)
End Function

Public Function fAppWndProc(ByVal hwnd As Long, _
    ByVal Msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Dim hdc     As Long
Dim lRet    As Long
Dim hBmp    As Long
Dim lFont   As Long
Dim sString As String
Dim tSize   As SIZE
Dim hPic    As StdPicture
Dim DIS     As DRAWITEMSTRUCT
'
' This callback function is called by the SetWindowLong API
' issued in the form load event.  The system calls this function
' whenever a message is sent to the parent of the one of our sub-
' classed controls.
'
If hwnd <> mhWndLast Or mlWndProc = 0 Then
    mlWndProc = Val(GetSetting("OwnerDraw", CStr(hwnd), "WndProcs"))
    mhWndLast = hwnd
End If

Select Case Msg
    Case WM_DRAWITEM
        '
        ' When a draw item message is received, copy the
        ' accompanying Draw Item Structure to local storage.
        '
        Call CopyMemory(DIS, ByVal lParam, Len(DIS))
        '
        ' See if the item to be drawn is our TabStrip, ListBox
        ' or ComboBox.
        '
        Select Case DIS.CtlType
            Case ODT_COMBOBOX
                '
                ' Drawing the ComboBox.
                '
                sString = Space$(24)
                With DIS
                    '
                    ' Get the text, indicated by DrawItemStruct.iItemID,
                    ' from the list portion of the combobox.
                    '
                    lRet = SendMessage(.hwndItem, CB_GETLBTEXT, .itemID, ByVal sString)
                    sString = Left$(sString, lRet)
                    '
                    ' Add the picture to the list portion of the combobox.
                    '
                    ' Get the associated ItemData value which is the index in
                    ' the resource file of the bitmap to display for that element.
                    '
                    ' Create a memory device context.
                    '
                    hBmp = LoadBitmap(0, ByVal .itemData)
                    hdc = CreateCompatibleDC(.hdc)
                    '
                    ' Select the bitmap into the new device context.
                    '
                    Call SelectObject(hdc, hBmp)
                    '
                    ' Set the color of the text in the list portion of
                    ' the combobox based on whether the item is selected
                    ' or not.
                    '
                    If .itemState And ODS_SELECTED Then
                        Call SetBkColor(.hdc, GetSysColor(COLOR_HIGHLIGHT))
                        Call SetTextColor(.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT))
                    Else
                        Call SetBkColor(.hdc, GetSysColor(COLOR_WINDOW))
                        Call SetTextColor(.hdc, GetSysColor(COLOR_WINDOWTEXT))
                    End If
                End With
                '
                ' Draw the bitmap.
                ' TextOut function writes a character string at the specified location,
                ' using the currently selected font, background color, and text color.
                '
                With DIS.rcItem
                    Call BitBlt(DIS.hdc, .Left, .Top, .Right - .Left, .Bottom - .Top, hdc, 0, 0, vbSrcCopy)
                    Call TextOut(DIS.hdc, .Left + GetSystemMetrics(SM_CXMENUSIZE), .Top, sString, Len(sString))
                End With
                Call DeleteDC(hdc)
                Call DeleteObject(hBmp)
            
            Case ODT_LISTBOX
                '
                ' Drawing the ListBox.
                '
                With DIS
                    '
                    ' Get the text, indicated by DrawItemStruct.iItemID,
                    ' from the listbox.
                    '
                    sString = Space$(128)
                    lRet = SendMessage(.hwndItem, LB_GETTEXT, .itemID, ByVal sString)
                    sString = Left$(sString, lRet)
                    '
                    ' Create a logical font based on the font specified
                    ' in the listbox.
                    '
                    lFont = CreateFont(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, sString)
                    '
                    ' Select the font into the new device context.
                    ' Compute the width and height of a string of text (font name).
                    ' Set the height of the listbox item accordingly.
                    '
                    Call SelectObject(.hdc, lFont)
                    Call GetTextExtentPoint32(.hdc, sString, Len(sString), tSize)
                    Call SendMessage(.hwndItem, LB_SETITEMHEIGHT, .itemID, ByVal tSize.cy)
                    '
                    ' Set the color of the text of the list item
                    ' based on whether the item is selected or not.
                    '
                    If .itemState And ODS_SELECTED Then
                        Call SetBkColor(.hdc, GetSysColor(COLOR_HIGHLIGHT))
                        Call SetTextColor(.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT))
                    Else
                        Call SetBkColor(.hdc, GetSysColor(COLOR_WINDOW))
                        Call SetTextColor(.hdc, GetSysColor(COLOR_WINDOWTEXT))
                    End If
                End With
                '
                ' Draw the text in the listbox.
                '
                With DIS.rcItem
                    Call TextOut(DIS.hdc, .Left, .Top, sString, Len(sString))
                End With
                Call DeleteObject(lFont)
            
            Case ODT_TAB
                '
                ' Drawing the TabStrip.
                '
                ' Select the resource file's bitmap into a
                ' device context.
                '
                hdc = CreateCompatibleDC(DIS.hdc)
                If DIS.itemID = 0 Then
                    Set hPic = LoadResPicture(101, 0)
                Else
                    Set hPic = LoadResPicture(102, 0)
                End If
                Call SelectObject(hdc, hPic)
                '
                ' Copy the bitmap to the tab.
                '
                With DIS.rcItem
                    Call BitBlt(DIS.hdc, .Left, .Top, .Right - .Left, .Bottom - .Top, hdc, 0, 0, vbSrcCopy)
                End With
                Set hPic = Nothing
                Call DeleteDC(hdc)
        End Select
        
        fAppWndProc = True
        Exit Function
    
    Case WM_PARENTNOTIFY
        '
        ' This message is sent to the parent of a child window when the child
        ' window is created or destroyed, or when the user clicks a mouse button
        ' while the cursor is over the child window. When the child window is
        ' being created, the system sends WM_PARENTNOTIFY just before the
        ' CreateWindow function that creates the window returns. When the child
        ' window is being destroyed, the system sends the message before any
        ' processing to destroy the window takes place.
        '
        If (wParam And &HFF) = WM_DESTROY Then GoTo DestroyIt
    
    Case WM_DESTROY
DestroyIt:
        '
        ' When the owner drawn control is destroyed, unsubclass it first.
        '
        Call DeleteSetting("OwnerDraw", CStr(hwnd))
        Call SetWindowLong(hwnd, GWL_WNDPROC, mlWndProc)
End Select
'
' Pass the message along to the original window procedure
' associated with the control's parent.
'
fAppWndProc = CallWindowProc(mlWndProc, hwnd, Msg, wParam, lParam)
End Function


