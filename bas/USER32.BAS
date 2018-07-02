Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
Declare Function AdjustWindowRect Lib "user32" Alias "AdjustWindowRectA" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Declare Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
Declare Function AnyPopup Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function ArrangeIconicWindows Lib "user32" (ByVal hWnd As Long) As Long
Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function BroadcastSystemMessage Lib "user32" (ByVal dw As Long, pdw As Long, ByVal un As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function CallMsgFilter Lib "user32" Alias "CallMsgFilterA" (lpMsg As msg, ByVal nCode As Long) As Long
Declare Function CharPrev Lib "user32" Alias "CharPrevA" (ByVal lpszStart As String, ByVal lpszCurrent As String) As String
Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Declare Function CharUpperBuff Lib "user32" Alias "CharUpperBuffA" (ByVal lpsz As String, ByVal cchLength As Long) As Long
Declare Function CheckDlgButton Lib "user32" Alias "CheckDLGButtonA" (ByVal hDlg As Long, ByVal nIDButton As Long, ByVal wCheck As Long) As Long
Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Boolean
Declare Function CheckRadioButton Lib "user32" Alias "CheckRadioButtonA" (ByVal hDlg As Long, ByVal nIDFirstButton As Long, ByVal nIDLastButton As Long, ByVal nIDCheckButton As Long) As Long
Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hWnd As Long, ByVal pt As POINTAPI, ByVal un As Long) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function CloseDesktop Lib "user32" (ByVal hDesktop As Long) As Boolean
Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function CloseWindowStation Lib "user32" (ByVal hWinSta As Long) As Boolean
Declare Function CopyAcceleratorTable Lib "user32" Alias "CopyAcceleratorTableA" (ByVal hAccelSrc As Long, lpAccelDst As ACCEL, ByVal cAccelEntries As Long) As Long
Declare Function CopyCursor Lib "user32" (ByVal hcur As Long) As Long
Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Declare Function CountClipboardFormats Lib "user32" () As Long
Declare Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableA" (lpaccl As ACCEL, ByVal cEntries As Long) As Long
Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateIcon Lib "user32" (ByVal hInstance As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Byte, ByVal nBitsPixel As Byte, lpANDbits As Byte, lpXORbits As Byte) As Long
Declare Function CreateIconFromResource Lib "user32" (presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As Long
Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Declare Function CreateMDIWindow Lib "user32" Alias "CreateMDIWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hInstance As Long, ByVal lParam As Long) As Long
Declare Function CreateMenu Lib "user32" () As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function CreateWindow Lib "user32" Alias "CreateWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function CreateWindowEx Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DdeAbandonTransaction Lib "user32" (ByVal idInst As Long, ByVal hConv As Long, ByVal idTransaction As Long) As Long
Declare Function DdeAccessData Lib "user32" Alias "DdeAccessDataA" (ByVal hData As Long, pcbDataSize As Long) As Long
Declare Function DdeAddData Lib "user32" Alias "DdeAddDataA" (ByVal hData As Long, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long) As Long
Declare Function DdeClientTransaction Lib "user32" (pData As Byte, ByVal cbData As Long, ByVal hConv As Long, ByVal hszItem As Long, ByVal wFmt As Long, ByVal wType As Long, ByVal dwTimeout As Long, pdwResult As Long) As Long
Declare Function DdeCmpStringHandles Lib "user32" (ByVal hsz1 As Long, ByVal hsz2 As Long) As Long
Declare Function DdeConnect Lib "user32" (ByVal idInst As Long, ByVal hszService As Long, ByVal hszTopic As Long, pCC As CONVCONTEXT) As Long
Declare Function DdeConnectList Lib "user32" (ByVal idInst As Long, ByVal hszService As Long, ByVal hszTopic As Long, ByVal hConvList As Long, pCC As CONVCONTEXT) As Long
Declare Function DdeCreateDataHandle Lib "user32" (ByVal idInst As Long, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long, ByVal hszItem As Long, ByVal wFmt As Long, ByVal afCmd As Long) As Long
Declare Function DdeCreateStringHandle Lib "user32" Alias "DdeCreateStringHandleA" (ByVal idInst As Long, ByVal psz As String, ByVal iCodePage As Long) As Long
Declare Function DdeDisconnect Lib "user32" (ByVal hConv As Long) As Long
Declare Function DdeDisconnectList Lib "user32" (ByVal hConvList As Long) As Long
Declare Function DdeEnableCallback Lib "user32" (ByVal idInst As Long, ByVal hConv As Long, ByVal wCmd As Long) As Long
Declare Function DdeFreeDataHandle Lib "user32" (ByVal hData As Long) As Long
Declare Function DdeFreeStringHandle Lib "user32" (ByVal idInst As Long, ByVal hsz As Long) As Long
Declare Function DdeGetData Lib "user32" Alias "DdeGetDataA" (ByVal hData As Long, pDst As Byte, ByVal cbMax As Long, ByVal cbOff As Long) As Long
Declare Function DdeGetLastError Lib "user32" (ByVal idInst As Long) As Long
Declare Function DdeImpersonateClient Lib "user32" (ByVal hConv As Long) As Long
Declare Function DdeInitialize Lib "user32" Alias "DdeInitializeA" (pidInst As Long, ByVal pfnCallback As Long, ByVal afCmd As Long, ByVal ulRes As Long) As Integer
Declare Function DdeKeepStringHandle Lib "user32" (ByVal idInst As Long, ByVal hsz As Long) As Long
Declare Function DdeNameService Lib "user32" (ByVal idInst As Long, ByVal hsz1 As Long, ByVal hsz2 As Long, ByVal afCmd As Long) As Long
Declare Function DefDlgProc Lib "user32" Alias "DefDlgProcA" (ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function DefFrameProc Lib "user32" Alias "DefFrameProcA" (ByVal hWnd As Long, ByVal hWndMDIClient As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyAcceleratorTable Lib "user32" (ByVal haccel As Long) As Long
Declare Function DestroyCaret Lib "user32" () As Long
Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function DialogBoxIndirectParam Lib "user32" Alias "DialogBoxIndirectParamA" (ByVal hInstance As Long, hDialogTemplate As DLGTEMPLATE, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Declare Function DlgDirList Lib "user32" Alias "DlgDirListA" (ByVal hDlg As Long, ByVal lpPathSpec As String, ByVal nIDListBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long
Declare Function DlgDirListComboBox Lib "user32" Alias "DlgDirListComboBoxA" (ByVal hDlg As Long, ByVal lpPathSpec As String, ByVal nIDComboBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long
Declare Function DlgDirSelectComboBoxEx Lib "user32" Alias "DlgDirSelectComboBoxExA" (ByVal hWndDlg As Long, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idComboBox As Long) As Long
Declare Function DlgDirSelectEx Lib "user32" Alias "DlgDirSelectExA" (ByVal hWndDlg As Long, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idListBox As Long) As Long
Declare Function DragDetect Lib "user32" (ByVal hWnd As Long, ByVal pt As POINTAPI) As Boolean
Declare Function DragObject Lib "user32" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal un As Long, ByVal dw As Long, ByVal hCursor As Long) As Long
Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Boolean
Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, pcRect As RECT, ByVal un As Long) As Boolean
Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Boolean
Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Boolean
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Declare Function EndDialog Lib "user32" (ByVal hDlg As Long, ByVal nResult As Long) As Long
Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Declare Function EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hWinSta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Declare Function EnumProps Lib "user32" Alias "EnumPropsA" (ByVal hWnd As Long, ByVal lpEnumFunc As Long) As Long
Declare Function EnumPropsEx Lib "user32" Alias "EnumPropsExA" (ByVal hWnd As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Boolean
Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Declare Function ExcludeUpdateRgn Lib "user32" (ByVal hDC As Long, ByVal hWnd As Long) As Long
Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function FreeDDElParam Lib "user32" (ByVal msg As Long, ByVal lParam As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetCapture Lib "user32" () As Long
Declare Function GetCaretBlinkTime Lib "user32" () As Long
Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long
Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetClassWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As Long
Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function GetClipboardOwner Lib "user32" () As Long
Declare Function GetClipboardViewer Lib "user32" () As Long
Declare Function GetClipCursor Lib "user32" (lprc As RECT) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDCEx Lib "user32" (ByVal hWnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDialogBaseUnits Lib "user32" () As Long
Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Declare Function GetDlgItemInt Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpTranslated As Long, ByVal bSigned As Long) As Long
Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function GetDoubleClickTime Lib "user32" () As Long
Declare Function GetFocus Lib "user32" () As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Declare Function GetInputState Lib "user32" () As Long
Declare Function GetKBCodePage Lib "user32" () As Long
Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function GetLastActivePopup Lib "user32" (ByVal hwndOwnder As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Boolean
Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Declare Function GetMessageExtraInfo Lib "user32" () As Long
Declare Function GetMessagePos Lib "user32" () As Long
Declare Function GetMessageTime Lib "user32" () As Long
Declare Function GetOpenClipboardWindow Lib "user32" () As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetPriorityClipboardFormat Lib "user32" (lpPriorityList As Long, ByVal nCount As Long) As Long
Declare Function GetProcessWindowStation Lib "user32" () As Long
Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long
Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Boolean
Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hDC As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long
Declare Function GetThreadDesktop Lib "user32" (ByVal dwThread As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function GetUpdateRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal fErase As Long) As Long
Declare Function GetUserObjectInformation Lib "user32" Alias "GetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long, lpnLengthNeeded As Long) As Boolean
Declare Function GetUserObjectSecurity Lib "user32" (ByVal hObj As Long, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowContextHelpId Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Declare Function HiliteMenuItem Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
Declare Function ImpersonateDdeClientWindow Lib "user32" (ByVal hWndClient As Long, ByVal hWndServer As Long) As Long
Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function InSendMessage Lib "user32" () As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByVal lpcMenuItemInfo As MENUITEMINFO) As Boolean
Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function InvalidateRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
Declare Function InvertRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hWnd As Long) As Long
Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageA" (ByVal hDlg As Long, lpMsg As msg) As Long
Declare Function IsDlgButtonChecked Lib "user32" (ByVal hDlg As Long, ByVal nIDButton As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Declare Function LoadAccelerators Lib "user32" Alias "LoadAcceleratorsA" (ByVal hInstance As Long, ByVal lpTableName As String) As Long
Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As String) As Long
Declare Function LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (ByVal lpMenuTemplate As Long) As Long
Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function LookupIconIdFromDirectory Lib "user32" (presbits As Byte, ByVal fIcon As Long) As Long
Declare Function LookupIconIdFromDirectoryEx Lib "user32" (presbits As Byte, ByVal fIcon As Boolean, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal flags As Long) As Long
Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Declare Function MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExA" (ByVal uCode As Long, ByVal uMapType As Long, ByVal dwhkl As Long) As Long
Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Declare Function MenuItemFromPoint Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long
Declare Function MessageBeep Lib "user32" Alias "BeepA" (ByVal wType As Long) As Long
Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectA" (lpMsgBoxParams As MSGBOXPARAMS) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Long) As Long
Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Declare Function OpenDesktop Lib "user32" Alias "OpenDesktopA" (ByVal lpszDesktop As String, ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Declare Function OpenInputDesktop Lib "user32" (ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenWindowStation Lib "user32" Alias "OpenWindowStationA" (ByVal lpszWinSta As String, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Declare Function PackDDElParam Lib "user32" (ByVal msg As Long, ByVal uiLo As Long, ByVal uiHi As Long) As Long
Declare Function PaintDesktop Lib "user32" (ByVal hDC As Long) As Boolean
Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Declare Function PostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function RegisterClass Lib "user32" (Class As WNDCLASS) As Long
Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function ReplyMessage Lib "user32" (ByVal lReply As Long) As Long
Declare Function ReuseDDElParam Lib "user32" (ByVal lParam As Long, ByVal msgIn As Long, ByVal msgOut As Long, ByVal uiLo As Long, ByVal uiHi As Long) As Long
Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Declare Function SendMessageCallback Lib "user32" Alias "SendMessageCallbackA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lpResultCallBack As Long, ByVal dwData As Long) As Long
Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
Declare Function SetCaretPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetClassWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Declare Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Sub SetDebugErrorLevel Lib "user32" (ByVal dwLevel As Long)
Declare Function SetDlgItemInt Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wValue As Long, ByVal bSigned As Long) As Long
Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Declare Function SetMenuContextHelpId Lib "user32" (ByVal hMenu As Long, ByVal dw As Long) As Boolean
Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Boolean
Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Boolean
Declare Function SetMessageExtraInfo Lib "user32" (ByVal lParam As Long) As Long
Declare Function SetMessageQueue Lib "user32" (ByVal cMessagesMax As Long) As Boolean
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetProcessWindowStation Lib "user32" (ByVal hWinSta As Long) As Boolean
Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Declare Function SetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Boolean
Declare Function SetThreadDesktop Lib "user32" (ByVal hDesktop As Long) As Boolean
Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function SetUserObjectInformation Lib "user32" Alias "SetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long) As Boolean
Declare Function SetUserObjectSecurity Lib "user32" (ByVal hObj As Long, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR) As Long
Declare Function SetWindowContextHelpId Lib "user32" (ByVal hWnd As Long, ByVal dw As Long) As Boolean
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Declare Function SubtractRect Lib "user32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Declare Function SwitchDesktop Lib "user32" (ByVal hDesktop As Long) As Boolean
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long
Declare Function TileWindows Lib "user32" (ByVal hWndParent As Long, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpKids As Long) As Integer
Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Declare Function ToAsciiEx Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpKeyState As Byte, lpChar As Integer, ByVal uFlags As Long, ByVal dwhkl As Long) As Long
Declare Function ToUnicode Lib "user32" (ByVal wVirtKey As Long, ByVal wScanCode As Long, lpKeyState As Byte, ByVal pwszBuff As String, ByVal cchBuff As Long, ByVal wFlags As Long) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long
Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As Long, lpTPMParams As TPMPARAMS) As Boolean
Declare Function TranslateAccelerator Lib "user32" Alias "TranslateAcceleratorA" (ByVal hWnd As Long, ByVal hAccTable As Long, lpMsg As msg) As Long
Declare Function TranslateMDISysAccel Lib "user32" (ByVal hWndClient As Long, lpMsg As msg) As Long
Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Declare Function UnhookWindowsHook Lib "user32" (ByVal nCode As Long, ByVal pfnFilterProc As Long) As Boolean
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function UnionRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Declare Function UnloadKeyboardLayout Lib "user32" (ByVal HKL As Long) As Long
Declare Function UnpackDDElParam Lib "user32" (ByVal msg As Long, ByVal lParam As Long, puiLo As Long, puiHi As Long) As Long
Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function ValidateRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Declare Function WaitMessage Lib "user32" () As Long
Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
