'|¯\|¯\¯|¯|¯|¯)¯|(¯(
'|__\_\_|___/_)_| )_)
'This BAS file was part compiled from older BAS files but
'mostly written by Anubis for AOL Add-On Programmming.
'Anubis nor Core is responsible for your actions with this
'BAS file.
'I strongly recommend you read all the subs and functions
'before you try to incorporate it into your programming
'project.  There are many many shortcuts in this BAS and
'should make programming a synch if you take the time to
'read everything.

'VBMsg Chat code
'Message$ = agGetStringFromLPSTR$(lparam)
'SN$ = Mid$(Message$, 3, InStr(Message$, ":") - 3)
'TXT$ = Mid$(Message$, InStr(Message$, ":") + 2)

'M-Chat Scroll Code [in VBMsg]
'(text1.text is the main VBMsg textbox)
'If Len(Text1.Text) >= 32000 Then Text1.Text = ""
'Text1.SelStart = Len(Text1.Text)
'Text1.SelText = Message$

'AOL Private Room URL: aol://2719:2-2-
'AOL Conference Room URL: aol://2719:3-37- (PC Devolment anyways)

'MoveWindow
'MoveWindow TargethWnd%, LeftCoord%, TopCoord%, hWndWidth%, hWndTop%
'hWnd = yer target Window handle

'Various
Type ConvertPointAPI
    xy As Long
    End Type
Type POINTAPI
    x As Integer
    y As Integer
End Type
Global Const MOUSE_MOVE = &HF012
Declare Function ExitWindows Lib "User" (ByVal RestartCode As Long, ByVal DOSReturnCode As Integer) As Integer
Declare Function GetKeyboardType Lib "Keyboard" (ByVal nTypeFlag As Integer) As Integer
Declare Function ShowCursor Lib "User" (ByVal bShow As Integer) As Integer

'API Spy
Declare Function WindowFromPoint Lib "User" (ByVal ptScreen As Any) As Integer
Declare Function GetModuleFileName Lib "Kernel" (ByVal hModule As Integer, ByVal lpFilename As String, ByVal nSize As Integer) As Integer
Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function GetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Long
Declare Function GetParent Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Declare Function getwindowtext Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
Declare Function GetActiveWindow% Lib "User" ()


'Speech Engine
Declare Function OpenSpeech Lib "fb_spch.dll" (ByVal hWnd%, ByVal mode%, ByVal voiceType&) As Long
Declare Function CloseSpeech Lib "fb_spch.dll" (ByVal lpSCB&) As Integer
Declare Function Say Lib "fb_spch.dll" (ByVal lpSCB&, ByVal phrase$) As Integer
Global lpSCB As Long
Global Const MCIERR_INVALID_DEVICE_ID = 30257
Global Const MCIERR_DEVICE_OPEN = 30263
Global Const MCIERR_CANNOT_LOAD_DRIVER = 30266
Global Const MCIERR_UNSUPPORTED_FUNCTION = 30274
Global Const MCIERR_INVALID_FILE = 30304
Global Const MCI_MODE_NOT_OPEN = 524
Global Const MCI_MODE_PLAY = 526
Global Const MCI_FORMAT_MILLISECONDS = 0
Global Const MCI_FORMAT_TMSF = 10

'311 [by Tartan]
Declare Function AOLGetList Lib "311.Dll" (ByVal p1%, ByVal p2$) As Integer
Declare Function AOLGetcombo% Lib "311.Dll" (ByVal index%, ByVal Buf$)

'MIDI
Declare Function mciExecute Lib "winmm.dll" Alias "mciExecute" (ByVal lpstrCommand As String) As Long

'AVI
Declare Function mciSendString Lib "mmsystem" (ByVal lpstrCommand$, ByVal lpstrReturnStr As Any, ByVal wReturnLen%, ByVal hCallBack%) As Long

'User
Declare Function DeleteMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function RemoveMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function GetFocus Lib "User" () As Integer
Declare Function getwindowtext Lib "User" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Option Compare Text
Declare Function ShowWindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Const c016E = 16
Const c014A = 12
Const c00AE = 2
Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const WM_USER = &H400
Global Const SW_NORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_MAXIMIZE = 3
Global Const SW_SHOWNOACTIVATE = 4
Global Const SW_SHOW = 5
Global Const SW_MINIMIZE = 6
Global Const SW_SHOWMINNOACTIVE = 7
Global Const SW_SHOWNA = 8
Global Const SW_RESTORE = 9
Global Const BM_SETCHECK = WM_USER + 1
Global Const LB_SETCURSEL = (WM_USER + 7)
Global Const LB_GETTEXT = (WM_USER + 10)
Global Const LB_GETTEXTLEN = (WM_USER + 11)
Global Const LB_ADDSTRING = WM_USER + 2
Global Const LB_DELETESTRING = (WM_USER + 3)
Type MODEL
  usVersion         As Integer
  fl                As Long
  pctlproc          As Long
  fsClassStyle      As Integer
  flWndStyle        As Long
  cbCtlExtra        As Integer
  idBmpPalette      As Integer
  npszDefCtlName    As Integer
  npszClassName     As Integer
  npszParentClassName As Integer
  npproplist        As Integer
  npeventlist       As Integer
  nDefProp          As String * 1
  nDefEvent         As String * 1
  nValueProp        As String * 1
  usCtlVersion      As Integer
End Type
Type Rect
  Left As Integer
  Top As Integer
  Right As Integer
  Bottom As Integer
End Type
Type HelpWinInfo
  wStructSize As Integer
  x As Integer
  y As Integer
  dx As Integer
  dy As Integer
  wMax As Integer
  rgChMember As String * 2
End Type

'                   API Subs and Functions
'                   ----------------------

'Subs and Functions for "User"
Declare Sub RedrawWindow Lib "User" (ByVal hWnd As Integer, lprcUpdate As Rect, ByVal hrgnUpdate As Integer, ByVal fuRedraw As Integer) 'As Integer
Declare Sub closewindow Lib "User" (ByVal hWnd As Integer)
Declare Sub UpdateWindow Lib "User" (ByVal hWnd As Integer)
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, lpRect As Rect)
Declare Sub SetWindowPos Lib "User" (ByVal H%, ByVal hb%, ByVal x%, ByVal y%, ByVal cX%, ByVal cY%, ByVal f%)
Declare Sub DrawMenuBar Lib "User" (ByVal hWnd As Integer)
Declare Sub GetScrollRange Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer, Lpminpos As Integer, lpmaxpos As Integer)
Declare Sub SetCursorPos Lib "User" (ByVal x As Integer, ByVal y As Integer)
Declare Sub UpdateWindow Lib "User" (ByVal hWnd%)
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Function GetNextWindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource%) As Integer
Declare Function GetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%) As Integer
Declare Function SetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%, ByVal wNewWord%) As Integer
Declare Function GetFocus% Lib "User" ()
Declare Function SetFocusAPI% Lib "User" Alias "SetFocus" (ByVal hWnd As Integer)
Declare Function GetWindow% Lib "User" (ByVal hWnd%, ByVal wCmd%)
Declare Function FindWindow% Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any)
Declare Function FindWindowByNum% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function FindWindowByString% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function ExitWindow% Lib "User" (ByVal dwReturnCode&, ByVal wReserved%)
Declare Function GetParent% Lib "User" (ByVal hWnd As Integer)
Declare Function SetParent% Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer)
Declare Function GetMessage% Lib "User" (lpMsg As String, ByVal hWnd As Integer, ByVal wMsgFilterMin As Integer, ByVal wMsgFilterMax As Integer)
Declare Function GetMenuString% Lib "User" (ByVal hMenu%, ByVal wIDItem%, ByVal lpString$, ByVal nMaxCount%, ByVal wFlag%)
Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lparam As Any) As Long
Declare Function sendmessagebystring& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lparam$)
Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lparam&)
Declare Function CreateMenu% Lib "User" ()
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function AppendMenuByString% Lib "User" Alias "AppendMenu" (ByVal hMenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem$)
Declare Function InsertMenu% Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any)
Declare Function WinHelp% Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any)
Declare Function WinHelpByString% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$)
Declare Function WinHelpByNum% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&)
Declare Function GetWindow% Lib "User" (ByVal hWnd%, ByVal wCmd%)
Declare Function getwindowtext% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer)
Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function setwindowtext% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Function GetActiveWindow% Lib "User" ()
Declare Function setactivewindow% Lib "User" (ByVal hWnd%)
Declare Function GetSysModalWindow% Lib "User" ()
Declare Function SetSysModalWindow% Lib "User" (ByVal hWnd As Integer)
Declare Function IsWindowVisible% Lib "User" (ByVal hWnd%)
Declare Function getcurrenttime& Lib "User" ()
Declare Function GetScrollPos Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer) As Integer
Declare Function GetCursor% Lib "User" ()
Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function GetNextDlgTabItem Lib "User" (ByVal hDlg As Integer, ByVal hctl As Integer, ByVal bPrevious As Integer) As Integer
Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetTopWindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function ArrangeIconicWindow% Lib "User" (ByVal hWnd%)
Declare Function getmenu% Lib "User" (ByVal hWnd%)
Declare Function GetMenuItemID% Lib "User" (ByVal hMenu%, ByVal nPos%)
Declare Function GetMenuItemCount% Lib "User" (ByVal hMenu%)
Declare Function GetMenuState% Lib "User" (ByVal hMenu%, ByVal wId%, ByVal wFlags%)
Declare Function GetSubMenu% Lib "User" (ByVal hMenu%, ByVal nPos%)
Declare Function GetSystemMetrics Lib "User" (ByVal nIndex%) As Integer
Declare Function GetDesktopWindow Lib "User" () As Integer
Declare Function GetDC Lib "User" (ByVal hWnd%) As Integer
Declare Function ReleaseDC Lib "User" (ByVal hWnd%, ByVal hDC%) As Integer
Declare Function GetWindowDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SwapMouseButton% Lib "User" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "User" (ByVal hwndparent%, ByVal lpenumfunc&, ByVal lparam&)
Declare Function FillRect Lib "User" (ByVal hDC As Integer, lpRect As Rect, ByVal hBrush As Integer) As Integer

'Subs and Functions for "Kernel"
Declare Function lStrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags%) As Long
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFilename$) As Integer

'Subs and Functions for "GDI"
Declare Function GetPixel Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Declare Sub setbkcolor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
Declare Function GetDeviceCaps Lib "GDI" (ByVal hDC%, ByVal nIndex%) As Integer
Declare Function TextOut Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function FloodFill Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal crColor As Long) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer

'Subs and Functions for "MMSystem"
Declare Function SndPlaySound Lib "MMSystem" (ByVal lpWavName$, ByVal FLAGS%) As Integer
'Declare Function mciSendString& Lib "MMSystem" (ByVal Cmd$, ByVal ReturnStr As Any, ByVal returnlen%, ByVal hCallBack%)

'Subs and Functions for "VBWFind.Dll"
Declare Function FindChild% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal Title$)
Declare Function FindChildByTitle% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal Title$)
Declare Function FindChildByClass% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal Title$)

'Subs and Functions for "APIGuide.Dll"
Declare Function agGetStringFromLPSTR$ Lib "APIGUIDE.DLL" (ByVal lpString&)
Declare Sub agCopyData Lib "APIGuide.Dll" (source As Any, dest As Any, ByVal nCount%)
Declare Sub agCopyDataBynum Lib "APIGuide.Dll" Alias "agCopyData" (ByVal source&, ByVal dest&, ByVal nCount%)
Declare Sub agDWordTo2Integers Lib "APIGuide.Dll" (ByVal L&, lw%, lh%)
Declare Sub agOutp Lib "APIGuide.Dll" (ByVal portid%, ByVal outval%)
Declare Sub agOutpw Lib "APIGuide.Dll" (ByVal portid%, ByVal outval%)
Declare Function agGetControlHwnd% Lib "APIGuide.Dll" (hctl As Control)
Declare Function agGetInstance% Lib "APIGuide.Dll" ()
Declare Function agGetAddressForObject& Lib "APIGuide.Dll" (object As Any)
Declare Function agGetAddressForInteger& Lib "APIGuide.Dll" Alias "agGetAddressForObject" (intnum%)
Declare Function agGetAddressForLong& Lib "APIGuide.Dll" Alias "agGetAddressForObject" (intnum&)
Declare Function agGetAddressForLPSTR& Lib "APIGuide.Dll" Alias "agGetAddressForObject" (ByVal lpString$)
Declare Function agGetAddressForVBString& Lib "APIGuide.Dll" (vbstring$)
Declare Function agGetControlName$ Lib "APIGuide.Dll" (ByVal hWnd%)
Declare Function agXPixelsToTwips& Lib "APIGuide.Dll" (ByVal pixels%)
Declare Function agYPixelsToTwips& Lib "APIGuide.Dll" (ByVal pixels%)
Declare Function agXTwipsToPixels% Lib "APIGuide.Dll" (ByVal Twips&)
Declare Function agYTwipsToPixels% Lib "APIGuide.Dll" (ByVal Twips&)
Declare Function agDeviceCapabilities& Lib "APIGuide.Dll" (ByVal hlib%, ByVal lpszDevice$, ByVal lpszPort$, ByVal fwCapability%, ByVal lpszOutput&, ByVal lpdm&)
Declare Function agDeviceMode% Lib "APIGuide.Dll" (ByVal hWnd%, ByVal hModule%, ByVal lpszDevice$, ByVal lpszOutput$)
Declare Function agExtDeviceMode% Lib "APIGuide.Dll" (ByVal hWnd%, ByVal hDriver%, ByVal lpdmOutput&, ByVal lpszDevice$, ByVal lpszPort$, ByVal lpdmInput&, ByVal lpszProfile&, ByVal fwMode%)
Declare Function agInp% Lib "APIGuide.Dll" (ByVal portid%)
Declare Function agInpw% Lib "APIGuide.Dll" (ByVal portid%)
Declare Function agHugeOffset& Lib "APIGuide.Dll" (ByVal addr&, ByVal offset&)
Declare Function agVBGetVersion% Lib "APIGuide.Dll" ()
Declare Function agVBSendControlMsg& Lib "APIGuide.Dll" (ctl As Control, ByVal Msg%, ByVal wp%, ByVal lp&)
Declare Function agVBSetControlFlags& Lib "APIGuide.Dll" (ctl As Control, ByVal mask&, ByVal value&)
Declare Function dwVBSetControlFlags& Lib "APIGuide.Dll" (ctl As Control, ByVal mask&, ByVal value&)


'Subs and Functions for "VBMsg.Vbx"
Declare Sub ptGetTypeFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptCopyTypeToAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptSetControlModel Lib "VBMsg.Vbx" (ctl As Control, lpm As MODEL)
Declare Function ptGetVariableAddress Lib "VBMsg.Vbx" (Var As Any) As Long
Declare Function ptGetTypeAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (Var As Any) As Long
Declare Function ptGetStringAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (ByVal S As String) As Long
Declare Function ptGetLongAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (L As Long) As Long
Declare Function ptGetIntegerAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (i As Integer) As Long
Declare Function ptGetIntegerFromAddress Lib "VBMsg.Vbx" (ByVal i As Long) As Integer
Declare Function ptGetLongFromAddress Lib "VBMsg.Vbx" (ByVal L As Long) As Long
Declare Function ptGetStringFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, ByVal cbBytes As Integer) As String
Declare Function ptMakelParam Lib "VBMsg.Vbx" (ByVal wLow As Integer, ByVal wHigh As Integer) As Long
Declare Function ptLoWord Lib "VBMsg.Vbx" (ByVal lparam As Long) As Integer
Declare Function ptHiWord Lib "VBMsg.Vbx" (ByVal lparam As Long) As Integer
Declare Function ptMakeUShort Lib "VBMsg.Vbx" (ByVal LongVal As Long) As Integer
Declare Function ptConvertUShort Lib "VBMsg.Vbx" (ByVal ushortVal As Integer) As Long
Declare Function ptMessagetoText Lib "VBMsg.Vbx" (ByVal uMsgID As Long, ByVal bFlag As Integer) As String
Declare Function ptRecreateControlHwnd Lib "VBMsg.Vbx" (ctl As Control) As Long
Declare Function ptGetControlModel Lib "VBMsg.Vbx" (ctl As Control, lpm As MODEL) As Long
Declare Function ptGetControlName Lib "VBMsg.Vbx" (ctl As Control) As String

'Subs and Functions for Other DLL's and VBX's
Declare Function VarPtr& Lib "VBRun300.Dll" (Param As Any)
Declare Function vbeNumChildWindow% Lib "VBStr.Dll" (ByVal Win%, ByVal iNum%)
Declare Function EnableWindow Lib "User" (ByVal hWnd As Integer, ByVal aBOOL As Integer) As Integer
Declare Sub MoveWindow Lib "User" (ByVal hWnd As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Integer)

'                   Sound Constants
'          ---------------

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
'                   Global Constants
'          ----------------
'OpenFile() Flags
Global Const OF_READ = &H0
Global Const OF_WRITE = &H1
Global Const OF_READWRITE = &H2
Global Const OF_SHARE_COMPAT = &H0
Global Const OF_SHARE_EXCLUSIVE = &H10
Global Const OF_SHARE_DENY_WRITE = &H20
Global Const OF_SHARE_DENY_READ = &H30
Global Const OF_SHARE_DENY_NONE = &H40
Global Const OF_PARSE = &H100
Global Const OF_DELETE = &H200
Global Const OF_VERIFY = &H400
Global Const OF_SEARCH = &H400
Global Const OF_CANCEL = &H800
Global Const OF_CREATE = &H1000
Global Const OF_PROMPT = &H2000
Global Const OF_EXIST = &H4000
Global Const OF_REOPEN = &H8000
Global Const TF_FORCEDRIVE = &H80

'GetDriveType return values
Global Const DRIVE_REMOVABLE = 2
Global Const DRIVE_FIXED = 3
Global Const DRIVE_REMOTE = 4

'Global Memory Flags
Global Const GMEM_FIXED = &H0
Global Const GMEM_MOVEABLE = &H2
Global Const GMEM_NOCOMPACT = &H10
Global Const GMEM_NODISCARD = &H20
Global Const GMEM_ZEROINIT = &H40
Global Const GMEM_MODIFY = &H80
Global Const GMEM_DISCARDABLE = &H100
Global Const GMEM_NOT_BANKED = &H1000
Global Const GMEM_SHARE = &H2000
Global Const GMEM_DDESHARE = &H2000
Global Const GMEM_NOTIFY = &H4000
Global Const GMEM_LOWER = GMEM_NOT_BANKED
Global Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Global Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

'Flags returned by GlobalFlags (in addition to GMEM_DISCARDABLE)
Global Const GMEM_DISCARDED = &H4000
Global Const GMEM_LOCKCOUNT = &HFF

'Predefined Resource Types
Global Const RT_CURSOR = 1&
Global Const RT_BITMAP = 2&
Global Const RT_ICON = 3&
Global Const RT_MENU = 4&
Global Const RT_DIALOG = 5&
Global Const RT_STRING = 6&
Global Const RT_FONTDIR = 7&
Global Const RT_FONT = 8&
Global Const RT_ACCELERATOR = 9&
Global Const RT_RCDATA = 10&

'GetFreeSystemResources
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

'GetWinFlags
Global Const WF_PMODE = &H1
Global Const WF_CPU286 = &H2
Global Const WF_CPU386 = &H4
Global Const WF_CPU486 = &H8
Global Const WF_STANDARD = &H10
Global Const WF_WIN286 = &H10
Global Const WF_ENHANCED = &H20
Global Const WF_WIN386 = &H20
Global Const WF_CPU086 = &H40
Global Const WF_CPU186 = &H80
Global Const WF_LARGEFRAME = &H100
Global Const WF_SMALLFRAME = &H200
Global Const WF_80x87 = &H400

'Parameter error checking
Global Const ERR_WARNING = 8
Global Const ERR_PARAM = 4
Global Const ERR_SIZE_MASK = 3
Global Const ERR_BYTE = 1
Global Const ERR_WORD = 2
Global Const ERR_DWORD = 3
Global Const ERR_BAD_VALUE = &H6001
Global Const ERR_BAD_FLAGS = &H6002
Global Const ERR_BAD_INDEX = &H6003
Global Const ERR_BAD_DVALUE = &H7004
Global Const ERR_BAD_DFLAGS = &H7005
Global Const ERR_BAD_DINDEX = &H7006
Global Const ERR_BAD_PTR = &H7007
Global Const ERR_BAD_FUNC_PTR = &H7008
Global Const ERR_BAD_SELECTOR = &H6009
Global Const ERR_BAD_STRING_PTR = &H700A
Global Const ERR_BAD_HANDLE = &H600B

'KERNEL parameter errors
Global Const ERR_BAD_HINSTANCE = &H6020
Global Const ERR_BAD_HMODULE = &H6021
Global Const ERR_BAD_GLOBAL_HANDLE = &H6022
Global Const ERR_BAD_LOCAL_HANDLE = &H6023
Global Const ERR_BAD_ATOM = &H6024
Global Const ERR_BAD_HFILE = &H6025

'USER parameter errors
Global Const ERR_BAD_HWND = &H6040
Global Const ERR_BAD_HMENU = &H6041
Global Const ERR_BAD_HCURSOR = &H6042
Global Const ERR_BAD_HICON = &H6043
Global Const ERR_BAD_HDWP = &H6044
Global Const ERR_BAD_CID = &H6045
Global Const ERR_BAD_HDRVR = &H6046

'GDI parameter errors
Global Const ERR_BAD_COORDS = &H7060
Global Const ERR_BAD_GDI_OBJECT = &H6061
Global Const ERR_BAD_HDC = &H6062
Global Const ERR_BAD_HPEN = &H6063
Global Const ERR_BAD_HFONT = &H6064
Global Const ERR_BAD_HBRUSH = &H6065
Global Const ERR_BAD_HBITMAP = &H6066
Global Const ERR_BAD_HRGN = &H6067
Global Const ERR_BAD_HPALETTE = &H6068
Global Const ERR_BAD_HMETAFILE = &H6069

'KERNEL errors
Global Const ERR_GALLOC = &H1
Global Const ERR_GREALLOC = &H2
Global Const ERR_GLOCK = &H3
Global Const ERR_LALLOC = &H4
Global Const ERR_LREALLOC = &H5
Global Const ERR_LLOCK = &H6
Global Const ERR_ALLOCRES = &H7
Global Const ERR_LOCKRES = &H8
Global Const ERR_LOADMODULE = &H9

'USER errors
Global Const ERR_CREATEDLG = &H40
Global Const ERR_CREATEDLG2 = &H41
Global Const ERR_REGISTERCLASS = &H42
Global Const ERR_DCBUSY = &H43
Global Const ERR_CREATEWND = &H44
Global Const ERR_STRUCEXTRA = &H45
Global Const ERR_LOADSTR = &H46
Global Const ERR_LOADMENU = &H47
Global Const ERR_NESTEDBEGINPAINT = &H48
Global Const ERR_BADINDEX = &H49
Global Const ERR_CREATEMENU = &H4A

'GDI errors
Global Const ERR_CREATEDC = &H80
Global Const ERR_CREATEMETA = &H81
Global Const ERR_DELOBJSELECTED = &H82
Global Const ERR_SELBITMAP = &H83

'Exit Window parameters

Global Const EW_RESTARTWindow = &H42
Global Const EW_REBOOTSYSTEM = &H43


'Stock system bitmaps
Global Const OBM_CLOSE = 32754
Global Const OBM_UPARROW = 32753
Global Const OBM_DNARROW = 32752
Global Const OBM_RGARROW = 32751
Global Const OBM_LFARROW = 32750
Global Const OBM_REDUCE = 32749
Global Const OBM_ZOOM = 32748
Global Const OBM_RESTORE = 32747
Global Const OBM_REDUCED = 32746
Global Const OBM_ZOOMD = 32745
Global Const OBM_RESTORED = 32744
Global Const OBM_UPARROWD = 32743
Global Const OBM_DNARROWD = 32742
Global Const OBM_RGARROWD = 32741
Global Const OBM_LFARROWD = 32740
Global Const OBM_MNARROW = 32739
Global Const OBM_COMBO = 32738
Global Const OBM_UPARROWI = 32737
Global Const OBM_DNARROWI = 32736
Global Const OBM_RGARROWI = 32735
Global Const OBM_LFARROWI = 32734
Global Const OBM_OLD_CLOSE = 32767
Global Const OBM_SIZE = 32766
Global Const OBM_OLD_UPARROW = 32765
Global Const OBM_OLD_DNARROW = 32764
Global Const OBM_OLD_RGARROW = 32763
Global Const OBM_OLD_LFARROW = 32762
Global Const OBM_BTSIZE = 32761
Global Const OBM_CHECK = 32760
Global Const OBM_CHECKBOXES = 32759
Global Const OBM_BTNCORNERS = 32758
Global Const OBM_OLD_REDUCE = 32757
Global Const OBM_OLD_ZOOM = 32756
Global Const OBM_OLD_RESTORE = 32755

'Stock system Icons
Global Const OCR_NORMAL = 32512
Global Const OCR_IBEAM = 32513
Global Const OCR_WAIT = 32514
Global Const OCR_CROSS = 32515
Global Const OCR_UP = 32516
Global Const OCR_SIZE = 32640
Global Const OCR_ICON = 32641
Global Const OCR_SIZENWSE = 32642
Global Const OCR_SIZENESW = 32643
Global Const OCR_SIZEWE = 32644
Global Const OCR_SIZENS = 32645
Global Const OCR_SIZEALL = 32646
Global Const OCR_ICOCUR = 32647
Global Const OIC_SAMPLE = 32512
Global Const OIC_HAND = 32513
Global Const OIC_QUES = 32514
Global Const OIC_BANG = 32515
Global Const OIC_NOTE = 32516

'Raster-ops (Binary)
Global Const R2_BLACK = 1 ' 0
Global Const R2_NOTMERGEPEN = 2 'DPon
Global Const R2_MASKNOTPEN = 3'DPna
Global Const R2_NOTCOPYPEN = 4'PN
Global Const R2_MASKPENNOT = 5'PDna
Global Const R2_NOT = 6 'Dn
Global Const R2_XORPEN = 7'DPx
Global Const R2_NOTMASKPEN = 8'DPan
Global Const R2_MASKPEN = 9 'DPa
Global Const R2_NOTXORPEN = 10'DPxn
Global Const R2_NOP = 11'D
Global Const R2_MERGENOTPEN = 12'DPno
Global Const R2_COPYPEN = 13'P
Global Const R2_MERGEPENNOT = 14'PDno
Global Const R2_MERGEPEN = 15 'DPo
Global Const R2_WHITE = 16' 1

'Raster-ops (Ternary)
Global Const SRCCOPY = &HCC0020
Global Const SRCPAINT = &HEE0086
Global Const SRCAND = &H8800C6
Global Const SRCINVERT = &H660046
Global Const SRCERASE = &H440328
Global Const NOTSRCCOPY = &H330008
Global Const NOTSRCERASE = &H1100A6
Global Const MERGECOPY = &HC000CA
Global Const MERGEPAINT = &HBB0226
Global Const PATCOPY = &HF00021
Global Const PATPAINT = &HFB0A09
Global Const PATINVERT = &H5A0049
Global Const DSTINVERT = &H550009
Global Const BLACKNESS = &H42&
Global Const WHITENESS = &HFF0062

'StretchBlt() Modes
Global Const BLACKONWHITE = 1
Global Const WHITEONBLACK = 2
Global Const COLORONCOLOR = 3

'PolyFill() Modes
Global Const ALTERNATE = 1
Global Const WINDING = 2

'Text Alignment Options
Global Const TA_NOUPDATECP = 0
Global Const TA_UPDATECP = 1
Global Const TA_LEFT = 0
Global Const TA_RIGHT = 2
Global Const TA_CENTER = 6
Global Const TA_TOP = 0
Global Const TA_BOTTOM = 8
Global Const TA_BASELINE = 24

'ExtTextOut flags
Global Const ETO_GRAYED = 1
Global Const ETO_OPAQUE = 2
Global Const ETO_CLIPPED = 4

'SetMapperFlags
Global Const ASPECT_FILTERING = &H1

'Metafile Functions
Global Const META_SETBKCOLOR = &H201
Global Const META_SETBKMODE = &H102
Global Const META_SETMAPMODE = &H103
Global Const META_SETROP2 = &H104
Global Const META_SETRELABS = &H105
Global Const META_SETPOLYFILLMODE = &H106
Global Const META_SETSTRETCHBLTMODE = &H107
Global Const META_SETTEXTCHAREXTRA = &H108
Global Const META_SETTEXTCOLOR = &H209
Global Const META_SETTEXTJUSTIFICATION = &H20A
Global Const META_SETWINDOWORG = &H20B
Global Const META_SETWINDOWEXT = &H20C
Global Const META_SETVIEWPORTORG = &H20D
Global Const META_SETVIEWPORTEXT = &H20E
Global Const META_OFFSETWINDOWORG = &H20F
Global Const META_SCALEWINDOWEXT = &H400
Global Const META_OFFSETVIEWPORTORG = &H211
Global Const META_SCALEVIEWPORTEXT = &H412
Global Const META_LINETO = &H213
Global Const META_MOVETO = &H214
Global Const META_EXCLUDECLIPRECT = &H415
Global Const META_INTERSECTCLIPRECT = &H416
Global Const META_ARC = &H817
Global Const META_ELLIPSE = &H418
Global Const META_FLOODFILL = &H419
Global Const META_PIE = &H81A
Global Const META_RECTANGLE = &H41B
Global Const META_ROUNDRECT = &H61C
Global Const META_PATBLT = &H61D
Global Const META_SAVEDC = &H1E
Global Const META_SETPIXEL = &H41F
Global Const META_OFFSETCLIPRGN = &H220
Global Const META_TEXTOUT = &H521
Global Const META_BITBLT = &H922
Global Const META_STRETCHBLT = &HB23
Global Const META_POLYGON = &H324
Global Const META_POLYLINE = &H325
Global Const META_ESCAPE = &H626
Global Const META_RESTOREDC = &H127
Global Const META_FILLREGION = &H228
Global Const META_FRAMEREGION = &H429
Global Const META_INVERTREGION = &H12A
Global Const META_PAINTREGION = &H12B
Global Const META_SELECTCLIPREGION = &H12C
Global Const META_SELECTOBJECT = &H12D
Global Const META_SETTEXTALIGN = &H12E
Global Const META_DRAWTEXT = &H62F
Global Const META_CHORD = &H830
Global Const META_SETMAPPERFLAGS = &H231
Global Const META_EXTTEXTOUT = &HA32
Global Const META_SETDIBTODEV = &HD33
Global Const META_SELECTPALETTE = &H234
Global Const META_REALIZEPALETTE = &H35
Global Const META_ANIMATEPALETTE = &H436
Global Const META_SETPALENTRIES = &H37
Global Const META_POLYPOLYGON = &H538
Global Const META_RESIZEPALETTE = &H139
Global Const META_DIBBITBLT = &H940
Global Const META_DIBSTRETCHBLT = &HB41
Global Const META_DIBCREATEPATTERNBRUSH = &H142
Global Const META_STRETCHDIB = &HF43
Global Const META_DELETEOBJECT = &H1F0
Global Const META_CREATEPALETTE = &HF7
Global Const META_CREATEBRUSH = &HF8
Global Const META_CREATEPATTERNBRUSH = &H1F9
Global Const META_CREATEPENINDIRECT = &H2FA
Global Const META_CREATEFONTINDIRECT = &H2FB
Global Const META_CREATEBRUSHINDIRECT = &H2FC
Global Const META_CREATEBITMAPINDIRECT = &H2FD
Global Const META_CREATEBITMAP = &H6FE
Global Const META_CREATEREGION = &H6FF

'Escape
Global Const NEWFRAME = 1
Global Const ABORTDOCCONST = 2
Global Const NEXTBAND = 3
Global Const SETCOLORTABLE = 4
Global Const GETCOLORTABLE = 5
Global Const FLUSHOUTPUT = 6
Global Const DRAFTMODE = 7
Global Const QUERYESCSUPPORT = 8
Global Const SETABORTPROCCONST = 9
Global Const STARTDOCCONST = 10
Global Const ENDDOCAPICONST = 11
Global Const GETPHYSPAGESIZE = 12
Global Const GETPRINTINGOFFSET = 13
Global Const GETSCALINGFACTOR = 14
Global Const MFCOMMENT = 15
Global Const GETPENWIDTH = 16
Global Const SETCOPYCOUNT = 17
Global Const SELECTPAPERSOURCE = 18
Global Const DEVICEDATA = 19
Global Const PASSTHROUGH = 19
Global Const GETTECHNOLGY = 20
Global Const GETTECHNOLOGY = 20
Global Const SETENDCAP = 21
Global Const SETLINEJOIN = 22
Global Const SETMITERLIMIT = 23
Global Const BANDINFO = 24
Global Const DRAWPATTERNRECT = 25
Global Const GETVECTORPENSIZE = 26
Global Const GETVECTORBRUSHSIZE = 27
Global Const ENABLEDUPLEX = 28
Global Const GETSETPAPERBINS = 29
Global Const GETSETPRINTORIENT = 30
Global Const ENUMPAPERBINS = 31
Global Const SETDIBSCALING = 32
Global Const EPSPRINTING = 33
Global Const ENUMPAPERMETRICS = 34
Global Const GETSETPAPERMETRICS = 35
Global Const POSTSCRIPT_DATA = 37
Global Const POSTSCRIPT_IGNORE = 38
Global Const GETEXTENDEDTEXTMETRICS = 256
Global Const GETEXTENTTABLE = 257
Global Const GETPAIRKERNTABLE = 258
Global Const GETTRACKKERNTABLE = 259
Global Const EXTTEXTOUTCONST = 512
Global Const ENABLERELATIVEWIDTHS = 768
Global Const ENABLEPAIRKERNING = 769
Global Const SETKERNTRACK = 770
Global Const SETALLJUSTVALUES = 771
Global Const SETCHARSET = 772
Global Const STRETCHBLTCONST = 2048
Global Const BEGIN_PATH = 4096
Global Const CLIP_TO_PATH = 4097
Global Const END_PATH = 4098
Global Const EXT_DEVICE_CAPS = 4099
Global Const RESTORE_CTM = 4100
Global Const SAVE_CTM = 4101
Global Const SET_ARC_DIRECTION = 4102
Global Const SET_BACKGROUND_COLOR = 4103
Global Const SET_POLY_MODE = 4104
Global Const SET_SCREEN_ANGLE = 4105
Global Const SET_SPREAD = 4106
Global Const TRANSFORM_CTM = 4107
Global Const SET_CLIP_BOX = 4108
Global Const SET_BOUNDS = 4109
Global Const SET_MIRROR_MODE = 4110

'Spooler Error Codes
Global Const SP_NOTREPORTED = &H4000
Global Const SP_ERROR = (-1)
Global Const SP_APPABORT = (-2)
Global Const SP_USERABORT = (-3)
Global Const SP_OUTOFDISK = (-4)
Global Const SP_OUTOFMEMORY = (-5)
Global Const PR_JOBSTATUS = &H0

'biCompression field constants for DIB
Global Const BI_RGB = 0&
Global Const BI_RLE8 = 1&
Global Const BI_RLE4 = 2&

'LOGFONT and TEXTMETRIC
Global Const OUT_DEFAULT_PRECIS = 0
Global Const OUT_STRING_PRECIS = 1
Global Const OUT_CHARACTER_PRECIS = 2
Global Const OUT_STROKE_PRECIS = 3
Global Const OUT_TT_PRECIS = 4
Global Const OUT_DEVICE_PRECIS = 5
Global Const OUT_RASTER_PRECIS = 6
Global Const OUT_TT_ONLY_PRECIS = 7
Global Const CLIP_DEFAULT_PRECIS = 0
Global Const CLIP_CHARACTER_PRECIS = 1
Global Const CLIP_STROKE_PRECIS = 2
Global Const CLIP_LH_ANGLES = &H10
Global Const CLIP_TT_ALWAYS = &H20
Global Const CLIP_EMBEDDED = &H80
Global Const DEFAULT_QUALITY = 0
Global Const DRAFT_QUALITY = 1
Global Const PROOF_QUALITY = 2
Global Const DEFAULT_PITCH = 0
Global Const FIXED_PITCH = 1
Global Const VARIABLE_PITCH = 2
Global Const TMPF_FIXED_PITCH = 1
Global Const TMPF_VECTOR = 2
Global Const TMPF_DEVICE = 8
Global Const TMPF_TRUETYPE = 4
Global Const ANSI_CHARSET = 0
Global Const DEFAULT_CHARSET = 1
Global Const SYMBOL_CHARSET = 2
Global Const SHIFTJIS_CHARSET = 128
Global Const OEM_CHARSET = 255
Global Const NTM_REGULAR = &H40&
Global Const NTM_BOLD = &H20&
Global Const NTM_ITALIC = &H1&
Global Const LF_FULLFACESIZE = 64
Global Const RASTER_FONTTYPE = 1
Global Const DEVICE_FONTTYPE = 2
Global Const TRUETYPE_FONTTYPE = 4


'Font Families
Global Const FF_DONTCARE = 0
Global Const FF_ROMAN = 16

'Times Roman, Century Schoolbook, etc.
Global Const FF_SWISS = 32

' Helvetica, Swiss, etc.
Global Const FF_MODERN = 48

' Pica, Elite, Courier, etc.
Global Const FF_SCRIPT = 64
Global Const FF_DECORATIVE = 80

'Font Weights
Global Const FW_DONTCARE = 0
Global Const FW_THIN = 100
Global Const FW_EXTRALIGHT = 200
Global Const FW_LIGHT = 300
Global Const FW_NORMAL = 400
Global Const FW_MEDIUM = 500
Global Const FW_SEMIBOLD = 600
Global Const FW_BOLD = 700
Global Const FW_EXTRABOLD = 800
Global Const FW_HEAVY = 900
Global Const FW_ULTRALIGHT = FW_EXTRALIGHT
Global Const FW_REGULAR = FW_NORMAL
Global Const FW_DEMIBOLD = FW_SEMIBOLD
Global Const FW_ULTRABOLD = FW_EXTRABOLD
Global Const FW_BLACK = FW_HEAVY

'Background Modes
'Global Const TRANSPARENT = 1
Global Const OPAQUE = 2

'Mapping Modes
Global Const MM_TEXT = 1
Global Const MM_LOMETRIC = 2
Global Const MM_HIMETRIC = 3
Global Const MM_LOENGLISH = 4
Global Const MM_HIENGLISH = 5
Global Const MM_TWIPS = 6
Global Const MM_ISOTROPIC = 7
Global Const MM_ANISOTROPIC = 8
'Coordinate Modes
Global Const ABSOLUTE = 1
Global Const RELATIVE = 2

'Stock Logical Objects
Global Const WHITE_BRUSH = 0
Global Const LTGRAY_BRUSH = 1
Global Const GRAY_BRUSH = 2
Global Const DKGRAY_BRUSH = 3
Global Const BLACK_BRUSH = 4
Global Const NULL_BRUSH = 5
Global Const HOLLOW_BRUSH = NULL_BRUSH
Global Const WHITE_PEN = 6
Global Const BLACK_PEN = 7
Global Const NULL_PEN = 8
Global Const OEM_FIXED_FONT = 10
Global Const ANSI_FIXED_FONT = 11
Global Const ANSI_VAR_FONT = 12
Global Const SYSTEM_FONT = 13
Global Const DEVICE_DEFAULT_FONT = 14
Global Const DEFAULT_PALETTE = 15
Global Const SYSTEM_FIXED_FONT = 16

'Brush Styles
Global Const BS_SOLID = 0
Global Const BS_NULL = 1
Global Const BS_HOLLOW = BS_NULL
Global Const BS_HATCHED = 2
Global Const BS_PATTERN = 3
Global Const BS_INDEXED = 4
Global Const BS_DIBPATTERN = 5

'Hatch Styles
Global Const HS_HORIZONTAL = 0
Global Const HS_VERTICAL = 1
Global Const HS_FDIAGONAL = 2
Global Const HS_BDIAGONAL = 3
Global Const HS_CROSS = 4
Global Const HS_DIAGCROSS = 5

'Pen Styles
Global Const PS_SOLID = 0
Global Const PS_DASH = 1
Global Const PS_DOT = 2
Global Const PS_DASHDOT = 3
Global Const PS_DASHDOTDOT = 4
Global Const PS_NULL = 5
Global Const PS_INSIDEFRAME = 6

'Bounds Rectangle Constants
Global Const DCB_RESET = 1
Global Const DCB_ACCUMULATE = 2
Global Const DCB_DIRTY = 2
Global Const DCB_SET = 3
Global Const DCB_ENABLE = 4
Global Const DCB_DISABLE = 8

'GetDeviceCaps() Device Parameters
Global Const DRIVERVERSION = 0
Global Const TECHNOLOGY = 2
Global Const HORZSIZE = 4
Global Const VERTSIZE = 6
Global Const HORZRES = 8
Global Const VERTRES = 10
Global Const BITSPIXEL = 12
Global Const PLANES = 14
Global Const NUMBRUSHES = 16
Global Const NUMPENS = 18
Global Const NUMMARKERS = 20
Global Const NUMFONTS = 22
Global Const NUMCOLORS = 24
Global Const PDEVICESIZE = 26
Global Const CURVECAPS = 28
Global Const LINECAPS = 30
Global Const POLYGONALCAPS = 32
Global Const TEXTCAPS = 34
Global Const CLIPCAPS = 36
Global Const RASTERCAPS = 38
Global Const ASPECTX = 40
Global Const ASPECTY = 42
Global Const ASPECTXY = 44
Global Const LOGPIXELSX = 88
Global Const LOGPIXELSY = 90
Global Const SIZEPALETTE = 104
Global Const NUMRESERVED = 106
Global Const COLORRES = 108

'Device Technologies
Global Const DT_PLOTTER = 0
Global Const DT_RASDISPLAY = 1
Global Const DT_RASPRINTER = 2
Global Const DT_RASCAMERA = 3
Global Const DT_CHARSTREAM = 4
Global Const DT_METAFILE = 5
Global Const DT_DISPFILE = 6

'Curve Capabilities
Global Const CC_NONE = 0
Global Const CC_CIRCLES = 1
Global Const CC_PIE = 2
Global Const CC_CHORD = 4
Global Const CC_ELLIPSES = 8
Global Const CC_WIDE = 16
Global Const CC_STYLED = 32
Global Const CC_WIDESTYLED = 64
Global Const CC_INTERIORS = 128

'Line Capabilities
Global Const LC_NONE = 0
Global Const LC_POLYLINE = 2
Global Const LC_MARKER = 4
Global Const LC_POLYMARKER = 8
Global Const LC_WIDE = 16
Global Const LC_STYLED = 32
Global Const LC_WIDESTYLED = 64
Global Const LC_INTERIORS = 128

'Polygonal Capabilities
Global Const PC_NONE = 0
Global Const PC_POLYGON = 1
Global Const PC_RECTANGLE = 2
Global Const PC_WINDPOLYGON = 4
Global Const PC_TRAPEZOID = 4
Global Const PC_SCANLINE = 8
Global Const PC_WIDE = 16
Global Const PC_STYLED = 32
Global Const PC_WIDESTYLED = 64
Global Const PC_INTERIORS = 128

'Polygonal Capabilities
Global Const CP_NONE = 0
Global Const CP_RECTANGLE = 1

'Text Capabilities
Global Const TC_OP_CHARACTER = &H1
Global Const TC_OP_STROKE = &H2
Global Const TC_CP_STROKE = &H4
Global Const TC_CR_90 = &H8
Global Const TC_CR_ANY = &H10
Global Const TC_SF_X_YINDEP = &H20
Global Const TC_SA_DOUBLE = &H40
Global Const TC_SA_INTEGER = &H80
Global Const TC_SA_CONTIN = &H100
Global Const TC_EA_DOUBLE = &H200
Global Const TC_IA_ABLE = &H400
Global Const TC_UA_ABLE = &H800
Global Const TC_SO_ABLE = &H1000
Global Const TC_RA_ABLE = &H2000
Global Const TC_VA_ABLE = &H4000
Global Const TC_RESERVED = &H8000

'Raster Capabilities
Global Const RC_BITBLT = 1
Global Const RC_BANDING = 2
Global Const RC_SCALING = 4
Global Const RC_BITMAP64 = 8
Global Const RC_GDI20_OUTPUT = &H10
Global Const RC_DI_BITMAP = &H80
Global Const RC_PALETTE = &H100
Global Const RC_DIBTODEV = &H200
Global Const RC_BIGFONT = &H400
Global Const RC_STRETCHBLT = &H800
Global Const RC_FLOODFILL = &H1000
Global Const RC_STRETCHDIB = &H2000

'palette entry flags
Global Const PC_RESERVED = &H1
Global Const PC_EXPLICIT = &H2
Global Const PC_NOCOLLAPSE = &H4

'DIB color table identifiers
Global Const DIB_RGB_COLORS = 0
Global Const DIB_PAL_COLORS = 1

'constants for Get/SetSystemPaletteUse()
Global Const SYSPAL_STATIC = 1
Global Const SYSPAL_NOSTATIC = 2

'constants for CreateDIBitmap
Global Const CBM_INIT = &H4&

'DrawText() Format Flags
Global Const DT_TOP = &H0
Global Const DT_LEFT = &H0
Global Const DT_CENTER = &H1
Global Const DT_RIGHT = &H2
Global Const DT_VCENTER = &H4
Global Const DT_BOTTOM = &H8
Global Const DT_WORDBREAK = &H10
Global Const DT_SINGLELINE = &H20
Global Const DT_EXPANDTABS = &H40
Global Const DT_TABSTOP = &H80
Global Const DT_NOCLIP = &H100
Global Const DT_EXTERNALLEADING = &H200
Global Const DT_CALCRECT = &H400
Global Const DT_NOPREFIX = &H800
Global Const DT_INTERNAL = &H1000

'ExtFloodFill style flags
Global Const FLOODFILLBORDER = 0
Global Const FLOODFILLSURFACE = 1


'Scroll Bar Constants
Global Const SB_HORZ = 0
Global Const SB_VERT = 1
Global Const SB_CTL = 2
Global Const SB_BOTH = 3

'Scroll Bar Commands
Global Const SB_LINEUP = 0
Global Const SB_LINEDOWN = 1
Global Const SB_PAGEUP = 2
Global Const SB_PAGEDOWN = 3
Global Const SB_THUMBPOSITION = 4
Global Const SB_THUMBTRACK = 5
Global Const SB_TOP = 6
Global Const SB_BOTTOM = 7
Global Const SB_ENDSCROLL = 8


'Old ShowWindow() Commands
Global Const HIDE_WINDOW = 0
Global Const SHOW_OPENWINDOW = 1
Global Const SHOW_ICONWINDOW = 2
Global Const SHOW_FULLSCREEN = 3
Global Const SHOW_OPENNOACTIVATE = 4

'Identifiers for the WM_SHOWWINDOW message
Global Const SW_PARENTCLOSING = 1
Global Const SW_OTHERZOOM = 2
Global Const SW_PARENTOPENING = 3
Global Const SW_OTHERUNZOOM = 4

'RedrawWindow flags
Global Const RDW_INVALIDATE = &H1
Global Const RDW_INTERNALPAINT = &H2
Global Const RDW_ERASE = &H4
Global Const RDW_VALIDATE = &H8
Global Const RDW_NOINTERNALPAINT = &H10
Global Const RDW_NOERASE = &H20
Global Const RDW_NOCHILDREN = &H40
Global Const RDW_ALLCHILDREN = &H80
Global Const RDW_UPDATENOW = &H100
Global Const RDW_ERASENOW = &H200
Global Const RDW_FRAME = &H400
Global Const RDW_NOFRAME = &H800

'ScrollWindowEx flags
Global Const SW_SCROLLCHILDREN = &H1
Global Const SW_INVALIDATE = &H2
Global Const SW_ERASE = &H4

'Region Flags
Global Const ERRORAPI = 0
Global Const NULLREGION = 1
Global Const SIMPLEREGION = 2
Global Const COMPLEXREGION = 3

'CombineRgn() Styles
Global Const RGN_AND = 1
Global Const RGN_OR = 2
Global Const RGN_XOR = 3
Global Const RGN_DIFF = 4
Global Const RGN_COPY = 5

'Virtual Keys, Standard Set
Global Const VK_LBUTTON = &H1
Global Const VK_RBUTTON = &H2
Global Const VK_CANCEL = &H3
Global Const VK_MBUTTON = &H4
Global Const VK_BACK = &H8
Global Const VK_TAB = &H9
Global Const VK_CLEAR = &HC
Global Const VK_RETURN = &HD
Global Const VK_SHIFT = &H10
Global Const VK_CONTROL = &H11
Global Const VK_MENU = &H12
Global Const VK_PAUSE = &H13
Global Const VK_CAPITAL = &H14
Global Const VK_ESCAPE = &H1B
Global Const VK_SPACE = &H20
Global Const VK_PRIOR = &H21
Global Const VK_NEXT = &H22
Global Const VK_END = &H23
Global Const VK_HOME = &H24
Global Const VK_LEFT = &H25
Global Const VK_UP = &H26
Global Const VK_RIGHT = &H27
Global Const VK_DOWN = &H28
Global Const VK_SELECT = &H29
Global Const VK_PRINT = &H2A
Global Const VK_EXECUTE = &H2B
Global Const VK_SNAPSHOT = &H2C
Global Const VK_INSERT = &H2D
Global Const VK_DELETE = &H2E
Global Const VK_HELP = &H2F
Global Const VK_NUMPAD0 = &H60
Global Const VK_NUMPAD1 = &H61
Global Const VK_NUMPAD2 = &H62
Global Const VK_NUMPAD3 = &H63
Global Const VK_NUMPAD4 = &H64
Global Const VK_NUMPAD5 = &H65
Global Const VK_NUMPAD6 = &H66
Global Const VK_NUMPAD7 = &H67
Global Const VK_NUMPAD8 = &H68
Global Const VK_NUMPAD9 = &H69
Global Const VK_MULTIPLY = &H6A
Global Const VK_ADD = &H6B
Global Const VK_SEPARATOR = &H6C
Global Const VK_SUBTRACT = &H6D
Global Const VK_DECIMAL = &H6E
Global Const VK_DIVIDE = &H6F
Global Const VK_F1 = &H70
Global Const VK_F2 = &H71
Global Const VK_F3 = &H72
Global Const VK_F4 = &H73
Global Const VK_F5 = &H74
Global Const VK_F6 = &H75
Global Const VK_F7 = &H76
Global Const VK_F8 = &H77
Global Const VK_F9 = &H78
Global Const VK_F10 = &H79
Global Const VK_F11 = &H7A
Global Const VK_F12 = &H7B
Global Const VK_F13 = &H7C
Global Const VK_F14 = &H7D
Global Const VK_F15 = &H7E
Global Const VK_F16 = &H7F
Global Const VK_F17 = &H80
Global Const VK_F18 = &H81
Global Const VK_F19 = &H82
Global Const VK_F20 = &H83
Global Const VK_F21 = &H84
Global Const VK_F22 = &H85
Global Const VK_F23 = &H86
Global Const VK_F24 = &H87
Global Const VK_NUMLOCK = &H90
Global Const VK_SCROLL = &H91

'Queue Status
Global Const QS_KEY = 1
Global Const QS_MOUSEMOVE = 2
Global Const QS_MOUSEBUTTON = 4
Global Const QS_MOUSE = 6
Global Const QS_POSTMESSAGE = 8
Global Const QS_TIMER = &H10
Global Const QS_PAINT = &H20
Global Const QS_SENDMESSAGE = &H40
Global Const QS_ALLINPUT = &H7F

'SetWindowHook() codes
Global Const WH_MSGFILTER = (-1)
Global Const WH_JOURNALRECORD = 0
Global Const WH_JOURNALPLAYBACK = 1
Global Const WH_KEYBOARD = 2
Global Const WH_GETMESSAGE = 3
Global Const WH_CALLWNDPROC = 4
Global Const WH_CBT = 5
Global Const WH_SYSMSGFILTER = 6
Global Const WH_WINDOWMGR = 7
Global Const WH_HARDWARE = 8
Global Const WH_SHELL = 10

'Hook Codes
Global Const HC_LPLPFNNEXT = (-2)
Global Const HC_LPFNNEXT = (-1)
Global Const HC_ACTION = 0
Global Const HC_GETNEXT = 1
Global Const HC_SKIP = 2
Global Const HC_NOREM = 3
Global Const HC_NOREMOVE = 3
Global Const HC_SYSMODALON = 4
Global Const HC_SYSMODALOFF = 5

'CBT Hook Codes
Global Const HCBT_MOVESIZE = 0
Global Const HCBT_MINMAX = 1
Global Const HCBT_QS = 2

'WH_MSGFILTER Filter Proc Codes
Global Const MSGF_DIALOGBOX = 0
Global Const MSGF_MESSAGEBOX = 1
Global Const MSGF_MENU = 2
Global Const MSGF_MOVE = 3
Global Const MSGF_SIZE = 4
Global Const MSGF_SCROLLBAR = 5
Global Const MSGF_NEXTWINDOW = 6

'Window Manager Hook Codes
Global Const WC_INIT = 1
Global Const WC_SWP = 2
Global Const WC_DEFWINDOWPROC = 3
Global Const WC_MINMAX = 4
Global Const WC_MOVE = 5
Global Const WC_SIZE = 6
Global Const WC_DRAWCAPTION = 7

'Window field offsets for GetWindowLong() and GetWindowWord()
Global Const GWL_WNDPROC = (-4)
Global Const GWW_HINSTANCE = (-6)
Global Const GWW_HWNDPARENT = (-8)
Global Const GWW_ID = (-12)
Global Const GWL_STYLE = (-16)
Global Const GWL_EXSTYLE = (-20)

'GetWindowLong and and GetWindowWord dialog box constants
Global Const DWL_MSGRESULT = 0
Global Const DWL_DLGPROC = 4
Global Const DWL_USER = 8

'Class field offsets for GetClassLong() and GetClassWord()
Global Const GCL_MENUNAME = (-8)
Global Const GCW_HBRBACKGROUND = (-10)
Global Const GCW_HCURSOR = (-12)
Global Const GCW_HICON = (-14)
Global Const GCW_HMODULE = (-16)
Global Const GCW_CBWNDEXTRA = (-18)
Global Const GCW_CBCLSEXTRA = (-20)
Global Const GCL_WNDPROC = (-24)
Global Const GCW_STYLE = (-26)
Global Const GCW_ATOM = (-32)

'SendMessage Flag
Global Const HWND_BROADCAST = -1

'Window Messages
Global Const WM_NULL = &H0
Global Const WM_CREATE = &H1
Global Const WM_DESTROY = &H2
Global Const WM_MOVE = &H3
Global Const WM_SIZE = &H5
Global Const WM_ACTIVATE = &H6
Global Const WM_SETFOCUS = &H7
Global Const WM_KILLFOCUS = &H8
Global Const WM_ENABLE = &HA
Global Const WM_SETREDRAW = &HB
Global Const WM_GETTEXT = &HD
Global Const WM_GETTEXTLENGTH = &HE
Global Const WM_PAINT = &HF
Global Const WM_CLOSE = &H10
Global Const WM_QUERYENDSESSION = &H11
Global Const WM_QUIT = &H12
Global Const WM_QUERYOPEN = &H13
Global Const WM_ERASEBKGND = &H14
Global Const WM_SYSCOLORCHANGE = &H15
Global Const WM_ENDSESSION = &H16
Global Const WM_SYSTEMERROR = &H17
Global Const WM_SHOWWINDOW = &H18
Global Const WM_CTLCOLOR = &H19
Global Const WM_WININICHANGE = &H1A
Global Const WM_DEVMODECHANGE = &H1B
Global Const WM_ACTIVATEAPP = &H1C
Global Const WM_FONTCHANGE = &H1D
Global Const WM_TIMECHANGE = &H1E
Global Const WM_CANCELMODE = &H1F
Global Const WM_SETCURSOR = &H20
Global Const WM_MOUSEACTIVATE = &H21
Global Const WM_CHILDACTIVATE = &H22
Global Const WM_QUEUESYNC = &H23
Global Const WM_GETMINMAXINFO = &H24
Global Const WM_PAINTICON = &H26
Global Const WM_ICONERASEBKGND = &H27
Global Const WM_NEXTDLGCTL = &H28
Global Const WM_SPOOLERSTATUS = &H2A
Global Const WM_DRAWITEM = &H2B
Global Const WM_MEASUREITEM = &H2C
Global Const WM_DELETEITEM = &H2D
Global Const WM_VKEYTOITEM = &H2E
Global Const WM_CHARTOITEM = &H2F
Global Const WM_SETFONT = &H30
Global Const WM_GETFONT = &H31
Global Const WM_COMMNOTIFY = &H44
Global Const WM_QUERYDRAGICON = &H37
Global Const WM_COMPAREITEM = &H39
Global Const WM_COMPACTING = &H41
Global Const WM_WINDOWPOSCHANGING = &H46
Global Const WM_WINDOWPOSCHANGED = &H47
Global Const WM_POWER = &H48
Global Const WM_NCCREATE = &H81
Global Const WM_NCDESTROY = &H82
Global Const WM_NCCALCSIZE = &H83
Global Const WM_NCHITTEST = &H84
Global Const WM_NCPAINT = &H85
Global Const WM_NCACTIVATE = &H86
Global Const WM_GETDLGCODE = &H87
Global Const WM_NCMOUSEMOVE = &HA0
Global Const WM_NCLBUTTONDOWN = &HA1
Global Const WM_NCLBUTTONUP = &HA2
Global Const WM_NCLBUTTONDBLCLK = &HA3
Global Const WM_NCRBUTTONDOWN = &HA4
Global Const WM_NCRBUTTONUP = &HA5
Global Const WM_NCRBUTTONDBLCLK = &HA6
Global Const WM_NCMBUTTONDOWN = &HA7
Global Const WM_NCMBUTTONUP = &HA8
Global Const WM_NCMBUTTONDBLCLK = &HA9
Global Const WM_KEYFIRST = &H100
Global Const WM_CHAR = &H102
Global Const WM_DEADCHAR = &H103
Global Const WM_SYSKEYDOWN = &H104
Global Const WM_SYSKEYUP = &H105
Global Const WM_SYSCHAR = &H106
Global Const WM_SYSDEADCHAR = &H107
Global Const WM_KEYLAST = &H108
Global Const WM_INITDIALOG = &H110
Global Const WM_COMMAND = &H111
Global Const WM_SYSCOMMAND = &H112
Global Const WM_TIMER = &H113
Global Const WM_HSCROLL = &H114
Global Const WM_VSCROLL = &H115
Global Const WM_INITMENU = &H116
Global Const WM_INITMENUPOPUP = &H117
Global Const WM_MENUSELECT = &H11F
Global Const WM_MENUCHAR = &H120
Global Const WM_ENTERIDLE = &H121
Global Const WM_MOUSEFIRST = &H200
Global Const WM_MOUSEMOVE = &H200
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Global Const WM_LBUTTONDBLCLK = &H203
Global Const WM_RBUTTONDOWN = &H204
Global Const WM_RBUTTONUP = &H205
Global Const WM_RBUTTONDBLCLK = &H206
Global Const WM_MBUTTONDOWN = &H207
Global Const WM_MBUTTONUP = &H208
Global Const WM_MBUTTONDBLCLK = &H209
Global Const WM_MOUSELAST = &H209
Global Const WM_PARENTNOTIFY = &H210
Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230
Global Const WM_DROPFILES = &H233
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Global Const WM_CLEAR = &H303
Global Const WM_UNDO = &H304
Global Const WM_RENDERFORMAT = &H305
Global Const WM_RENDERALLFORMATS = &H306
Global Const WM_DESTROYCLIPBOARD = &H307
Global Const WM_DRAWCLIPBOARD = &H308
Global Const WM_PAINTCLIPBOARD = &H309
Global Const WM_VSCROLLCLIPBOARD = &H30A
Global Const WM_SIZECLIPBOARD = &H30B
Global Const WM_ASKCBFORMATNAME = &H30C
Global Const WM_CHANGECBCHAIN = &H30D
Global Const WM_HSCROLLCLIPBOARD = &H30E
Global Const WM_QUERYNEWPALETTE = &H30F
Global Const WM_PALETTEISCHANGING = &H310
Global Const WM_PALETTECHANGED = &H311
'WM_SYNCTASK Commands
Global Const ST_BEGINSWP = 0
Global Const ST_ENDSWP = 1


'WM_ACTIVATE constants
Global Const WA_INACTIVE = 0
Global Const WA_ACTIVE = 1
Global Const WA_CLICKACTIVE = 2


'WinWhere() Area Codes
Global Const HTERROR = (-2)
Global Const HTTRANSPARENT = (-1)
Global Const HTNOWHERE = 0
Global Const HTCLIENT = 1
Global Const HTCAPTION = 2
Global Const HTSYSMENU = 3
Global Const HTGROWBOX = 4
Global Const HTSIZE = HTGROWBOX
Global Const HTMENU = 5
Global Const HTHSCROLL = 6
Global Const HTVSCROLL = 7
Global Const HTREDUCE = 8
Global Const HTZOOM = 9
Global Const HTLEFT = 10
Global Const HTRIGHT = 11
Global Const HTTOP = 12
Global Const HTTOPLEFT = 13
Global Const HTTOPRIGHT = 14
Global Const HTBOTTOM = 15
Global Const HTBOTTOMLEFT = 16
Global Const HTBOTTOMRIGHT = 17
Global Const HTSIZEFIRST = HTLEFT
Global Const HTSIZELAST = HTBOTTOMRIGHT

'WM_MOUSEACTIVATE Return Codes
Global Const MA_ACTIVATE = 1
Global Const MA_ACTIVATEANDEAT = 2
Global Const MA_NOACTIVATE = 3
Global Const MA_NOACTIVATEANDEAT = 4


'Size Message Commands
Global Const SIZENORMAL = 0
Global Const SIZEICONIC = 1
Global Const SIZEFULLSCREEN = 2
Global Const SIZEZOOMSHOW = 3
Global Const SIZEZOOMHIDE = 4

'Key State Masks for Mouse Messages
Global Const MK_LBUTTON = &H1
Global Const MK_RBUTTON = &H2
Global Const MK_SHIFT = &H4
Global Const MK_CONTROL = &H8
Global Const MK_MBUTTON = &H10

'Window Styles
Global Const WS_OVERLAPPED = &H0&
Global Const WS_POPUP = &H80000000
Global Const WS_CHILD = &H40000000
Global Const WS_MINIMIZE = &H20000000
Global Const WS_VISIBLE = &H10000000
Global Const WS_DISABLED = &H8000000
Global Const WS_CLIPSIBLINGS = &H4000000
Global Const WS_CLIPCHILDREN = &H2000000
Global Const WS_MAXIMIZE = &H1000000
Global Const WS_CAPTION = &HC00000
Global Const WS_BORDER = &H800000
Global Const WS_DLGFRAME = &H400000
Global Const WS_VSCROLL = &H200000
Global Const WS_HSCROLL = &H100000
Global Const WS_SYSMENU = &H80000
Global Const WS_THICKFRAME = &H40000
Global Const WS_GROUP = &H20000
Global Const WS_TABSTOP = &H10000
Global Const WS_MINIMIZEBOX = &H20000
Global Const WS_MAXIMIZEBOX = &H10000
Global Const WS_TILED = WS_OVERLAPPED
Global Const WS_ICONIC = WS_MINIMIZE
Global Const WS_SIZEBOX = WS_THICKFRAME
'Common Window Styles
Global Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Global Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Global Const WS_CHILDWINDOW = (WS_CHILD)
Global Const WS_TILEDWINDOW = (WS_OVERLAPPEDWINDOW)

'Extended Window Styles
Global Const WS_EX_DLGMODALFRAME = &H1&
Global Const WS_EX_NOPARENTNOTIFY = &H4&
Global Const WS_EX_TOPMOST = &H8&
Global Const WS_EX_ACCEPTFILES = &H10&
Global Const WS_EX_TRANSPARENT = &H20&

' MDI style allows use of all child styles
Global Const MDIS_ALLCHILDSTYLES = &H1&

'Class styles
Global Const CS_VREDRAW = &H1
Global Const CS_HREDRAW = &H2
Global Const CS_KEYCVTWINDOW = &H4
Global Const CS_DBLCLKS = &H8
Global Const CS_OWNDC = &H20
Global Const CS_CLASSDC = &H40
Global Const CS_PARENTDC = &H80
Global Const CS_NOKEYCVT = &H100
Global Const CS_NOCLOSE = &H200
Global Const CS_SAVEBITS = &H800
Global Const CS_BYTEALIGNCLIENT = &H1000
Global Const CS_BYTEALIGNWINDOW = &H2000
Global Const CS_GLOBALCLASS = &H4000

'Predefined Clipboard Formats
Global Const CF_TEXT = 1
Global Const CF_BITMAP = 2
Global Const CF_METAFILEPICT = 3
Global Const CF_SYLK = 4
Global Const CF_DIF = 5
Global Const CF_TIFF = 6
Global Const CF_OEMTEXT = 7
Global Const CF_DIB = 8
Global Const CF_PALETTE = 9
Global Const CF_OWNERDISPLAY = &H80
Global Const CF_DSPTEXT = &H81
Global Const CF_DSPBITMAP = &H82
Global Const CF_DSPMETAFILEPICT = &H83

'"Private" formats don't get GlobalFree()'d
Global Const CF_PRIVATEFIRST = &H200
Global Const CF_PRIVATELAST = &H2FF

'"GDIOBJ" formats do get DeleteObject()'d
Global Const CF_GDIOBJFIRST = &H300
Global Const CF_GDIOBJLAST = &H3FF


'Owner draw control types
Global Const ODT_MENU = 1
Global Const ODT_LISTBOX = 2
Global Const ODT_COMBOBOX = 3
Global Const ODT_BUTTON = 4

'Owner draw actions
Global Const ODA_DRAWENTIRE = &H1
Global Const ODA_SELECT = &H2
Global Const ODA_FOCUS = &H4

'Owner draw state
Global Const ODS_SELECTED = &H1
Global Const ODS_GRAYED = &H2
Global Const ODS_DISABLED = &H4
Global Const ODS_CHECKED = &H8
Global Const ODS_FOCUS = &H10


'PeekMessage() Options
Global Const PM_NOREMOVE = &H0
Global Const PM_REMOVE = &H1
Global Const PM_NOYIELD = &H2

'Flags for _lopen
Global Const READAPI = 0
Global Const WRITEAPI = 1
Global Const READ_WRITE = 2


'Window placement flags
Global Const CW_USEDEFAULT = &H8000
Global Const WPF_SETMINPOSITION = 1
Global Const WPF_RESTORETOMAXIMIZED = 2

'SetWindowPos Flags
Global Const SWP_NOSIZE = &H1
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOZORDER = &H4
Global Const SWP_NOREDRAW = &H8
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_DRAWFRAME = &H20
Global Const SWP_SHOWWINDOW = &H40
Global Const SWP_HIDEWINDOW = &H80
Global Const SWP_NOCOPYBITS = &H100
Global Const SWP_NOREPOSITION = &H200
Global Const SWP_NOSENDCHANGING = &H400
Global Const SWP_DEFERERASE = &H2000

'SetWindowPos() hwndInsertAfter values
Global Const HWND_TOP = 0
Global Const HWND_BOTTOM = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const DLGWINDOWEXTRA = 30

'GetSystemMetrics() codes
Global Const SM_CXSCREEN = 0
Global Const SM_CYSCREEN = 1
Global Const SM_CXVSCROLL = 2
Global Const SM_CYHSCROLL = 3
Global Const SM_CYCAPTION = 4
Global Const SM_CXBORDER = 5
Global Const SM_CYBORDER = 6
Global Const SM_CXDLGFRAME = 7
Global Const SM_CYDLGFRAME = 8
Global Const SM_CYVTHUMB = 9
Global Const SM_CXHTHUMB = 10
Global Const SM_CXICON = 11
Global Const SM_CYICON = 12
Global Const SM_CXCURSOR = 13
Global Const SM_CYCURSOR = 14
Global Const SM_CYMENU = 15
Global Const SM_CXFULLSCREEN = 16
Global Const SM_CYFULLSCREEN = 17
Global Const SM_CYKANJIWINDOW = 18
Global Const SM_MOUSEPRESENT = 19
Global Const SM_CYVSCROLL = 20
Global Const SM_CXHSCROLL = 21
Global Const SM_DEBUG = 22
Global Const SM_SWAPBUTTON = 23
Global Const SM_RESERVED1 = 24
Global Const SM_RESERVED2 = 25
Global Const SM_RESERVED3 = 26
Global Const SM_RESERVED4 = 27
Global Const SM_CXMIN = 28
Global Const SM_CYMIN = 29
Global Const SM_CXSIZE = 30
Global Const SM_CYSIZE = 31
Global Const SM_CXFRAME = 32
Global Const SM_CYFRAME = 33
Global Const SM_CXMINTRACK = 34
Global Const SM_CYMINTRACK = 35
Global Const SM_CXDOUBLECLK = 36
Global Const SM_CYDOUBLECLK = 37
Global Const SM_CXICONSPACING = 38
Global Const SM_CYICONSPACING = 39
Global Const SM_MENUDROPALIGNMENT = 40
Global Const SM_PENWindow = 41
Global Const SM_DBCSENABLED = 42

'System parameters support

Global Const SPI_GETBEEP = 1
Global Const SPI_SETBEEP = 2
Global Const SPI_GETMOUSE = 3
Global Const SPI_SETMOUSE = 4
Global Const SPI_GETBORDER = 5
Global Const SPI_SETBORDER = 6
Global Const SPI_GETKEYBOARDSPEED = 10
Global Const SPI_SETKEYBOARDSPEED = 11
Global Const SPI_LANGDRIVER = 12
Global Const SPI_ICONHORIZONTALSPACING = 13
Global Const SPI_GETSCREENSAVEPause = 14
Global Const SPI_SETSCREENSAVEPause = 15
Global Const SPI_GETSCREENSAVEACTIVE = 16
Global Const SPI_SETSCREENSAVEACTIVE = 17
Global Const SPI_GETGRIDGRANULARITY = 18
Global Const SPI_SETGRIDGRANULARITY = 19
Global Const SPI_SETDESKWALLPAPER = 20
Global Const SPI_SETDESKPATTERN = 21
Global Const SPI_GETKEYBOARDDELAY = 22
Global Const SPI_SETKEYBOARDDELAY = 23
Global Const SPI_ICONVERTICALSPACING = 24
Global Const SPI_GETICONTITLEWRAP = 25
Global Const SPI_SETICONTITLEWRAP = 26
Global Const SPI_GETMENUDROPALIGNMENT = 27
Global Const SPI_SETMENUDROPALIGNMENT = 28
Global Const SPI_SETDOUBLECLKWIDTH = 29
Global Const SPI_SETDOUBLECLKHEIGHT = 30
Global Const SPI_GETICONTITLELOGFONT = 31
Global Const SPI_SETDOUBLECLICKTIME = 32
Global Const SPI_SETMOUSEBUTTONSWAP = 33
Global Const SPI_SETICONTITLELOGFONT = 34
Global Const SPI_GETFASTTASKSWITCH = 35
Global Const SPI_SETFASTTASKSWITCH = 36

'SystemParametersInfo flags
Global Const SPIF_UPDATEINIFILE = 1
Global Const SPIF_SENDWININICHANGE = 2
'MessageBox() Flags
Global Const MB_OK = &H0
Global Const MB_OKCANCEL = &H1
Global Const MB_ABORTRETRYIGNORE = &H2
Global Const MB_YESNOCANCEL = &H3
Global Const MB_YESNO = &H4
Global Const MB_RETRYCANCEL = &H5
Global Const MB_ICONHAND = &H10
Global Const MB_ICONQUESTION = &H20
Global Const MB_ICONEXCLAMATION = &H30
Global Const MB_ICONASTERISK = &H40
Global Const MB_ICONINFORMATION = MB_ICONASTERISK
Global Const MB_ICONSTOP = MB_ICONHAND
Global Const MB_DEFBUTTON1 = &H0
Global Const MB_DEFBUTTON2 = &H100
Global Const MB_DEFBUTTON3 = &H200
Global Const MB_APPLMODAL = &H0
Global Const MB_SYSTEMMODAL = &H1000
Global Const MB_TASKMODAL = &H2000
Global Const MB_NOFOCUS = &H8000
Global Const MB_TYPEMASK = &HF
Global Const MB_ICONMASK = &HF0
Global Const MB_DEFMASK = &HF00
Global Const MB_MODEMASK = &H3000
Global Const MB_MISCMASK = &HC000

'Color Types
Global Const CTLCOLOR_MSGBOX = 0
Global Const CTLCOLOR_EDIT = 1
Global Const CTLCOLOR_LISTBOX = 2
Global Const CTLCOLOR_BTN = 3
Global Const CTLCOLOR_DLG = 4
Global Const CTLCOLOR_SCROLLBAR = 5
Global Const CTLCOLOR_STATIC = 6
Global Const CTLCOLOR_MAX = 8 'three bits max
Global Const COLOR_SCROLLBAR = 0
Global Const COLOR_BACKGROUND = 1
Global Const COLOR_ACTIVECAPTION = 2
Global Const COLOR_INACTIVECAPTION = 3
Global Const COLOR_MENU = 4
Global Const COLOR_WINDOW = 5
Global Const COLOR_WINDOWFRAME = 6
Global Const COLOR_MENUTEXT = 7
Global Const COLOR_WINDOWTEXT = 8
Global Const COLOR_CAPTIONTEXT = 9
Global Const COLOR_ACTIVEBORDER = 10
Global Const COLOR_INACTIVEBORDER = 11
Global Const COLOR_APPWORKSPACE = 12
Global Const COLOR_HIGHLIGHT = 13
Global Const COLOR_HIGHLIGHTTEXT = 14
Global Const COLOR_BTNFACE = 15
Global Const COLOR_BTNSHADOW = 16
Global Const COLOR_GRAYTEXT = 17
Global Const COLOR_BTNTEXT = 18
Global Const COLOR_INACTIVECAPTIONTEXT = 19
Global Const COLOR_BTNHIGHLIGHT = 20

'GetWindow() Constants
Global Const GW_HWNDFIRST = 0
Global Const GW_HWNDLAST = 1
Global Const GW_HWNDNEXT = 2
Global Const GW_HWNDPREV = 3
Global Const GW_OWNER = 4
Global Const GW_CHILD = 5

'GetDCEx flags
Global Const DCX_WINDOW = &H1&
Global Const DCX_CACHE = &H2&
Global Const DCX_CLIPCHILDREN = &H8&
Global Const DCX_CLIPSIBLINGS = &H10&
Global Const DCX_PARENTCLIP = &H20&
Global Const DCX_EXCLUDERGN = &H40&
Global Const DCX_INTERSECTRGN = &H80&
Global Const DCX_LOCKWINDOWUPDATE = &H400&
Global Const DCX_USESTYLE = &H10000

'Menu flags for Add/Check/EnableMenuItem()
Global Const MF_INSERT = &H0
Global Const MF_CHANGE = &H80
Global Const MF_APPEND = &H100
Global Const MF_DELETE = &H200
Global Const MF_REMOVE = &H1000
Global Const MF_BYCOMMAND = &H0
Global Const MF_BYPOSITION = &H400
Global Const MF_SEPARATOR = &H800
Global Const MF_ENABLED = &H0
Global Const MF_GRAYED = &H1
Global Const MF_DISABLED = &H2
Global Const MF_UNCHECKED = &H0
Global Const MF_CHECKED = &H8
Global Const MF_USECHECKBITMAPS = &H200
Global Const MF_STRING = &H0
Global Const MF_BITMAP = &H4
Global Const MF_OWNERDRAW = &H100
Global Const MF_POPUP = &H10
Global Const MF_MENUBARBREAK = &H20
Global Const MF_MENUBREAK = &H40
Global Const MF_UNHILITE = &H0
Global Const MF_HILITE = &H80
Global Const MF_SYSMENU = &H2000
Global Const MF_HELP = &H4000
Global Const MF_MOUSESELECT = &H8000
Global Const MF_END = &H80

'TrackPopupMenu flags
Global Const TPM_LEFTBUTTON = &H0
Global Const TPM_RIGHTBUTTON = &H2
Global Const TPM_LEFTALIGN = &H0
Global Const TPM_CENTERALIGN = &H4
Global Const TPM_RIGHTALIGN = &H8

'System Menu Command Values
Global Const SC_SIZE = &HF000
Global Const SC_MOVE = &HF010
Global Const SC_MINIMIZE = &HF020
Global Const SC_MAXIMIZE = &HF030
Global Const SC_NEXTWINDOW = &HF040
Global Const SC_PREVWINDOW = &HF050
Global Const SC_CLOSE = &HF060
Global Const SC_VSCROLL = &HF070
Global Const SC_HSCROLL = &HF080
Global Const SC_MOUSEMENU = &HF090
Global Const SC_KEYMENU = &HF100
Global Const SC_ARRANGE = &HF110
Global Const SC_RESTORE = &HF120
Global Const SC_TASKLIST = &HF130
Global Const SC_ICON = SC_MINIMIZE
Global Const SC_ZOOM = SC_MAXIMIZE

'Standard Cursor IDs
Global Const IDC_ARROW = 32512&
Global Const IDC_IBEAM = 32513&
Global Const IDC_WAIT = 32514&
Global Const IDC_CROSS = 32515&
Global Const IDC_UPARROW = 32516&
Global Const IDC_SIZE = 32640&
Global Const IDC_ICON = 32641&
Global Const IDC_SIZENWSE = 32642&
Global Const IDC_SIZENESW = 32643&
Global Const IDC_SIZEWE = 32644&
Global Const IDC_SIZENS = 32645&
Global Const ORD_LANGDRIVER = 1

'Standard Icon IDs
Global Const IDI_APPLICATION = 32512&
Global Const IDI_HAND = 32513&
Global Const IDI_QUESTION = 32514&
Global Const IDI_EXCLAMATION = 32515&
Global Const IDI_ASTERISK = 32516&

'Dialog Box Command IDs
Global Const IDOK = 1
Global Const IDCANCEL = 2
Global Const IDABORT = 3
Global Const IDRETRY = 4
Global Const IDIGNORE = 5
Global Const IDYES = 6
Global Const IDNO = 7

'Edit Control Styles
Global Const ES_LEFT = &H0&
Global Const ES_CENTER = &H1&
Global Const ES_RIGHT = &H2&
Global Const ES_MULTILINE = &H4&
Global Const ES_UPPERCASE = &H8&
Global Const ES_LOWERCASE = &H10&
Global Const ES_PASSWORD = &H20&
Global Const ES_AUTOVSCROLL = &H40&
Global Const ES_AUTOHSCROLL = &H80&
Global Const ES_NOHIDESEL = &H100&
Global Const ES_OEMCONVERT = &H400&
Global Const ES_READONLY = &H800&
Global Const ES_WANTRETURN = &H1000&

'Edit Control Notification Codes
Global Const EN_SETFOCUS = &H100
Global Const EN_KILLFOCUS = &H200
Global Const EN_CHANGE = &H300
Global Const EN_UPDATE = &H400
Global Const EN_ERRSPACE = &H500
Global Const EN_MAXTEXT = &H501
Global Const EN_HSCROLL = &H601
Global Const EN_VSCROLL = &H602
Global Const WB_LEFT = 0
Global Const WB_RIGHT = 1
Global Const WB_ISDELIMITER = 2

'Button Control Styles
Global Const BS_PUSHBUTTON = &H0&
Global Const BS_DEFPUSHBUTTON = &H1&
Global Const BS_CHECKBOX = &H2&
Global Const BS_AUTOCHECKBOX = &H3&
Global Const BS_RADIOBUTTON = &H4&
Global Const BS_3STATE = &H5&
Global Const BS_AUTO3STATE = &H6&
Global Const BS_GROUPBOX = &H7&
Global Const BS_USERBUTTON = &H8&
Global Const BS_AUTORADIOBUTTON = &H9&
Global Const BS_PUSHBOX = &HA&
Global Const BS_OWNERDRAW = &HB&
Global Const BS_LEFTTEXT = &H20&

'User Button Notification Codes
Global Const BN_CLICKED = 0
Global Const BN_PAINT = 1
Global Const BN_HILITE = 2
Global Const BN_UNHILITE = 3
Global Const BN_DISABLE = 4
Global Const BN_DOUBLECLICKED = 5

'Static Control Constants
Global Const SS_LEFT = &H0&
Global Const SS_CENTER = &H1&
Global Const SS_RIGHT = &H2&
Global Const SS_ICON = &H3&
Global Const SS_BLACKRECT = &H4&
Global Const SS_GRAYRECT = &H5&
Global Const SS_WHITERECT = &H6&
Global Const SS_BLACKFRAME = &H7&
Global Const SS_GRAYFRAME = &H8&
Global Const SS_WHITEFRAME = &H9&
Global Const SS_USERITEM = &HA&
Global Const SS_SIMPLE = &HB&
Global Const SS_LEFTNOWORDWRAP = &HC&
Global Const SS_NOPREFIX = &H80&


'Dialog Styles
Global Const DS_ABSALIGN = &H1&
Global Const DS_SYSMODAL = &H2&
Global Const DS_LOCALEDIT = &H20&
Global Const DS_SETFONT = &H40&
Global Const DS_MODALFRAME = &H80&
Global Const DS_NOIDLEMSG = &H100&
Global Const DC_HASDEFID = &H534

'Dialog Codes
Global Const DLGC_WANTARROWS = &H1
Global Const DLGC_WANTTAB = &H2
Global Const DLGC_WANTALLKEYS = &H4
Global Const DLGC_WANTMESSAGE = &H4
Global Const DLGC_HASSETSEL = &H8
Global Const DLGC_DEFPUSHBUTTON = &H10
Global Const DLGC_UNDEFPUSHBUTTON = &H20
Global Const DLGC_RADIOBUTTON = &H40
Global Const DLGC_WANTCHARS = &H80
Global Const DLGC_STATIC = &H100
Global Const DLGC_BUTTON = &H2000

'Scroll Bar Styles
Global Const SBS_HORZ = &H0&
Global Const SBS_VERT = &H1&
Global Const SBS_TOPALIGN = &H2&
Global Const SBS_LEFTALIGN = &H2&
Global Const SBS_BOTTOMALIGN = &H4&
Global Const SBS_RIGHTALIGN = &H4&
Global Const SBS_SIZEBOXTOPLEFTALIGN = &H4&
Global Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Global Const SBS_SIZEBOX = &H8&

'WaitSoundState() Constants
Global Const S_QUEUEEMPTY = 0
Global Const S_THRESHOLD = 1
Global Const S_ALLTHRESHOLD = 2

'Accent Modes
Global Const S_NORMAL = 0
Global Const S_LEGATO = 1
Global Const S_STACCATO = 2

'SetSoundNoise() Sources
Global Const S_PERIOD512 = 0'
Global Const S_PERIOD1024 = 1
Global Const S_PERIOD2048 = 2
Global Const S_PERIODVOICE = 3
Global Const S_WHITE512 = 4
Global Const S_WHITE1024 = 5
Global Const S_WHITE2048 = 6
Global Const S_WHITEVOICE = 7
Global Const S_SERDVNA = (-1)
Global Const S_SEROFM = (-2)
Global Const S_SERMACT = (-3)
Global Const S_SERQFUL = (-4)
Global Const S_SERBDNT = (-5)
Global Const S_SERDLN = (-6)
Global Const S_SERDCC = (-7)
Global Const S_SERDTP = (-8)
Global Const S_SERDVL = (-9)
Global Const S_SERDMD = (-10)
Global Const S_SERDSH = (-11)
Global Const S_SERDPT = (-12)
Global Const S_SERDFQ = (-13)
Global Const S_SERDDR = (-14)
Global Const S_SERDSR = (-15)
Global Const S_SERDST = (-16)

'COMM declarations
Global Const NOPARITY = 0
Global Const ODDPARITY = 1
Global Const EVENPARITY = 2
Global Const MARKPARITY = 3
Global Const SPACEPARITY = 4
Global Const ONESTOPBIT = 0
Global Const ONE5STOPBITS = 1
Global Const TWOSTOPBITS = 2
Global Const IGNORE = 0
Global Const INFINITE = &HFFFF

'COMM Error Flags
Global Const CE_RXOVER = &H1
Global Const CE_OVERRUN = &H2
Global Const CE_RXPARITY = &H4
Global Const CE_FRAME = &H8
Global Const CE_BREAK = &H10
Global Const CE_CTSTO = &H20
Global Const CE_DSRTO = &H40
Global Const CE_RLSDTO = &H80
Global Const CE_TXFULL = &H100
Global Const CE_PTO = &H200
Global Const CE_IOE = &H400
Global Const CE_DNS = &H800
Global Const CE_OOP = &H1000
Global Const CE_MODE = &H8000
Global Const IE_BADID = (-1)
Global Const IE_OPEN = (-2)
Global Const IE_NOPEN = (-3)
Global Const IE_MEMORY = (-4)
Global Const IE_DEFAULT = (-5)
Global Const IE_HARDWARE = (-10)
Global Const IE_BYTESIZE = (-11)
Global Const IE_BAUDRATE = (-12)

'COMM Events
Global Const EV_RXCHAR = &H1
Global Const EV_RXFLAG = &H2
Global Const EV_TXEMPTY = &H4
Global Const EV_CTS = &H8
Global Const EV_DSR = &H10
Global Const EV_RLSD = &H20
Global Const EV_BREAK = &H40
Global Const EV_ERR = &H80
Global Const EV_RING = &H100
Global Const EV_PERR = &H200
Global Const EV_CTSS = &H400
Global Const EV_DSRS = &H800
Global Const EV_RLSDS = &H1000


'COMM Escape Functions
Global Const SETXOFF = 1'Simulate XOFF received
Global Const SETXON = 2 'Simulate XON received
Global Const SETRTS = 3 'Set RTS high
Global Const CLRRTS = 4 'Set RTS low
Global Const SETDTR = 5 'Set DTR high
Global Const CLRDTR = 6 'Set DTR low
Global Const RESETDEV = 7 'Reset device if possible
Global Const GETMAXLPT = 8
Global Const GETMAXCOM = 9
Global Const GETBASEIRQ = 10
Global Const CBR_110 = &HFF10
Global Const CBR_300 = &HFF11
Global Const CBR_600 = &HFF12
Global Const CBR_1200 = &HFF13
Global Const CBR_2400 = &HFF14
Global Const CBR_4800 = &HFF15
Global Const CBR_9600 = &HFF16
Global Const CBR_14400 = &HFF17
Global Const CBR_19200 = &HFF18
Global Const CBR_38400 = &HFF1B
Global Const CBR_56000 = &HFF1F
Global Const CBR_128000 = &HFF23
Global Const CBR_256000 = &HFF27

'COMM notifications on WM_COMMNOTIFY messages
Global Const CN_RECEIVE = &H1
Global Const CN_TRANSMIT = &H2
Global Const CN_EVENT = &H4

'COMM status flags
Global Const CSTF_CTSHOLD = &H1
Global Const CSTF_DSRHOLD = &H2
Global Const CSTF_RLSDHOLD = &H4
Global Const CSTF_XOFFHOLD = &H8
Global Const CSTF_XOFFSENT = &H10
Global Const CSTF_EOF = &H20
Global Const CSTF_TXIM = &H40
Global Const LPTx = &H80

'Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1
Global Const HELP_QUIT = &H2
Global Const HELP_INDEX = &H3
Global Const HELP_HELPONHELP = &H4
Global Const HELP_SETINDEX = &H5
Global Const HELP_CONTEXTPOPUP = &H8
Global Const HELP_FORCEFILE = &H9
Global Const HELP_KEY = &H101
Global Const HELP_COMMAND = &H102
Global Const HELP_PARTIALKEY = &H105
Global Const HELP_MULTIKEY = &H201
Global Const HELP_SETWINPOS = &H203

'Field selection bits
Global Const DM_ORIENTATION = &H1&
Global Const DM_PAPERSIZE = &H2&
Global Const DM_PAPERLENGTH = &H4&
Global Const DM_PAPERWIDTH = &H8&
Global Const DM_SCALE = &H10&
Global Const DM_COPIES = &H100&
Global Const DM_DEFAULTSOURCE = &H200&
Global Const DM_PRINTQUALITY = &H400&
Global Const DM_COLOR = &H800&
Global Const DM_DUPLEX = &H1000&
Global Const DM_YRESOLUTION = &H2000&
Global Const DM_TTOPTION = &H4000&

'Printer orientation selections
Global Const DMORIENT_PORTRAIT = 1
Global Const DMORIENT_LANDSCAPE = 2

'Paper selections
Global Const DMPAPER_LETTER = 1
Global Const DMPAPER_LETTERSMALL = 2
Global Const DMPAPER_TABLOID = 3
Global Const DMPAPER_LEDGER = 4
Global Const DMPAPER_LEGAL = 5
Global Const DMPAPER_STATEMENT = 6
Global Const DMPAPER_EXECUTIVE = 7
Global Const DMPAPER_A3 = 8
Global Const DMPAPER_A4 = 9
Global Const DMPAPER_A4SMALL = 10
Global Const DMPAPER_A5 = 11
Global Const DMPAPER_B4 = 12
Global Const DMPAPER_B5 = 13
Global Const DMPAPER_FOLIO = 14
Global Const DMPAPER_QUARTO = 15
Global Const DMPAPER_10X14 = 16
Global Const DMPAPER_11X17 = 17
Global Const DMPAPER_NOTE = 18
Global Const DMPAPER_ENV_9 = 19
Global Const DMPAPER_ENV_10 = 20
Global Const DMPAPER_ENV_11 = 21
Global Const DMPAPER_ENV_12 = 22
Global Const DMPAPER_ENV_14 = 23
Global Const DMPAPER_CSHEET = 24
Global Const DMPAPER_DSHEET = 25
Global Const DMPAPER_ESHEET = 26
Global Const DMPAPER_ENV_DL = 27
Global Const DMPAPER_ENV_C5 = 28
Global Const DMPAPER_ENV_C3 = 29
Global Const DMPAPER_ENV_C4 = 30
Global Const DMPAPER_ENV_C6 = 31
Global Const DMPAPER_ENV_C65 = 32
Global Const DMPAPER_ENV_B4 = 33
Global Const DMPAPER_ENV_B5 = 34
Global Const DMPAPER_ENV_B6 = 35
Global Const DMPAPER_ENV_ITALY = 36
Global Const DMPAPER_ENV_MONARCH = 37
Global Const DMPAPER_ENV_PERSONAL = 38
Global Const DMPAPER_FANFOLD_US = 39
Global Const DMPAPER_FANFOLD_STD_GERMAN = 40
Global Const DMPAPER_FANFOLD_LGL_GERMAN = 41
Global Const DMPAPER_USER = 256

'Printer bin selections
Global Const DMBIN_UPPER = 1
Global Const DMBIN_ONLYONE = 1
Global Const DMBIN_LOWER = 2
Global Const DMBIN_MIDDLE = 3
Global Const DMBIN_MANUAL = 4
Global Const DMBIN_ENVELOPE = 5
Global Const DMBIN_ENVMANUAL = 6
Global Const DMBIN_AUTO = 7
Global Const DMBIN_TRACTOR = 8
Global Const DMBIN_SMALLFMT = 9
Global Const DMBIN_LARGEFMT = 10
Global Const DMBIN_LARGECAPACITY = 11
Global Const DMBIN_CASSETTE = 14
Global Const DMBIN_USER = 256

'Print qualities
Global Const DMRES_DRAFT = -1
Global Const DMRES_LOW = -2
Global Const DMRES_MEDIUM = -3
Global Const DMRES_HIGH = -4

'Color enable/disable for color printers
Global Const DMCOLOR_MONOCHROME = 1
Global Const DMCOLOR_COLOR = 2

'Printer duplex enable
Global Const DMDUP_SIMPLEX = 1
Global Const DMDUP_VERTICAL = 2
Global Const DMDUP_HORIZONTAL = 3

'TrueType options
Global Const DMTT_BITMAP = 1
Global Const DMTT_DOWNLOAD = 2
Global Const DMTT_SUBDEV = 3

'Device mode function modes
Global Const DM_UPDATE = 1
Global Const DM_COPY = 2
Global Const DM_PROMPT = 4
Global Const DM_MODIFY = 8
Global Const DM_IN_BUFFER = 8
Global Const DM_IN_PROMPT = 4
Global Const DM_OUT_BUFFER = 2
Global Const DM_OUT_DEFAULT = 1

'Device capabilities indices
Global Const DC_FIELDS = 1
Global Const DC_PAPERS = 2
Global Const DC_PAPERSIZE = 3
Global Const DC_MINEXTENT = 4
Global Const DC_MAXEXTENT = 5
Global Const DC_BINS = 6
Global Const DC_DUPLEX = 7
Global Const DC_SIZE = 8
Global Const DC_EXTRA = 9
Global Const DC_VERSION = 10
Global Const DC_DRIVER = 11
Global Const DC_BINNAMES = 12
Global Const DC_ENUMRESOLUTIONS = 13
Global Const DC_FILEDEPENDENCIES = 14
Global Const DC_TRUETYPE = 15
Global Const DC_PAPERNAMES = 16
Global Const DC_ORIENTATION = 17
Global Const DC_COPIES = 18

'DC_TRUETYPE bit fields
Global Const DCTT_BITMAP = &H1&
Global Const DCTT_DOWNLOAD = &H2&
Global Const DCTT_SUBDEV = &H4&

'LZ encode constants
Global Const LZERROR_BADINHANDLE = -1
Global Const LZERROR_BADOUTHANDLE = -2
Global Const LZERROR_READ = -3
Global Const LZERROR_WRITE = -4
Global Const LZERROR_GLOBALLOC = -5
Global Const LZERROR_GLOBLOCK = -6
Global Const LZERROR_BADVALUE = -7
Global Const LZERROR_UNKNOWNALG = -8

'Version Control Resources
Global Const VS_FILE_INFO = 16
Global Const VS_VERSION_INFO = 1
Global Const VS_USER_DEFINED = 100

'Version control flags
Global Const VS_FFI_SIGNATURE = &HFEEF04BD
Global Const VS_FFI_STRUCVERSION = &H10000
Global Const VS_FFI_FILEFLAGSMASK = &H3F&
Global Const VS_FF_DEBUG = &H1&
Global Const VS_FF_PRERELEASE = &H2&
Global Const VS_FF_PATCHED = &H4&
Global Const VS_FF_PRIVATEBUILD = &H8&
Global Const VS_FF_INFOINFERRED = &H10&
Global Const VS_FF_SPECIALBUILD = &H20&

'Version control OS flags
Global Const VOS_UNKNOWN = &H0&
Global Const VOS_DOS = &H10000
Global Const VOS_OS216 = &H20000
Global Const VOS_OS232 = &H30000
Global Const VOS_NT = &H40000
Global Const VOS__BASE = &H0&
Global Const VOS__Window16 = &H1&
Global Const VOS__PM16 = &H2&
Global Const VOS__PM32 = &H3&
Global Const VOS__Window32 = &H4&
Global Const VOS_DOS_Window16 = &H10001
Global Const VOS_DOS_Window32 = &H10004
Global Const VOS_OS216_PM16 = &H20002
Global Const VOS_OS232_PM32 = &H30003
Global Const VOS_NT_Window32 = &H40004

'Version control file types
Global Const VFT_UNKNOWN = &H0&
Global Const VFT_APP = &H1&
Global Const VFT_DLL = &H2&
Global Const VFT_DRV = &H3&
Global Const VFT_FONT = &H4&
Global Const VFT_VXD = &H5&
Global Const VFT_STATIC_LIB = &H7&

' VS_VERSION.dwFileSubtype for VFT_Window_DRV
Global Const VFT2_UNKNOWN = &H0&
Global Const VFT2_DRV_PRINTER = &H1&
Global Const VFT2_DRV_KEYBOARD = &H2&
Global Const VFT2_DRV_LANGUAGE = &H3&
Global Const VFT2_DRV_DISPLAY = &H4&
Global Const VFT2_DRV_MOUSE = &H5&
Global Const VFT2_DRV_NETWORK = &H6&
Global Const VFT2_DRV_SYSTEM = &H7&
Global Const VFT2_DRV_INSTALLABLE = &H8&
Global Const VFT2_DRV_SOUND = &H9&
Global Const VFT2_DRV_COMM = &HA&

' VS_VERSION.dwFileSubtype for VFT_Window_FONT
Global Const VFT2_FONT_RASTER = &H1&
Global Const VFT2_FONT_VECTOR = &H2&
Global Const VFT2_FONT_TRUETYPE = &H3&

'VerFindFile() flags
Global Const VFFF_ISSHAREDFILE = &H1
Global Const VFF_CURNEDEST = &H1
Global Const VFF_FILEINUSE = &H2
Global Const VFF_BUFFTOOSMALL = &H4

'VerInstallFile() flags
Global Const VIFF_FORCEINSTALL = &H1
Global Const VIFF_DONTDELETEOLD = &H2
Global Const VIF_TEMPFILE = &H1&
Global Const VIF_MISMATCH = &H2&
Global Const VIF_SRCOLD = &H4&
Global Const VIF_DIFFLANG = &H8&
Global Const VIF_DIFFCODEPG = &H10&
Global Const VIF_DIFFTYPE = &H20&
Global Const VIF_WRITEPROT = &H40&
Global Const VIF_FILEINUSE = &H80&
Global Const VIF_OUTOFSPACE = &H100&
Global Const VIF_ACCESSVIOLATION = &H200&
Global Const VIF_SHARINGVIOLATION = &H400&
Global Const VIF_CANNOTCREATE = &H800&
Global Const VIF_CANNOTDELETE = &H1000&
Global Const VIF_CANNOTRENAME = &H2000&
Global Const VIF_CANNOTDELETECUR = &H4000&
Global Const VIF_OUTOFMEMORY = &H8000&
Global Const VIF_CANNOTREADSRC = &H10000
Global Const VIF_CANNOTREADDST = &H20000
Global Const VIF_BUFFTOOSMALL = &H40000

'WM_POWER window message and DRV_POWER driver notification
Global Const PWR_OK = 1
Global Const PWR_FAIL = (-1)
Global Const PWR_SUSPENDREQUEST = 1
Global Const PWR_SUSPENDRESUME = 2
Global Const PWR_CRITICALRESUME = 3

'Network operation return values
Global Const WN_SUCCESS = 0
Global Const WN_NOT_SUPPORTED = 1
Global Const WN_NET_ERROR = 2
Global Const WN_MORE_DATA = 3
Global Const WN_BAD_POINTER = 4
Global Const WN_BAD_VALUE = 5
Global Const WN_BAD_PASSWORD = 6
Global Const WN_ACCESS_DENIED = 7
Global Const WN_FUNCTION_BUSY = 8
Global Const WN_Window_ERROR = 9
Global Const WN_BAD_USER = &HA
Global Const WN_OUT_OF_MEMORY = &HB
Global Const WN_CANCEL = &HC
Global Const WN_CONTINUE = &HD

'Network Connection errors
Global Const WN_NOT_CONNECTED = &H30
Global Const WN_OPEN_FILES = &H31
Global Const WN_BAD_NETNAME = &H32
Global Const WN_BAD_LOCALNAME = &H33
Global Const WN_ALREADY_CONNECTED = &H34
Global Const WN_DEVICE_ERROR = &H35
Global Const WN_CONNECTION_CLOSED = &H36
Global XSound As String
Global Info(1 To 5) As String
Global ClickNum As Integer
Global DialogCaption As String
Global Trk
Global TotalTrk
Global Flag
Global NewCount
Global CNCL As Integer

Global Const LB_GETCOUNT = (WM_USER + 12)
Declare Function CreatePopupMenu Lib "User" () As Integer
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Global Const WM_SETTEXT = &HC
Global Const WM_KEYDOWN = &H100
Global Const WM_KEYUP = &H101
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cX As Integer, ByVal cY As Integer, ByVal wFlags As Integer)
Declare Sub GetCursorPos Lib "User" (lpPoint As Long)
Declare Function DefWindowProc Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lparam As Any) As Long
Declare Function CreateCompatibleDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Function GetWindowDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function DeleteDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Function ReleaseDC Lib "User" (ByVal hWnd As Integer, ByVal hDC As Integer) As Integer
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function LoadBitmap Lib "User" (ByVal hInstance As Integer, ByVal lpBitmapName As Any) As Integer
Declare Function APIGetObject Lib "GDI" Alias "GetObject" (ByVal hObject As Integer, ByVal nCount As Integer, lpObject As Any) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Function GetSystemMetrics Lib "User" (ByVal nIndex As Integer) As Integer
Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, lpRect As Any)

'MAPI Globals
Global Const RECIPTYPE_ORIG = 0
Global Const RECIPTYPE_TO = 1
Global Const RECIPTYPE_CC = 2
Global Const RECIPTYPE_BCC = 3
Global Const ATTACHTYPE_DATA = 0
Global Const ATTACHTYPE_EOLE = 1
Global Const ATTACHTYPE_SOLE = 2
'Action
Global Const MESSAGE_FETCH = 1             ' Load all messages from message store
Global Const MESSAGE_SENDDLG = 2           ' Send mail bring up default mapi dialog
Global Const MESSAGE_SEND = 3              ' Send mail without default mapi dialog
Global Const MESSAGE_SAVEMSG = 4           ' Save message in the compose buffer
Global Const MESSAGE_COPY = 5              ' Copy current message to compose buffer
Global Const MESSAGE_COMPOSE = 6           ' Initialize compose buffer (previous
					   ' data is lostGlobal Const MESSAGE_REPLY = 7
					   ' Fill Compose buffer as REPLY
Global Const MESSAGE_REPLYALL = 8          ' Fill Compose buffer as REPLY ALL
Global Const MESSAGE_FORWARD = 9           ' Fill Compose buffer as FORWARD
Global Const MESSAGE_DELETE = 10           ' Delete current message
Global Const MESSAGE_SHOWADBOOK = 11       ' Show Address book
Global Const MESSAGE_SHOWDETAILS = 12      ' Show details of the current recipient
Global Const MESSAGE_RESOLVENAME = 13      ' Resolve the display name of the recipient
Global Const RECIPIENT_DELETE = 14            ' Fill Compose buffer as FORWARD
Global Const ATTACHMENT_DELETE = 15          ' Delete current message
'MAPI Errors
Global Const SUCCESS_SUCCESS = 32000
Global Const MAPI_USER_ABORT = 32001
Global Const MAPI_E_FAILURE = 32002
Global Const MAPI_E_LOGIN_FAILURE = 32003
Global Const MAPI_E_DISK_FULL = 32004
Global Const MAPI_E_INSUFFICIENT_MEMORY = 32005
Global Const MAPI_E_ACCESS_DENIED = 32006
Global Const MAPI_E_TOO_MANY_SESSIONS = 32008
Global Const MAPI_E_TOO_MANY_FILES = 32009
Global Const MAPI_E_TOO_MANY_RECIPIENTS = 32010
Global Const MAPI_E_ATTACHMENT_NOT_FOUND = 32011
Global Const MAPI_E_ATTACHMENT_OPEN_FAILURE = 32012
Global Const MAPI_E_ATTACHMENT_WRITE_FAILURE = 32013
Global Const MAPI_E_UNKNOWN_RECIPIENT = 32014
Global Const MAPI_E_BAD_RECIPTYPE = 32015
Global Const MAPI_E_NO_MESSAGES = 32016
Global Const MAPI_E_INVALID_MESSAGE = 32017
Global Const MAPI_E_TEXT_TOO_LARGE = 32018
Global Const MAPI_E_INVALID_SESSION = 32019
Global Const MAPI_E_TYPE_NOT_SUPPORTED = 32020
Global Const MAPI_E_AMBIGUOUS_RECIPIENT = 32021
Global Const MAPI_E_MESSAGE_IN_USE = 32022
Global Const MAPI_E_NETWORK_FAILURE = 32023
Global Const MAPI_E_INVALID_EDITFIELDS = 32024
Global Const MAPI_E_INVALID_RECIPS = 32025
Global Const MAPI_E_NOT_SUPPORTED = 32026

Global Const CONTROL_E_SESSION_EXISTS = 32050
Global Const CONTROL_E_INVALID_BUFFER = 32051
Global Const CONTROL_E_INVALID_READ_BUFFER_ACTION = 32052
Global Const CONTROL_E_NO_SESSION = 32053
Global Const CONTROL_E_INVALID_RECIPIENT = 32054
Global Const CONTROL_E_INVALID_COMPOSE_BUFFER_ACTION = 32055
Global Const CONTROL_E_FAILURE = 32056
Global Const CONTROL_E_NO_RECIPIENTS = 32057
Global Const CONTROL_E_NO_ATTACHMENTS = 32058
'MAPI Session
Global Const SESSION_SIGNON = 1
Global Const SESSION_SIGNOFF = 2

'NON Rectangular Forms
Declare Function CreatePolygonRgn Lib "GDI" (lpPoints As POINTAPI, ByVal nCount As Integer, ByVal nPolyFillMode As Integer) As Integer
Declare Function CreateEllipticRgn Lib "GDI" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
Declare Function PaintRgn Lib "GDI" (ByVal hDC As Integer, ByVal hRgn As Integer) As Integer

'Read Only Text Boxes
Global Const EM_SETREADONLY = (WM_USER + 31)

Sub AC_SignOff ()
If AC_AOLVersion() = 25 Then
    AC_Close FindWindow("AOL Frame25", 0&)
    Do Until sOff% <> 0
	DoEvents
	sOff% = FindWindow("_AOL_MODAL", 0&)
	Timeout .001
	Loop
    yBttn% = FindChildByTitle(sOff%, "&Yes")
    AC_Click yBttn%
   Else
    AC_Close FindWindow("AOL Frame25", 0&)
    End If
End Sub

Sub AC_TurboBustSetup (ByVal Caption As String, ByVal Room As String)
'This edits the Goto.INI via the _AOL_MODAL so you can
'create a quick room buster with the tenth item on that
'menu.

AC_RunMenuByString "Edit Go To Menu", "&Go To"
Do Until AMod% <> 0
    DoEvents
    AMod% = FindWindow("_AOL_MODAL", "Favorite Places")
    Timeout (.001)
    Loop
Q% = ShowWindow(AMod%, SW_HIDE)
SaveBttn% = FindChildByTitle(AMod%, "Save Changes")
NameEd% = AC_GetAOLWin(AMod%, "_AOL_EDIT", 19)
URLEd% = AC_GetAOLWin(AMod%, "_AOL_EDIT", 20)
AC_SetText NameEd%, Caption$
AC_SetText URLEd%, "aol://2719:2-2-" & Room$
AC_Click SaveBttn%
End Sub

Sub AC_UnUpChat ()
AOM% = FindWindow("_AOL_MODAL", 0&)
If Left$(AC_GetWinText(AOM%), 4) <> "File" Then MsgBox "No file transfer detected.", 0, "Error": Exit Sub
MoveWindow AOM%, 0, 0, 370, 180, -1
Dim rRect As Rect
AC_Hide AOM%
GetWindowRect AOM%, rRect
AF_CenterAny AOM%
AC_Show AOM%
Q% = SetFocusAPI(AOM%)
DoEvents
End Sub

Sub AC_UpChat ()
AOM% = FindWindow("_AOL_MODAL", 0&)
If Left$(AC_GetWinText(AOM%), 4) <> "File" Then MsgBox "No file transfer detected.", 0, "Error": Exit Sub
MoveWindow AOM%, 0, 0, 160, -95, 1
AC_KillModal
End Sub

Sub AC_WaitForChange ()
Old% = GetFocus()
Boy% = GetFocus()
Do Until Boy% <> Old%
    Q% = DoEvents()
    Boy% = GetFocus()
    Loop
End Sub

Sub AC_WaitForChangeHide ()
Old% = GetFocus()
Boy% = GetFocus()
Do Until Boy% <> Old%
    x = DoEvents()
    Boy% = GetFocus()
    Loop
Q% = ShowWindow(GetParent(Boy%), SW_HIDE)
DoEvents
End Sub

Sub AC_WaitForCursor ()
'This sub will pause your code until the cursor is no
'longer the hourglass.  Win95 only.
CurType = GetCursor()
Do Until CurType <> 5094
     CurType = GetCursor()
     Loop
End Sub

Sub AC_WaitForOK ()
Do Until OKhWnd% <> 0
    DoEvents
    OKhWnd% = FindWindow("#32770", "America Online")
    Timeout (.001)
    Loop
K% = FindChildByTitle(OKhWnd%, "OK")
AC_Click K%
End Sub

Sub AF_AbstractWindow (FRM As Form)

ReDim myPolygon(5) As POINTAPI
Dim myPolyPolygon() As POINTAPI
Dim myPolyCount() As Integer
Dim hPPRegion As Integer, hTRegion As Integer
Dim PPCount%
Dim i%, j%
myPolygon(0).x = 0:  myPolygon(0).y = 0
myPolygon(1).x = 0:  myPolygon(1).y = 120
myPolygon(2).x = 284:    myPolygon(2).y = 120
myPolygon(3).x = 284:    myPolygon(3).y = 29
myPolygon(4).x = 260:    myPolygon(4).y = 0
poly = CreatePolygonRgn(myPolygon(0), 5, WINDING)
butt = CreateEllipticRgn(0, 0, 10, 10)
    
Q& = PaintRgn(FRM.hWnd, poly)
End Sub

Sub AF_AlignBottomCenter (FRM As Form)
moveLeft = (screen.Width - FRM.Width) / 2
moveTop = (screen.Height - FRM.Height)
FRM.Move moveLeft, moveTop
End Sub

Sub AF_AlignLowerLeft (FRM As Form)
moveValue = screen.Height - FRM.Height
FRM.Move 0, moveValue
End Sub

Sub AF_AlignLowerRight (FRM As Form)
moveLeft = screen.Width - FRM.Width
moveTop = screen.Height - FRM.Height
FRM.Move moveLeft, moveTop
End Sub

Sub AF_AlignTopCenter (FRM As Form)
moveLeft = (screen.Width - FRM.Height) / 2
FRM.Move moveLeft, 0
End Sub

Sub AF_AlignUpperLeft (FRM As Form)
FRM.Move 0, 0
End Sub

Sub AF_AlignUpperRight (FRM As Form)
moveValue = screen.Width - FRM.Width
FRM.Move moveValue, 0
End Sub

Function AF_AppDupeCheck () As Integer
'If AF_AppDupeCheck() = True Then End --- used for ending
'the program if it is already loaded.

If app.PrevInstance = True Then MsgBox "You cannot have more than one of the specified program running at once.", 0, "Error": AF_AppDupeCheck% = True Else AF_AppDupeCheck% = False
End Function

Sub AF_Apply3D (myForm As Form, MyCtl As Control)
    'Place in Form|Paint; worx best with Grey background
    myForm.ScaleMode = 3
    myForm.CurrentX = MyCtl.Left - 1
    myForm.CurrentY = MyCtl.Top + MyCtl.Height
    myForm.Line -Step(0, -(MyCtl.Height + 1)), RGB(92, 92, 92)
    myForm.Line -Step(MyCtl.Width + 1, 0), RGB(92, 92, 92)
    myForm.Line -Step(0, MyCtl.Height + 1), RGB(255, 255, 255)
    myForm.Line -Step(-(MyCtl.Width + 1), 0), RGB(255, 255, 255)
End Sub

Function AF_Bracket (ByVal S) As String
AF_Bracket$ = CStr("•–^v–ª[" & S & "]ª–v^–•")
End Function

Sub AF_Center (FRM As Form)
Dim abc As Variant
Dim xyz As Variant
    abc = (screen.Width - FRM.Width) / 2
    xyz = (screen.Height - FRM.Height) / 2
    FRM.Move abc, xyz
End Sub

Sub AF_CenterAndTop (FRM As Form)
Dim abc As Variant
Dim xyz As Variant
    abc = (screen.Width - FRM.Width) / 2
    xyz = (screen.Height - FRM.Height) / 2
    FRM.Move abc, xyz
SetWindowPos FRM.hWnd, -1, 0, 0, 0, 0, &H50
End Sub

Sub AF_CenterAny (ByVal chWnd As Integer)
'This sub will center any non-VB window you tell it to.
Dim cRect As Rect
GetWindowRect chWnd%, cRect
cX = ((screen.Width / screen.TwipsPerPixelY) - (cRect.Right - cRect.Left))
cY = ((screen.Height / screen.TwipsPerPixelX) - (cRect.Bottom - cRect.Top))
MoveWindow chWnd%, cX / 2, cY / 2, (cRect.Right - cRect.Left), (cRect.Bottom - cRect.Top), 0
End Sub

Function AF_Enter ()
AF_Enter = CStr(Chr(13) & Chr(10))
End Function

Sub AF_Explode (FRM As Form, CFlag As Integer, Steps As Integer)
'This is awesome... put it in Form_Load
Dim FRect As Rect
Dim FWidth, fHeight As Integer
Dim i, x, y, cX, cY As Integer
Dim hScreen, Brush As Integer, OldBrush
' If CFlag = True, then explode from center of form, otherwise
' explode from upper left corner.
    GetWindowRect FRM.hWnd, FRect
    FWidth = (FRect.Right - FRect.Left)
    fHeight = FRect.Bottom - FRect.Top
' Create brush with Form's background color.
    hScreen = GetDC(0)
    Brush = CreateSolidBrush(FRM.BackColor)
    OldBrush = SelectObject(hScreen, Brush)
' Draw rectangles in larger sizes filling in the area to be occupied
' by the form.
    For i = 1 To Steps
	cX = FWidth * (i / Steps)
	cY = fHeight * (i / Steps)
	If CFlag Then
	    x = FRect.Left + (FWidth - cX) / 2
	    y = FRect.Top + (fHeight - cY) / 2
	Else
	    x = FRect.Left
	    y = FRect.Top
	End If
	Rectangle hScreen, x, y, x + cX, y + cY
    Next i
' Release the device context to free memory.
' Make the Form visible
    If ReleaseDC(0, hScreen) = 0 Then
	MsgBox "Unable to Release Device Context", 16, "Device Error"
    End If
    DeleteObject (Brush)
    FRM.Show
End Sub

Sub AF_ExplodeAddX (FRM As Form, Steps As Integer, ByVal nWidth As Long)
'This sub will explode a section of the form to expand
'horizontally.  Pretty cool; written by Anubis.
'FRM is the Form.Name, Steps is the speed and nWidth is
'the new Width the form will be.
Dim FRect As Rect
Dim i, x, y, cX, cY As Integer
Dim hScreen, Brush, OldBrush As Integer
GetWindowRect FRM.hWnd, FRect
wVal = (FRM.Width / screen.TwipsPerPixelY)
nHeight2 = (FRM.Height / screen.TwipsPerPixelY)
nWidth2 = (nWidth / screen.TwipsPerPixelX) - wVal
hScreen = GetDC(0)
Brush = CreateSolidBrush(&H0)
OldBrush = SelectObject(hScreen, Brush)
For i = 1 To Steps
    DoEvents
    cX = nWidth2 * (i / Steps)
    cY = nHeight2 * (i / Steps)
    x = FRect.Right
    y = FRect.Top
    Rectangle hScreen, x, y, x + cX, y + cY
    Next i
DeleteObject (Brush)
FRM.Width = nWidth
FRM.Refresh
End Sub

Sub AF_ExplodeAddY (FRM As Form, Steps As Integer, ByVal nHeight As Long)
'This sub will explode a section of the form to expand
'vertically.  Pretty cool; written by Anubis.
'FRM is the Form.Name, Steps is the speed and nHeight is
'the new Height the form will be.
Dim FRect As Rect
Dim i, x, y, cX, cY As Integer
Dim hScreen, Brush, OldBrush As Integer
GetWindowRect FRM.hWnd, FRect
hVal = (FRM.Height / screen.TwipsPerPixelX)
nHeight2 = (nHeight / screen.TwipsPerPixelY) - hVal
nWidth2 = (FRM.Width / screen.TwipsPerPixelX)
hScreen = GetDC(0)
Brush = CreateSolidBrush(&H0)
OldBrush = SelectObject(hScreen, Brush)
For i = 1 To Steps
    DoEvents
    cX = nWidth2 * (i / Steps)
    cY = nHeight2 * (i / Steps)
    x = FRect.Left
    y = FRect.Bottom
    Rectangle hScreen, x, y, x + cX, y + cY
    Next i
DeleteObject (Brush)
FRM.Height = nHeight
FRM.Refresh
End Sub

Sub AF_ExplodeX (FRM As Form, Steps As Integer)
'This Sub explodes the form horizontally; by Anubis.
'Steps is the speed.
Dim FRect As Rect
Dim i, x, y, cX, cY As Integer
Dim hScreen, Brush, OldBrush As Integer
GetWindowRect FRM.hWnd, FRect
nHeight = (FRM.Height / screen.TwipsPerPixelY)
nWidth = (FRM.Width / screen.TwipsPerPixelX)
hScreen = GetDC(0)
Brush = CreateSolidBrush(&H0)
OldBrush = SelectObject(hScreen, Brush)
For i = 1 To Steps
    DoEvents
    cX = nWidth * (i / Steps)
    cY = nHeight
    x = FRect.Left
    y = FRect.Top
    Rectangle hScreen, x, y, x + cX, y + cY
    Next i
DeleteObject (Brush)
FRM.Show
End Sub

Sub AF_ExplodeY (FRM As Form, Steps As Integer)
'This Sub explodes the form horizontally; by Anubis.
'Steps is the speed.
Dim FRect As Rect
Dim i, x, y, cX, cY As Integer
Dim hScreen, Brush, OldBrush As Integer
GetWindowRect FRM.hWnd, FRect
nHeight = (FRM.Height / screen.TwipsPerPixelY)
nWidth = (FRM.Width / screen.TwipsPerPixelX)
hScreen = GetDC(0)
Brush = CreateSolidBrush(&H0)
OldBrush = SelectObject(hScreen, Brush)
For i = 1 To Steps
    DoEvents
    cX = nWidth
    cY = nHeight * (i / Steps)
    x = FRect.Left
    y = FRect.Top
    Rectangle hScreen, x, y, x + cX, y + cY
    Next i
DeleteObject (Brush)
FRM.Show
End Sub

Sub AF_FadeIn (Lab As Label)
Lab.ForeColor = &HC0C0C0
Timeout (.3)
Lab.ForeColor = &H808080
Timeout (.3)
Lab.ForeColor = &H0&
End Sub

Sub AF_FadeOut (Lab As Label)
Lab.ForeColor = &H0&
Timeout (.3)
Lab.ForeColor = &H808080
Timeout (.3)
Lab.ForeColor = &HC0C0C0
End Sub

Function AF_KillNull (ByVal TXT As String) As String
'This function removes all spaces from a string, including
'ending, begining and all spaces inbetween.
Kill_1$ = Trim$(AF_Script(TXT$, " ", ""))
AF_KillNull$ = AF_Script(Kill_1$, Chr(160), "")
End Function

Sub AF_MoveForm (hWnd As Form)
'This will move a form that has no title just by clicking
'and dragging.  Place in "Form_MouseDown"

Dim mpos As Long
Dim p As ConvertPointAPI
Dim ret As Integer
Call GetCursorPos(mpos)
ret = SendMessage(hWnd.hWnd, WM_LBUTTONUP, 0, mpos)
ret = SendMessage(hWnd.hWnd, WM_SYSCOMMAND, MOUSE_MOVE, mpos)
DoEvents
End Sub

Sub AF_OnBottom (FRM As Form)
'This will set a form always on the bottom.
SetWindowPos FRM.hWnd, 1, 0, 0, 0, 0, &H50
End Sub

Sub AF_PlayWav (ByVal XSound As String)

End Sub

Function AF_Quote (ByVal TXT) As String
AF_Quote$ = CStr(Chr(34) & TXT & Chr(34))
End Function

Function AF_Range (Lower As Integer, Upper As Integer) As Integer
Randomize
AF_Range% = Int((Upper% - Lower% + 1) * Rnd + Lower%)
End Function

Sub AF_ReadListINI (ByVal INIFile As String, Lst As Control)
'This is a simple Sub that will read a INI file into a
'list or combobox control.  Use AF_WriteListINI to write
'to an INI file.
If AC_FileExist(INIFile$) = False Then Exit Sub
Open INIFile$ For Input As #1
    Do Until EOF(1)
	Input #1, n$
	If n$ <> "" Then Lst.AddItem n$
	Loop
    Close #1
End Sub

Sub AF_RemDupes (Lst As ListBox)
For i = 0 To Lst.ListCount - 1
    For n = 0 To Lst.ListCount - 1
	If LCase(Lst.List(i)) Like LCase(Lst.List(n)) And i <> n Then Lst.RemoveItem (n)
	Next n
    Next i
End Sub

Function AF_Script (ByVal Strt As String, ByVal ReplaceMe As String, ByVal ReplaceWith As String) As String
Start$ = Strt
Do While InStr(Start$, ReplaceMe$) <> 0
    x% = DoEvents()
    pos% = InStr(Start$, ReplaceMe$)
    Start$ = Left(Start$, pos% - 1) & ReplaceWith$ & Right(Start$, Len(Start$) - pos% - Len(ReplaceMe$) + 1)
    Loop
AF_Script$ = Start$
End Function

Sub AF_StayOnTop (FRM As Form)
SetWindowPos FRM.hWnd, -1, 0, 0, 0, 0, &H50
End Sub

Sub AF_TextBoxReadOnly (TXT As TextBox)
'By Neo
'Sets a text box to read only, which means you can't edit it.

DoEvents
Q% = SendMessage(TXT.hWnd, EM_SETREADONLY, 1, 0)
End Sub

Sub AF_WriteListINI (ByVal INIPath As String, Lst As Control)
'This is a simple Sub that will write a list or combobox
'to an INI file you specify.  Use AF_ReadListINI to read
'an INI file written with this sub.

If AC_FileExist(INIPath$) Then Kill INIPath$
FNum = FreeFile
Open INIPath$ For Output As #FNum
    For i = 0 To Lst.ListCount
	If Lst.List(i) <> "" Then Print #FNum, Lst.List(i)
	Next i
    Close #FNum
End Sub

Sub AOT_Turkey ()
'I'm sorry... I just needed to include this. =]
Send "The AOTurkey Logo"
Send "Please excuse me while I scroll my turkey"
Timeout (2)
Send "                                  ,+*^^*+___+++_"
Send "                               ,*^^^^              )"
Send "                            _+*                     ^**+_"
Timeout (2)
Send "                          +^       _ _++*+_+++_,         )"
Send "              _+^^*+_    (     ,+*^ ^          \+_        )"
Send "             {       )  (    ,(    ,_+--+--,      ^)      ^\"
Timeout (2)
Send "            { (@)    } f   ,(  ,+-^ __*_*_  ^^\_   ^\       )"
Send "           {:;-/    (_+*-+^^^^^+*+*<_ _++_)_    )    )     | /"
Send "          ( /  (    (        ,___    ^*+_+* )   <    <      \"
Timeout (2)
Send "           U _/     )    *--<  ) ^\-----++__)   )    )       )"
Send "            (      )  _(^)^^))  )  )\^^^^^))^*+/    /       /"
Send "          (      /  (_))_^)) )  )  ))^^^^^))^^^)__/     +^^"
Timeout (2)
Send "         (     ,/    (^))^))  )  ) ))^^^^^^^))^^)       _)"
Send "          *+__+*       (_))^)  ) ) |     ))^^^^^^))^^^^^)____*^"
Send "          \             \_)^)_)) ))^^^^^^^^^^))^^^^)"
Timeout (2)
Send "           (_             ^\__^^^^^^^^^^^^))^^^^^^^)"
Send "             ^\___            ^\__^^^^^^))^^^^^^^^)\\"
Send "             |     ^^^^^\uuu/^^\uuu/^^^^\^\^\^\^\^\^\^\"
Timeout (2)
Send "                     ___) >____) >___   ^\_\_\_\_\_\_\)"
Send "                    ^^^//\\_^^//\\_^       ^(\_\_\_\)"
Send "                      ^^^ ^^ ^^^ ^^"
Timeout (2)
Send "-={[ Brought to you by AOTurkey ]}=-"
Send "-={[The Turkey revived by Anubis]}=-"
Timeout (2)
End Sub

Function INI_Read (AppName, KeyName, filename As String) As String
'Example: List1.AddItem INI_Read("Elysium", "RoomList", app.Path & "\ely.ini")
sRet = String(255, Chr(0))
INI_Read$ = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function

Sub INI_Write (sAppname As String, sKeyName As String, sNewString As String, sFileName As String)
'Example: INI_Write("Elysium", "Rooms", ListComp$, app.path + "\ely.ini")
Q% = WritePrivateProfileString(sAppname$, sKeyName$, sNewString$, sFileName$)
End Sub

Function MM_CreateMailList (Lst As ListBox) As String
For i = 0 To Lst.ListCount - 1
    Final$ = Final$ & "," & Lst.List(i)
    Next i
MM_CreateMailList$ = "( " & Final$ & " )"
End Function

Function MM_Email1Name (T As String, Subj As String, Message As String)
AOL% = FindWindow("AOL Frame25", 0&)
Ver% = AC_AOLVersion()
AOL% = FindWindow("AOL Frame25", 0&)
b% = FindChildByClass(AOL%, "AOL Toolbar")
DoEvents
C% = AC_GetAOLWin(b%, "_AOL_Icon", 2)
DoEvents
AC_Click (C%)
DoEvents
Do
DoEvents
heh% = FindChildByTitle(AOL%, "Compose Mail")
DoEvents
Timeout (.001)
Loop Until heh% <> 0
DoEvents
    
    
    SendBttn% = FindChildByClass(heh%, "_AOL_Icon")
    tot = FindChildByTitle(heh%, "To:")
    To1% = GetWindow(tot, GW_HWNDNEXT)
    cct = GetWindow(To1%, GW_HWNDNEXT)
    cc% = GetWindow(cct, GW_HWNDNEXT)
    subjectt% = GetWindow(cc%, GW_HWNDNEXT)
    Subjec% = GetWindow(subjectt%, GW_HWNDNEXT)
    If AC_AOLVersion() = 25 Then Textz2% = AC_GetAOLWin(heh%, "_AOL_Edit", 4)
    If AC_AOLVersion() = 3 Then
	d00d1 = GetWindow(Subjec%, GW_HWNDNEXT)
	d00d2 = GetWindow(d00d1, GW_HWNDNEXT)
	d00d3 = GetWindow(d00d2, GW_HWNDNEXT)
	d00d4 = GetWindow(d00d3, GW_HWNDNEXT)
	d00d5 = GetWindow(d00d4, GW_HWNDNEXT)
	d00d6 = GetWindow(d00d5, GW_HWNDNEXT)
	d00d7 = GetWindow(d00d6, GW_HWNDNEXT)
	d00d8 = GetWindow(d00d7, GW_HWNDNEXT)
	d00d9 = GetWindow(d00d8, GW_HWNDNEXT)
	d00d10 = GetWindow(d00d9, GW_HWNDNEXT)
	d00d11 = GetWindow(d00d10, GW_HWNDNEXT)
	d00d12 = GetWindow(d00d11, GW_HWNDNEXT)
	Textz2% = GetWindow(d00d12, GW_HWNDNEXT)
	Textz2% = GetWindow(Textz2%, GW_HWNDNEXT)
	Textz2% = GetWindow(Textz2%, GW_HWNDNEXT)
    End If
DoEvents
Q% = sendmessagebystring(To1%, WM_SETTEXT, 0, T$)
Q% = sendmessagebystring(Subjec%, WM_SETTEXT, 0, Subj$)
Q% = sendmessagebystring(Textz2%, WM_SETTEXT, 0, CStr(Message$))
AC_Click (SendBttn%)
DoEvents
Mssg% = 0
Er% = 0
Do Until Mssg% <> 0 Or Er% <> 0
    DoEvents
    Er% = FindChildByTitle(AOL%, "Error")
    If AC_AOLVersion() = 25 Then Mssg% = FindWindow("#32770", "America Online")
    If AC_AOLVersion() = 3 Then Mssg% = FindWindow("_AOL_MODAL", 0&)
    Timeout (.001)
    Loop
If Er% <> 0 Then
    Q% = SendMessage(heh%, WM_CLOSE, 0, 0&)
    Q% = SendMessage(Er%, WM_CLOSE, 0, 0&)
    MM_Email1Name = T$
    DoEvents
   Else
    Bttn% = FindChildByTitle(Mssg%, "OK")
    AC_Click (Bttn%)
    DoEvents
    MM_Email1Name = ""
    End If
End Function

Sub MM_FillFwd (T As String, Message As String, KeepFwd As Integer)
'If KeepFwd = False then MM_FillFwd will remove the first 5
'characters of the "Subject:" line ("Fwd: ").

AOL% = AC_AOL()
Ver% = AC_AOLVersion()
FwdWin% = FindChildByTitle(AOL%, "Fwd:")
If FwdWin% = 0 Then Debug.Print "Fwd Window NOT found!": Exit Sub
ToEd% = FindChildByClass(FwdWin%, "_AOL_EDIT")
SubjEd% = AC_GetAOLWin(FwdWin%, "_AOL_EDIT", 3)
If Ver% = 25 Then MainEd% = AC_GetAOLWin(FwdWin%, "_AOL_EDIT", 4) Else MainEd% = FindChildByClass(FwdWin%, "RICHCNTL")
SendBttn% = FindChildByClass(FwdWin%, "_AOL_ICON")
DoEvents
AC_SetText ToEd%, T$
AC_SetText MainEd%, CStr(Message$)
If KeepFwd% = False Then
    DoEvents
    SubjText$ = AC_GetWinText(SubjEd%)
    SubjText$ = Mid$(SubjText$, 5)
    AC_SetText SubjEd%, SubjText$
    End If
AC_Click SendBttn%
End Sub

Function MM_FindFwd () As Integer
AOL% = FindWindow("AOL Frame25", 0&)
If AC_AOLVersion() = 25 Then Fwd% = GetParent(FindChildByClass(AOL%, "_AOL_EDIT"))
If AC_AOLVersion() = 3 Then Fwd% = GetParent(FindChildByClass(AOL%, "RICHCNTL"))
SN% = GetParent(FindChildByTitle(AOL%, "Send Now"))
SL% = GetParent(FindChildByTitle(AOL%, "Send Later"))
Ab% = GetParent(FindChildByTitle(AOL%, "Address" & Chr(13) & "Book"))
f% = GetParent(FindChildByTitle(SN%, "Forward")) 'this should be zero if everything is right.
Timeout (.001)
If Fwd% = SN% And SN% = SL% And SL% = Ab% And SL% <> f% Then MM_FindFwd% = Fwd% Else MM_FindFwd% = 0
End Function

Sub MM_FixErrMulti (Lst As ListBox)
'Created by Ruler.
ErrorBox% = FindChildByTitle(AC_AOL(), "Error")
AOL% = AC_AOL()
View% = FindChildByClass(ErrorBox%, "_AOL_VIEW")
InView$ = AC_GetWinText(View%)
AC_Click FindChildByTitle(ErrorBox%, "OK")
InView$ = AF_Script(InView$, "The following problems occurred while processing your request:", "")
For i = 0 To Lst.ListCount - 1
    If InStr(InView$, Lst.List(i)) <> 0 Then Lst.RemoveItem (i)
    Next i
End Sub

Function MM_FixError (MList As String) As String
AOL% = FindWindow("AOL Frame25", 0&)
Do Until ErrorS <> 0
    DoEvents
    ErrorS = FindChildByTitle(AOL%, "Error")
    ErrorM = FindChildByClass(ErrorS, "_AOL_View")
    Timeout (.001)
    Loop
Do Until Len(ErrorText$) <> 0
    DoEvents
    z = SendMessageByNum(ErrorM, WM_GETTEXTLENGTH, 0, 0&)
    ErrorText2$ = String$(z + 1, 0)
    G% = sendmessagebystring(ErrorM, WM_GETTEXT, 0, ErrorText2$)
    ErrorText$ = Left(ErrorText2$, G%)
    Timeout (.001)
    Loop
ErrorOK = FindChildByTitle(ErrorS, "OK")
AC_Click Int(ErrorOK)
DoEvents
Do Until DashPos <> 0
    DoEvents
    DashPos = InStr(ErrorText$, "-")
    Timeout (.001)
    Loop
ErrorText$ = Left$(ErrorText$, DashPos - 2)
ErrorText$ = AF_Script(ErrorText$, "The following problems occurred while processing your request:" & Chr(13) & Chr(10) & Chr(13) & Chr(10), "")
MM_FixError = AF_Script(LCase(MList$), LCase(ErrorText$), "")
End Function

Function MM_FixErrorGetSN (MList As String) As String
AOL% = FindWindow("AOL Frame25", 0&)
Do Until ErrorS <> 0
    DoEvents
    ErrorS = FindChildByTitle(AOL%, "Error")
    ErrorM = FindChildByClass(ErrorS, "_AOL_View")
    Timeout (.001)
    Loop
Do Until Len(ErrorText$) <> 0
    DoEvents
    z = SendMessageByNum(ErrorM, WM_GETTEXTLENGTH, 0, 0&)
    ErrorText2$ = String$(z + 1, 0)
    G% = sendmessagebystring(ErrorM, WM_GETTEXT, 0, ErrorText2$)
    ErrorText$ = Left(ErrorText2$, G%)
    Timeout (.001)
    Loop
ErrorOK = FindChildByTitle(ErrorS, "OK")
ButtonDown = SendMessageByNum(ErrorOK, WM_LBUTTONDOWN, 0&, 0&)
ButtonUp = SendMessageByNum(ErrorOK, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do Until DashPos <> 0
    DoEvents
    DashPos = InStr(ErrorText$, "-")
    Timeout (.001)
    Loop
ErrorText$ = Left$(ErrorText$, DashPos - 2)
ErrorText$ = AF_Script(ErrorText$, "The following problems occurred while processing your request:" & Chr(13) & Chr(10) & Chr(13) & Chr(10), "")
MM_FixErrorGetSN = ErrorText$
End Function

Sub MM_FullEmail (SN As String, Subj As String, Mes As String, Attach As String, SendMe As Integer)
'This sub'll allow you to attach a file and select if the
'mail should be sent or not...
'Set SendMe% to 1 if you want to send immediately, 2 if
'you wish to send later and 3 if you wish to not send at
'all.
If SN$ = "" Then MsgBox "You have not specified anyone to send mail to.", 0, "Error": Exit Sub
If Subj$ = "" Then Subj$ = Chr(160)
If Mes$ = "" Then Mes$ = Chr(160)
If SendMe% < 1 Or SendMe% > 3 Then SendMe% = 3
AOL% = FindWindow("AOL Frame25", 0&)
Ver% = AC_AOLVersion()
Tool% = FindChildByClass(AOL%, "AOL Toolbar")
NMBttn% = AC_GetAOLWin(Tool%, "_AOL_ICON", 2)
AC_Click NMBttn%
AC_WaitForChange
CompMail% = FindChildByTitle(AOL%, "Compose Mail")
ToEd% = FindChildByClass(CompMail%, "_AOL_EDIT")
SubjEd% = AC_GetAOLWin(CompMail%, "_AOL_EDIT", 3)
If Ver% = 25 Then MainEd% = AC_GetAOLWin(CompMail%, "_AOL_EDIT", 4) Else MainEd% = FindChildByClass(CompMail%, "RICHCNTL")
SendBttn% = FindChildByClass(CompMail%, "_AOL_ICON")
SendLaterBttn% = AC_GetAOLWin(CompMail%, "_AOL_ICON", 2)
AttachBttn% = AC_GetAOLWin(CompMail%, "_AOL_ICON", 3)
AC_SetText ToEd%, CStr(SN$)
AC_SetText SubjEd%, Subj$
AC_SetText MainEd%, CStr(Mes$)
If Len(Attach$) <> 0 Then
    AC_Click AttachBttn%
    AC_WaitForChange
    DlgWin% = FindWindow("#32770", "Attach File")
    FilEd% = FindChildByClass(DlgWin%, "Edit")
    OKBttn% = FindChildByTitle(DlgWin%, "OK")
    AC_SetText FilEd%, Attach$
    AC_Click OKBttn%
    For e = 0 To 7
	DoEvents
	ErrWin% = FindWindow("#32770", "Attach File")
	If ErrWin% <> DlgWin% Then
	    OKBttn% = FindChildByTitle(ErrWin%, "OK")
	    AC_Click ok%
	    AC_Close DlgWin%
	    ErrRetCode = 1
	    SendMe% = 2
	    Exit For
	    End If
	Timeout (.001)
	Next e
    End If
    If SendMe% = 1 Then
Resend:
	AC_Click SendBttn%
	ErrWin% = 0
	Do Until ErrWin% <> 0 Or Success% <> 0
	    DoEvents
	    ErrWin% = FindChildByTitle(AOL%, "Error")
	    If Ver% = 25 Then Success% = FindWindow("#32770", "America Online") Else Success% = FindWindow("_AOL_MODAL", 0&)
	    Timeout (.001)
	    Loop
	If ErrWin% <> 0 Then
	    SN$ = CStr(MM_FixError(SN$))
	    If Len(SN$) = 0 Or SN$ = "()" Then
		ErrRetCode = 2
		AC_SetText ToEd%, SN$
		GoTo Final
		End If
	    GoTo Resend
	    End If
	If Len(Attach$) <> 0 Then GoTo Final
	If Success% <> 0 Then
	    OKBttn% = FindChildByTitle(Success%, "OK")
	    AC_Click OKBttn%
	    GoTo Final
	    End If
	End If
If SendMe% = 2 Then
    AC_Click SendLaterBttn%
    Success% = 0
    Do Until Success% <> 0
	DoEvents
	If Ver% = 25 Then Success% = FindWindow("#32770", "America Online") Else Success% = FindWindow("_AOL_MODAL", 0&)
	Timeout (.001)
	Loop
    OKBttn% = FindChildByTitle(Success%, "OK")
    AC_Click OKBttn%
    GoTo Final
    End If
If SendMe% = 3 Then GoTo Final
Final:
If ErrRetCode = 1 Then MsgBox CStr(LCase(Attach$) & AF_Enter() & AF_Enter() & "does not exist so it could not be attached.  The mail message was not sent."), 0, "Error"
If ErrRetCode = 2 Then MsgBox "None of the screen names specified were known AOL members so the mail could not be sent.", 0, "Error"
End Sub

Function MM_KillDupes (UseDelete As Integer) As Integer
'This Sub will remove all duplicate mails in the New Mail
'box.  UseDelete is Boolean (True or False) and determines
'if you want the mail deleted or just ignored.  True will
'cause it all to be deleted while False will just use the
'ignore button.  Returns the number of mails deleted.

AOL% = AC_AOL()
If AOL% = 0 Then Exit Function
MailCount% = MM_OpenMailGetCount()
MailWin% = FindChildByTitle(AOL%, "New Mail")
Q% = ShowWindow(MailWin%, SW_HIDE)
Tree% = FindChildByClass(MailWin%, "_AOL_TREE")
DelBttn% = FindChildByTitle(NewMail%, "Delete")
IgBttn% = FindChildByTitle(NewMail%, "Ignore")
ReDim srcList(0 To MailCount% - 1)
ReDim delList(0 To MailCount% - 1)
For i = 0 To MailCount% - 1
    MailStr$ = String$(255, " ")
    Q% = sendmessagebystring(Tree%, LB_GETTEXT, i, MailStr$)
    NoDate$ = Mid$(MailStr$, InStr(MailStr$, "/") + 4)
    NoSN$ = Mid$(NoDate$, InStr(NoDate$, Chr(9)) + 1)
    srcList(i) = TrimNull(NoSN$)
    DoEvents
    Next i
Q% = ShowWindow(MailWin%, SW_SHOW)
For i = 0 To MailCount% - 1
    delList(i) = srcList(i)
    Next i
For i = 0 To MailCount% - 1
    curMail$ = srcList(i)
    For n = 0 To MailCount% - 1
	If n <> i And curMail$ Like delList(i) Then delList(i) = "delme"
	Next n
    Next i
delCount = 0
For i = 0 To MailCount% - 1
    DoEvents
    If delList(i) Like "delme" Then
	Q% = SendMessageByNum(Tree%, LB_SETCURSEL, i, 0&)
	If UseDelete% = True Then AC_Click DelBttn% Else AC_Click IgBttn%
	delCount = delCount + 1
	End If
    Next i
MM_KillDupes% = Int(delCount)
End Function

Function MM_KillError (ChipsAhoy As String) As String
AOL% = FindWindow("AOL Frame25", 0&)
e% = FindChildByTitle(AOL%, "Error")
EView% = FindChildByClass(e%, "_AOL_VIEW")
EOk% = FindChildByClass(e%, "_AOL_BUTTON")
y$ = AC_GetWinText(EView%)
y$ = AF_Script(y$, "The following problems occurred while processing your request:" & Chr(13) & Chr(10) & Chr(13) & Chr(10), "")
DashPos = InStr(y$, "-")
If Len(y$) <> 0 Then z$ = TrimNull(Left(y$, DashPos - 1))
Debug.Print "SN=" & z$
K% = FindChildByClass(e%, "_AOL_BUTTON")
AC_Click (K%)
DoEvents
FinalList$ = AF_Script(ChipsAhoy$, z$, "")
MM_KillError$ = FinalList$
End Function

Function MM_LocateMail (FindFwd As Integer) As Integer
'This locates the Mail window... if FindFwd = true then
'it returns the value of the Forward button on the mail
'window, otherwise it returns the value of the mail win-
'dow itself.

AOL% = AC_AOL()
Target% = GetFocus()
mail% = 0
Do Until mail% <> 0
    DoEvents
    Q% = SetFocusAPI(Target%)
    Stat1% = GetParent(FindChildByTitle(AOL%, "Reply"))
    Stat2% = GetParent(FindChildByTitle(AOL%, "Forward"))
    Stat3% = GetParent(FindChildByTitle(AOL%, "Reply to All"))
    If Stat1% = Stat2% And Stat2% = Stat3% And Stat3% <> 0 Then
	mail% = Stat1%: Exit Do
	End If
    Target% = GetWindow(AOL%, GW_HWNDNEXT)
    Loop
If FindFwd% = False Then MM_LocateMail% = mail%: Exit Function
FwdBttn% = AC_GetAOLWin(mail%, "_AOL_ICON", 2)
MM_LocateMail% = FwdBttn%
End Function

Sub MM_MailToClipboard ()
Final$ = MM_MailToString()
ClipBoard.SetText CStr(Final$)
DoEvents
End Sub

Sub MM_MailToList (Lst As ListBox)
AOL% = FindWindow("AOL Frame25", 0&)
If AOL% = 0 Then Exit Sub
MailCount% = MM_OpenMailGetCount()
MailWin% = FindChildByTitle(AOL%, "New Mail")
Q% = ShowWindow(MailWin%, SW_HIDE)
Tree% = FindChildByClass(MailWin%, "_AOL_TREE")
For i = 0 To MailCount% - 1
    MailStr$ = String$(255, " ")
    Q% = sendmessagebystring(Tree%, LB_GETTEXT, i, MailStr$)
    NoDate$ = Mid$(MailStr$, InStr(MailStr$, "/") + 4)
    NoSN$ = Mid$(NoDate$, InStr(NoDate$, Chr(9)) + 1)
    Lst.AddItem TrimNull(NoSN$)
    DoEvents
    Next i
Q% = ShowWindow(MailWin%, SW_SHOW)
DoEvents
End Sub

Function MM_MailToString () As String
AOL% = FindWindow("AOL Frame25", 0&)
If AOL% = 0 Then Exit Function
MailCount% = MM_OpenMailGetCount()
MailWin% = FindChildByTitle(AOL%, "New Mail")
Q% = ShowWindow(MailWin%, SW_HIDE)
Tree% = FindChildByClass(MailWin%, "_AOL_TREE")
ReDim MailBuf(0 To MailCount%)
For i = 0 To MailCount% - 1
    MailStr$ = String$(255, " ")
    Q% = sendmessagebystring(Tree%, LB_GETTEXT, i, MailStr$)
    NoDate$ = Mid$(MailStr$, InStr(MailStr$, "/") + 4)
    NoSN$ = Mid$(NoDate$, InStr(NoDate$, Chr(9)) + 1)
    MailBuf(i) = TrimNull(NoSN$)
    DoEvents
    Next i
For i = 0 To MailCount%
    Final$ = Final$ & MailBuf(i) & AF_Enter()
    Next i
Q% = ShowWindow(MailWin%, SW_SHOW)
MM_MailToString$ = CStr(Final$)
DoEvents
End Function

Sub MM_MMPref ()
AOL% = FindWindow("AOL Frame25", 0&)
Ver% = AC_AOLVersion()
If Ver% = 25 Then Pr$ = "Set Preferences"
If Ver% = 3 Then Pr$ = "Preferences"
Call AC_RunMenuByString(Pr$, "Mem&bers")
Do Until Pref% <> 0
    DoEvents
    Pref% = FindChildByTitle(AOL%, "Preferences")
    Timeout .001
    Loop
MailPrefBttn% = AC_GetAOLWin(Pref%, "_AOL_Icon", 6)
AC_Click (MailPrefBttn%)
Do Until MailPref% <> 0
    DoEvents
    MailPref% = FindWindow("_AOL_Modal", "Mail Preferences")
    Timeout .001
    Loop
OKBttn% = FindChildByTitle(MailPref%, "OK")
C1% = FindChildByTitle(MailPref%, "Confirm mail after it has been sent")
C2% = FindChildByTitle(MailPref%, "Close mail after it has been sent")
DoEvents
Q% = SendMessageByNum(C1%, BM_SETCHECK, True, 0&)
Q% = SendMessageByNum(C2%, BM_SETCHECK, True, 0&)
AC_Click OKBttn%
AC_Close Pref%
End Sub

Function MM_OpenMailGetCount () As Integer
AOL% = FindWindow("AOL Frame25", 0&)
NewMail% = FindChildByTitle(AOL%, "New Mail")
If NewMail% = 0 Then
    Do Until NewMail% <> 0 Or Mssg% <> 0
	DoEvents
	Tool% = FindChildByClass(AOL%, "AOL Toolbar")
	NewMailBttn% = FindChildByClass(Tool%, "_AOL_ICON")
	AC_Click (NewMailBttn%)
	Timeout (1)
	NewMail% = FindChildByTitle(AOL%, "New Mail")
	Mssg% = AC_AOLMsgBox()
	DoEvents
	Timeout (.001)
	Loop
	DoEvents
	End If
If Mssg% <> 0 Then
    OKBttn% = FindChildByTitle(Mssg%, "OK")
    AC_Click OKBttn%
    AC_OpenMailGetCount% = 0
    Exit Function
    End If
Do Until NewMail% <> 0
    DoEvents
    NewMail% = FindChildByTitle(AOL%, "New Mail")
    Timeout (.001)
    Loop
Hand% = FindChildByClass(NewMail%, "_AOL_Tree")
Do Until MailNum <> 0 And MailNum = MailNum2 And MailNum2 = MailNum3
    DoEvents
    MailNum = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0&)
    Call Timeout(.7)
    MailNum2 = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0&)
    Call Timeout(1.5)
    MailNum3 = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0&)
    Timeout (.346)
    Loop
Hand% = FindChildByClass(NewMail%, "_AOL_TREE")
Buffer = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0)
MM_OpenMailGetCount% = Buffer
End Function

Function MM_OpenMailGetCount2 (MailType As String) As Integer
'Valid MailTypes are "Old Mail", "Outgoing Mail",
'"New Mail"
AOL% = FindWindow("AOL Frame25", 0&)
NewMail% = FindChildByTitle(AOL%, "MailType$")
If NewMail% = 0 Then
    Do Until NewMail% <> 0 Or Mssg% <> 0
	DoEvents
	Tool% = FindChildByClass(AOL%, "AOL Toolbar")
	NewMailBttn% = FindChildByClass(Tool%, "_AOL_ICON")
	AC_Click (NewMailBttn%)
	Timeout (1)
	NewMail% = FindChildByTitle(AOL%, "New Mail")
	Mssg% = AC_AOLMsgBox()
	DoEvents
	Timeout (.001)
	Loop
	DoEvents
	End If
If Mssg% <> 0 Then
    OKBttn% = FindChildByTitle(Mssg%, "OK")
    AC_Click OKBttn%
    AC_OpenMailGetCount% = 0
    Exit Function
    End If
Do Until NewMail% <> 0
    DoEvents
    NewMail% = FindChildByTitle(AOL%, "New Mail")
    Timeout (.001)
    Loop
Hand% = FindChildByClass(NewMail%, "_AOL_Tree")
Do Until MailNum <> 0 And MailNum = MailNum2 And MailNum2 = MailNum3
    DoEvents
    MailNum = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0&)
    Call Timeout(.7)
    MailNum2 = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0&)
    Call Timeout(1.5)
    MailNum3 = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0&)
    Timeout (.346)
    Loop
Hand% = FindChildByClass(NewMail%, "_AOL_TREE")
Buffer = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0)
MM_OpenMailGetCount2% = Buffer
End Function

Sub PlaySound (XSound As String)
Q% = SndPlaySound(XSound, 1)
DoEvents
End Sub

Sub PlayVideo (ByVal VidPath As String, ByVal wSize As Integer)
'wSize can be set to 0 for a window or -1 for fullscreen.
If wSize = 0 Then sSize$ = "window " Else sSize$ = "fullscreen "
If AC_FileExist(VidPath$) = False Then Exit Sub: Debug.Print "Video not found."
CmdStr$ = "Play %PATH% & sSize$"
Q& = mciSendString(AF_Script(CmdStr$, "%PATH%", VidPath$), 0&, 0, 0&)
End Sub

Function AC_AOLMsgBox () As Integer
'This function locates AOL's Message Boxes like the ones
'that say if IMs are off or a room is full.

AC_AOLMsgBox% = FindWindow("#32770", "America Online")
End Function

Function AC_AOLVersion ()
'This figures out which AOL version they are using:
'AOL 3.0 or 2.5
AOL% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(AOL%, "Welcome, " + AC_GetSN())
aol3% = FindChildByClass(Wel%, "RICHCNTL")
If aol3% = 0 Then AC_AOLVersion = 25: Exit Function
If aol3% <> 0 Then
    If AC_GetWinText(AOL%) <> "America Online" Then AC_AOLVersion = 3 Else AC_AOLVersion = 4
    End If
End Function

Function AC_Available (ByVal SN As String) As Integer
Call AC_PrepIM(SN$)
O% = 0
IM% = 0
C% = 0
Ver% = AC_AOLVersion()
AOL% = FindWindow("AOL Frame25", 0&)
IM% = FindChildByTitle(AOL%, "Send Instant Message")
If Ver% = 25 Then Avail% = FindChildByTitle(IM%, "Available?")
If Ver% = 3 Then
    Waste% = FindChildByClass(IM%, "RICHCNTL")
    Waste2% = GetWindow(Waste%, GW_HWNDNEXT)
    Avail% = GetWindow(Waste2%, GW_HWNDNEXT)
    End If
AC_Click Avail%
Do Until O% <> 0
    DoEvents
    O% = FindWindow("#32770", "America Online")
    Timeout (.001)
    Loop
MsgBoxTxt$ = AC_GetMsgText(O%)
If InStr(MsgBoxTxt$, "is online and able to receive Instant Messages.") <> 0 Then AC_Available = True Else AC_Available = False
K% = FindChildByTitle(O%, "OK")
AC_Click K%
AC_Close IM%
End Function

Sub AC_Click (Bttn As Integer)
DoEvents
Q% = SendMessageByNum(Bttn%, WM_LBUTTONDOWN, 0, 0&)
Q% = SendMessageByNum(Bttn%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AC_ClickList (Lis As Integer)
'This will double click a list box's item; the item
'that's currently selected (Use LB_SETCURSEL)

Q% = SendMessageByNum(Lis, WM_LBUTTONDBLCLK, 0, 0&)
DoEvents
End Sub

Sub AC_Close (Target As Integer)
Q% = SendMessage(Target%, WM_CLOSE, 0, 0&)
DoEvents
End Sub

Sub AC_CloseBuddy ()
AOL% = FindWindow("AOL Frame25", 0&)
Bud% = FindChildByTitle(AOL%, "Buddy Lists")
Q% = SendMessage(Bud%, WM_CLOSE, 0, 0&)
DoEvents
End Sub

Sub AC_CreateMenu (mnuTitle As String, mnuPopUps As String)
'This sub will append menus to AOL.  You need to assign
'mnuPopUps$ a series of menus and indexes;
'<menuName:Index;menuName:Index>
'Here is an Example:

'MenusToAdd$ = "New Item:1;&File:2;Killer:3"
'Call AC_CreateMenu("&Test", MenusToAdd$)

AOL% = FindWindow("AOL Frame25", 0&)
If AOL% = 0 Then Exit Sub
AOLMenu% = getmenu(AOL%)
hMenuPopup% = CreatePopupMenu()
SplitterCounter% = 0
For i = 1 To Len(mnuPopUps$)
    ExamineChar$ = Mid$(mnuPopUps$, i, 1)
    If ExamineChar$ = ":" Then SplitterCounter% = SplitterCounter% + 1
    Next i
Egg$ = mnuPopUps$
For i = 0 To SplitterCounter%
    If Egg$ = "" Then Exit For
    If InStr(Egg$, ":") = 0 Then Exit For
    mnuName$ = Left$(Egg$, InStr(Egg$, ":") - 1)
    Egg$ = AF_Script(Egg$, mnuName$ & ":", "")
    If InStr(Egg$, ";") <> 0 Then
	mnuIndex% = CInt(Left(Egg$, InStr(Egg$, ";") - 1))
       Else
	mnuIndex% = CInt(Egg$)
	End If
    Egg$ = AF_Script(Egg$, AF_Script(Str(mnuIndex%), " ", "") & ";", "")
    Q% = AppendMenu(hMenuPopup%, MF_ENABLED Or MF_STRING, mnuIndex%, mnuName$)
    Next i
Q% = AppendMenu(AOLMenu%, MF_STRING Or MF_POPUP, hMenuPopup%, mnuTitle$)
DrawMenuBar AOL%
End Sub

Function AC_Disc (ByVal in As String) As String
On Error Resume Next
AC_Disc$ = Left$(in$, InStr(Marbro$, Chr$(0)) - 1)
End Function

Sub AC_Email (T As String, Subj As String, Message As String)
AOL% = FindWindow("AOL Frame25", 0&)
Ver% = AC_AOLVersion()
Tool% = FindChildByClass(AOL%, "AOL Toolbar")
C% = AC_GetAOLWin(Tool%, "_AOL_Icon", 2)
AC_Click (C%)
DoEvents
Do
DoEvents
heh% = FindChildByTitle(AOL%, "Compose Mail")
DoEvents
Timeout (.001)
Loop Until heh% <> 0
DoEvents
    
    
    SendBttn% = FindChildByClass(heh%, "_AOL_Icon")
    tot = FindChildByTitle(heh%, "To:")
    To1% = GetWindow(tot, GW_HWNDNEXT)
    cct = GetWindow(To1%, GW_HWNDNEXT)
    cc% = GetWindow(cct, GW_HWNDNEXT)
    subjectt% = GetWindow(cc%, GW_HWNDNEXT)
    Subjec% = GetWindow(subjectt%, GW_HWNDNEXT)
    If AC_AOLVersion() = 25 Then Textz2% = AC_GetAOLWin(heh%, "_AOL_Edit", 4)
    If AC_AOLVersion() = 3 Then
	d00d1 = GetWindow(Subjec%, GW_HWNDNEXT)
	d00d2 = GetWindow(d00d1, GW_HWNDNEXT)
	d00d3 = GetWindow(d00d2, GW_HWNDNEXT)
	d00d4 = GetWindow(d00d3, GW_HWNDNEXT)
	d00d5 = GetWindow(d00d4, GW_HWNDNEXT)
	d00d6 = GetWindow(d00d5, GW_HWNDNEXT)
	d00d7 = GetWindow(d00d6, GW_HWNDNEXT)
	d00d8 = GetWindow(d00d7, GW_HWNDNEXT)
	d00d9 = GetWindow(d00d8, GW_HWNDNEXT)
	d00d10 = GetWindow(d00d9, GW_HWNDNEXT)
	d00d11 = GetWindow(d00d10, GW_HWNDNEXT)
	d00d12 = GetWindow(d00d11, GW_HWNDNEXT)
	Textz2% = GetWindow(d00d12, GW_HWNDNEXT)
	Textz2% = GetWindow(Textz2%, GW_HWNDNEXT)
	Textz2% = GetWindow(Textz2%, GW_HWNDNEXT)
    End If
DoEvents
Q% = sendmessagebystring(To1%, WM_SETTEXT, 0, T$)
Q% = sendmessagebystring(Subjec%, WM_SETTEXT, 0, Subj$)
Q% = sendmessagebystring(Textz2%, WM_SETTEXT, 0, CStr(Message$))
AC_Click (SendBttn%)
DoEvents
Mssg% = 0
Do Until Mssg% <> 0 Or Er% <> 0
    DoEvents
    If AC_AOLVersion() = 25 Then Mssg% = FindWindow("#32770", "America Online")
    If AC_AOLVersion() = 3 Then Mssg% = FindWindow("_AOL_MODAL", 0&)
    Timeout (.001)
    Loop
Bttn% = FindChildByTitle(Mssg%, "OK")
AC_Click (Bttn%)
DoEvents
End Sub

Sub AC_EmailNoSend (T As String, Subj As String, Message As String)
AOL% = FindWindow("AOL Frame25", 0&)
Ver% = AC_AOLVersion()
b% = FindChildByClass(AOL%, "AOL Toolbar")
DoEvents
C% = AC_GetAOLWin(b%, "_AOL_Icon", 2)
DoEvents
AC_Click (C%)
DoEvents
Do
DoEvents
heh% = FindChildByTitle(AOL%, "Compose Mail")
DoEvents
Timeout (.001)
Loop Until heh% <> 0
DoEvents
    
    
    SendBttn% = FindChildByClass(heh%, "_AOL_Icon")
    tot = FindChildByTitle(heh%, "To:")
    To1% = GetWindow(tot, GW_HWNDNEXT)
    cct = GetWindow(To1%, GW_HWNDNEXT)
    cc% = GetWindow(cct, GW_HWNDNEXT)
    subjectt% = GetWindow(cc%, GW_HWNDNEXT)
    Subjec% = GetWindow(subjectt%, GW_HWNDNEXT)
    If AC_AOLVersion() = 25 Then Textz2% = AC_GetAOLWin(heh%, "_AOL_Edit", 4)
    If AC_AOLVersion() = 3 Then
	d00d1 = GetWindow(Subjec%, GW_HWNDNEXT)
	d00d2 = GetWindow(d00d1, GW_HWNDNEXT)
	d00d3 = GetWindow(d00d2, GW_HWNDNEXT)
	d00d4 = GetWindow(d00d3, GW_HWNDNEXT)
	d00d5 = GetWindow(d00d4, GW_HWNDNEXT)
	d00d6 = GetWindow(d00d5, GW_HWNDNEXT)
	d00d7 = GetWindow(d00d6, GW_HWNDNEXT)
	d00d8 = GetWindow(d00d7, GW_HWNDNEXT)
	d00d9 = GetWindow(d00d8, GW_HWNDNEXT)
	d00d10 = GetWindow(d00d9, GW_HWNDNEXT)
	d00d11 = GetWindow(d00d10, GW_HWNDNEXT)
	d00d12 = GetWindow(d00d11, GW_HWNDNEXT)
	Textz2% = GetWindow(d00d12, GW_HWNDNEXT)
	Textz2% = GetWindow(Textz2%, GW_HWNDNEXT)
	Textz2% = GetWindow(Textz2%, GW_HWNDNEXT)
    End If
DoEvents
Q% = sendmessagebystring(To1%, WM_SETTEXT, 0, T$)
Q% = sendmessagebystring(Subjec%, WM_SETTEXT, 0, Subj$)
Q% = sendmessagebystring(Textz2%, WM_SETTEXT, 0, CStr(Message$))
End Sub

Sub AC_ExitWindows (ByVal nExitOption As Integer, MBox As Integer)
'This sub will Exit Windows and continue the operation
'depending on your choice.  MBox = False = no message box
'whereas true will ask the user if they wish to exit win-
'dows.
If MBox = True Then
MB = MsgBox("Do you want to exit Windows?", 36, "Windows")
    If MB = 7 Then Exit Sub 'User chose NO
    End If
    Select Case nExitOption
	Case 1
	n = ExitWindows(67, 0) 'reboot the computer
	Case 2
	n = ExitWindows(66, 0) 'restart Windows
	Case 3
	n = ExitWindows(0, 0) 'exit Windows
	End Select
End Sub

Function AC_ExtractPW (AOLDir As String, ByVal SN As String) As String
ScreenName$ = SN$ & Space(10 - Len(SN$))
BytesRead = 1
Do
    PW$ = ""
    DoEvents
    On Error Resume Next
    Open AOLDir$ & "\idb\main.idx" For Binary As #1
	If Err Then AC_ExtractPW$ = "": Exit Function
	PW$ = String(32000, 0)
	Get #1, BytesRead, PW$
	Close #1
    Open AOLDir$ & "\idb\main.idx" For Binary As #2
	WherePW = InStr(1, PW$, NumSpaces + Chr(0), 1)
	If WherePW Then
40 :
	    DoEvents
	    Mid(PW$, WherePW) = "Pass Word "
	    midsn = Mid(PW$, WherePW + Len(NumSpaces) + 1, 8)
	    midsn = TrimNull(midsn)
	    midpw = Mid(PW$, WherePW + Len(NumSpaces) + 1 + Len(midsn), 1)
	    If midpw <> Chr(0) Then GoTo 45
	    If Len(midsn) < 4 Then GoTo 45
	    If Len(midsn) = "" Then GoTo 45
	    AC_ExtractPW$ = midsn
45 :
	    WherePW = InStr(1, PW$, NumSpaces + Chr(0), 1)
	    If WherePW Then DoEvents: GoTo 40
	End If
	BytesRead = BytesRead + 32000
	FileLength = LOF(2)
	Close #2
    If BytesRead > FileLength Then GoTo 30
Loop
30 :
End Function

Function AC_FileExist (ByVal sFileName As String) As Integer
'Example: If Not AC_FileExist(app.Path & "\test.ini") then...
Dim i As Integer
On Error Resume Next
i = Len(Dir$(sFileName))
    If Err Or i = 0 Then
	AC_FileExist = False
	Else
	AC_FileExist = True
	End If
Resume Next
End Function

Function AC_GetAOLDir () As String
Dim sModuleFileName As String * 100
AOL% = FindWindow("AOL Frame25", 0&)
If AOL% = 0 Then MsgBox "You must have AOL open to use this program.", 0, "Error": Exit Function
hInstance = GetWindowWord(AOL%, GWW_HINSTANCE)
x = GetModuleFileName(hInstance, sModuleFileName, 100)
FullDir$ = Left$(sModuleFileName, x)
AC_GetAOLDir$ = LCase(AF_Script(FullDir$, "\WAOL.EXE", ""))
End Function

Function AC_GetAOLWin (Parent As Integer, ByVal ClassToFind As String, Num As Integer) As Integer
'This will get any item on a window.  You need only count
'the number of items b/w you and what you want.

'Example:
'PplConIcon% = AC_GetAOLWin(Tool%, "_AOL_ICON", 5)

'This function was originally created by DRAGN.

If Parent% = 0 Then Debug.Print "The parent window was NOT found.": Exit Function
Init% = FindChildByClass(Parent%, ClassToFind$)
If Init% = 0 Then Debug.Print "That classname does not exist on " & Str(Parent%): Exit Function
Count% = Count% + 1
If Count% = Num% Then
    AC_GetAOLWin% = Init%
    Exit Function
    End If
Waste% = Init%
Do Until Found% <> 0
    DoEvents
    Waste% = GetWindow(Waste%, GW_HWNDNEXT)
    Buf$ = String$(255, " ")
    Q% = GetClassName(Waste%, Buf$, 254)
    DoEvents
    If LCase(TrimNull(Buf$)) Like LCase(ClassToFind$) Then Count% = Count% + 1
    If Count% = Num% Then
	AC_GetAOLWin% = Waste%
	Found% = Waste%
	End If
    Loop
End Function

Sub AC_GetComboNames (Lst As Control)
Lst.Clear
For index% = 0 To 25
    SN$ = String$(20, "")
    Q% = AOLGetcombo(index%, SN$)
    If Len(Trim$(names$)) <= 2 Then Exit For
    SN$ = TrimNull(SN$)
    Lst.AddItem (SN$)
    Next index%
End Sub

Function AC_GetHostName (AOLDir As String) As String
If InStr(AOLDir$, "3") <> 0 Then Version = 3 Else Version = 25
Host$ = String$(40, " ")
If Version = 3 Then
    chat$ = "aolchat.aol"
    PNum = 4761
    V% = 9
   Else
    chat$ = "chat.aol"
    PNum = 6887
    V% = 6
    End If
Open "C:\" + AOLDir$ + "\tool\" & chat$ For Binary As #1
Seek #1, PNum
Get #1, PNum, Host$
Close #1
Host$ = Left$(Host$, InStr(Host$, "%s") - V%)
AC_GetHostName$ = Trim(Host$)
End Function

Function AC_GetIMText (IM As Integer) As String
Ver% = AC_AOLVersion()
If Ver% = 25 Then S$ = "_AOL_VIEW" Else S$ = "RICHCNTL"
View% = FindChildByClass(IM%, S$)
AC_GetIMText$ = AC_GetWinText(View%)
End Function

Function AC_GetLastChatLine () As String
'Inspired by Master by VSTDCoord and Progger by Progger
'Place in a timer and set the timer interval to 5 or 7.

If AC_SetToView() = 0 Then Exit Function
RoomText$ = AC_GetWinText(AC_SetToView())
For i = Len(RoomText$) - 1 To 1 Step -1
    If Mid$(RoomText$, i, 1) Like Chr(13) Then Exit For
    Next i
AC_GetLastChatLine$ = Mid$(RoomText$, i + 1)
End Function

Function AC_GetListIndex (Lst As Integer, findstring As String) As Integer
Const LB_FINDSTRINGEXACT = &H400 + 5
AC_GetListIndex = sendmessagebystring(Lst%, LB_FINDSTRINGEXACT, 1, findstring)
End Function

Function AC_GetMsgText (MBox As Integer) As String
Stat1% = FindChildByClass(MBox%, "STATIC")
Stat2% = GetWindow(Stat1%, GW_HWNDNEXT)
Stat$ = AC_GetWinText(Stat2%)
If Stat$ = "" Then Stat$ = AC_GetWinText(Stat1%)
AC_GetMsgText$ = Stat$
End Function

Function AC_GetSN () As String
'This gets the user's SN from the Welcome window.
AOL = FindWindow("AOL Frame25", 0&)
Wel = FindChildByTitle(AOL, "Welcome,")
If Wel = 0 Then AC_GetSN = "Not Online": Exit Function
namelen = SendMessage(Wel, WM_GETTEXTLENGTH, 0, 0)
Buffer$ = String$(namelen, 0)
x = sendmessagebystring(Wel, WM_GETTEXT, namelen, Buffer$)
a = InStr(Buffer$, ",")
SN$ = Mid$(Buffer$, a + 2, (Len(Buffer$) - (a + 1)))
SN$ = TrimNull(SN$)
AC_GetSN$ = SN$
End Function

Function AC_GetSNFromIM (IM As Integer) As String
If IM% = 0 Then AC_GetSNFromIM$ = "No IM."
AC_GetSNFromIM$ = TrimNull(Mid$(AC_GetWinText(IM%), InStr(AC_GetWinText(IM%), ":") + 1))
End Function

Function AC_GetWinText (GetThis As Integer) As String
'This can get a window's caption or get text from just
'about anything that has text including _AOL_EDIT.

'Example:
'WinCaption$ = AC_GetWinText(Pref%)

BufLen% = SendMessageByNum(GetThis%, WM_GETTEXTLENGTH, 0, 0)
Buffer$ = String(BufLen%, 0)
Q% = sendmessagebystring(GetThis%, WM_GETTEXT, BufLen% + 1, Buffer$)
DoEvents
AC_GetWinText$ = TrimNull(Buffer$)
End Function

Sub AC_Hide (Target As Integer)
'This Sub will hide a target window.
Q% = ShowWindow(Target%, SW_HIDE)
DoEvents
End Sub

Sub AC_HideToolbar ()
AOL% = FindWindow("AOL Frame25", 0&)
Tool% = FindChildByClass(AOL%, "AOL Toolbar")
Q% = ShowWindow(Tool%, SW_HIDE)
Q% = ShowWindow(AOL%, SW_HIDE)
Q% = ShowWindow(AOL%, SW_MINIMIZE)
Q% = ShowWindow(AOL%, SW_MAXIMIZE)
Q% = ShowWindow(AOL%, SW_SHOW)
DoEvents
End Sub

Sub AC_HideWelcome ()
AOL% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(AOL%, "Welcome, " & AC_GetSN() & "!")
Q% = ShowWindow(Wel%, SW_HIDE)
DoEvents
End Sub

Sub AC_IM (ByVal SN As String, ByVal Mesg As String)
If AC_Online() = 0 Then Exit Sub
Ver% = AC_AOLVersion()
AOL% = FindWindow("AOL Frame25", 0&)
AOL% = FindChildByClass(AOL%, "MDIClient")
AC_RunMenuByString "Send an Instant Message", "Mem&bers"
AC_WaitForChange
IMWin% = FindChildByTitle(AOL%, "Send Instant Message")
Q% = SetFocusAPI(IMWin%)
Ed1% = FindChildByClass(IMWin%, "_AOL_EDIT")
If Ver% = 25 Then MainEd% = AC_GetAOLWin(IMWin%, "_AOL_EDIT", 2) Else MainEd% = FindChildByClass(IMWin%, "RICHCNTL")
SendBttn% = GetWindow(MainEd%, GW_HWNDNEXT)
If Ed1% <> 0 And SendBttn% <> 0 And MainEd% <> 0 Then IMFound = True Else IMWin% = GetWindow(IMWin%, GW_HWNDNEXT): DoEvents
Q% = SetFocusAPI(IMWin%)
Do Until Len(EdText$) <> 0
    DoEvents
    AC_SetText Ed1%, SN$
    EdText$ = AC_GetWinText(Ed1%)
    Timeout .001
    Loop
Do Until Len(MsgText$) <> 0
    DoEvents
    AC_SetText MainEd%, CStr(Mesg$)
    MsgText$ = AC_GetWinText(MainEd%)
    Timeout .001
    Loop
DoEvents
Q% = SetFocusAPI(IMWin%)
AC_Click SendBttn%
AC_Close IMWin%
DoEvents
For i = 0 To 3
    ErrWin% = FindWindow("#32770", "America Online")
    If ErrWin% <> 0 Then
	OkieBttn% = FindChildByTitle(ErrWin%, "OK")
	AC_Click OkieBttn%
	Exit For
	End If
    Next i
End Sub

Function AC_IMAnswer (ByVal AnsMsg As String) As String
If AC_Online() = 0 Then Exit Function
Ver% = AC_AOLVersion()
AOL% = FindWindow("AOL Frame25", 0&)
IM% = FindChildByTitle(AOL%, ">Instant Message From:")
If IM% = 0 Then Exit Function
If Ver% = 25 Then
    View% = FindChildByClass(IM%, "_AOL_VIEW")
    Res% = FindChildByTitle(IM%, "Respond")
    AC_Click Res%
    Ed% = FindChildByClass(IM%, "_AOL_EDIT")
    SendBttn% = FindChildByTitle(IM%, "Send")
   Else
    View% = FindChildByClass(IM%, "RICHCNTL")
    Res% = FindChildByClass(IM%, "_AOL_ICON")
    AC_Click Res%
    Ed% = AC_GetAOLWin(IM%, "RICHCNTL", 2)
    SendBttn% = GetWindow(Ed%, GW_HWNDNEXT)
    End If
AC_SetText Ed%, AnsMsg$
AC_IMAnswer$ = CStr(AC_GetWinText(View%))
AC_Click SendBttn%
AC_Close IM%
DoEvents
End Function

Sub AC_IMRoom (Lst As ListBox, ByVal Msg As String)
For i = 0 To Lst.ListCount - 1
    SN$ = Lst.List(i)
    Msg$ = AF_Script(Msg$, "%SN%", SN$)
    Msg$ = AF_Script(Msg$, "%TIME%", Format$(Now, "h:mm AM/PM"))
    Msg$ = AF_Script(Msg$, "%DATE%", Format$(Now, "m-dd-yy"))
    AC_IM SN$, Msg$
    For n = 0 To 7
	DoEvents
	ErrBox% = FindWindow("#32770", "America Online")
	Timeout (.001)
	Next n
    If ErrBox% <> 0 Then
	K% = FindChildByTitle(ErrBox%, "OK")
	AC_Click K%
	End If
    Next i
End Sub

Sub AC_IMsOnOrOff (Choice As String)
If AC_Online() = 0 Then Exit Sub
If LCase(Choice$) = "off" Then FillIn$ = "$IM_OFF"
If LCase(Choice$) = "on" Then FillIn$ = "$IM_ON"
AC_IM FillIn$, "Turning Instant Messages " & Choice$ & "."
Do Until Mes% <> 0
    DoEvents
    Mes% = FindWindow("#32770", "America Online")
    Timeout (.001)
    Loop
MesBttn% = FindChildByTitle(Mes%, "OK")
AC_Click (MesBttn%)
AOL% = FindWindow("AOL Frame25", 0&)
IM% = FindChildByTitle(AOL%, "Send Instant Message")
Q% = SendMessage(IM%, WM_CLOSE, 0, 0&)
DoEvents
End Sub

Function AC_IsOnline () As Integer
AOL% = FindWindow("AOL Frame25", 0&)
Wlcm% = FindChildByTitle(AOL%, "Welcome, ")
Com% = FindChildByClass(Wlcm%, "_AOL_COMBOBOX")
If Com% <> 0 Then Wlcm% = 0
If Wlcm% = 0 Then
    Beep
    MsgBox "You must be signed on to use this program.", 48, app.Path
    AC_IsOnline% = 0
    Exit Function
    End If
AC_IsOnline% = Wlcm%
End Function

Sub AC_Keyword (KW As String)
AOL% = FindWindow("AOL Frame25", 0&)
If AOL% = 0 Then
    Beep
    MsgBox "AOL was not detected (Keyword)", 0, "Error"
    Exit Sub
    End If
Call AC_RunMenuByString("Keyword...", "&Go To")
Do Until key% <> 0
    DoEvents
    key% = FindChildByTitle(AOL%, "Keyword")
    Timeout (.001)
    Loop
Q% = ShowWindow(key%, SW_HIDE): DoEvents
KEdit% = FindChildByClass(key%, "_AOL_EDIT")
Q% = sendmessagebystring(KEdit%, WM_SETTEXT, 0, KW$)
GoBttn% = FindChildByClass(key%, "_AOL_ICON")
AC_Click (GoBttn%)
DoEvents
End Sub

Sub AC_KillModal ()
DoEvents
AOM% = FindWindow("_AOL_MODAL", 0&)
Q% = EnableWindow(AOM%, 0)
AOL% = FindWindow("AOL Frame25", 0&)
Q% = EnableWindow(AOL%, 1)
End Sub

Sub AC_KillPalette ()
Pal% = FindWindow("_AOL_PALETTE", "America Online")
If Pal% <> 0 Then
    ok% = FindChildByTitle(Pal%, "OK")
    Do Until Pal% = 0
	DoEvents
	Pal% = FindWindow("_AOL_PALETTE", "America Online")
	AC_Click (ok%)
	DoEvents
	Timeout (.001)
	Loop
    End If
End Sub

Sub AC_KillWait ()
Q% = ShowCursor(False)
Ver% = AC_AOLVersion()
AOL% = FindWindow("AOL Frame25", 0&)
If Ver% = 25 Then
    Call AC_RunMenuByString("Exit Free Area", "&Go To")
    End If
If Ver% = 3 Then
    Call AC_RunMenuByString("Exit Unlimited Use area", "&Go To")
    DoEvents
    End If
Q% = ShowCursor(True)
End Sub

Function AC_LocateChat () As Integer
'This is a no-fault locate chat room

AOL% = FindWindow("AOL Frame25", 0&)
Do Until Room1% <> 0 And Room1% = Room2% And Room2% = Room3%
    Room1% = GetParent(FindChildByClass(AOL%, "_AOL_LISTBOX"))
    Room2% = GetParent(FindChildByClass(AOL%, "_AOL_VIEW"))
    Room3% = GetParent(FindChildByClass(AOL%, "_AOL_EDIT"))
    Timeout (.01)
    Loop
AC_LocateChat% = Room1%
End Function

Function AC_LocateChatNoLoop () As Integer
DoEvents
AOL% = FindWindow("AOL Frame25", 0&)
startWin% = GetParent(FindChildByClass(AOL%, "_AOL_LISTBOX"))
If AC_GetWinText(startWin%) Like "Buddy List Window" Then startWin% = GetParent(FindChildByClass(AOL%, "_AOL_VIEW"))
Q% = SetFocusAPI(startWin%)
For i = 0 To 6
    DoEvents
    Room1% = GetParent(FindChildByClass(AOL%, "_AOL_LISTBOX"))
    Room2% = GetParent(FindChildByClass(AOL%, "_AOL_VIEW"))
    Room3% = GetParent(FindChildByClass(AOL%, "_AOL_EDIT"))
    If Room1% <> 0 And Room1% = Room2% And Room2% = Room3% Then AC_LocateChatNoLoop% = Room2%: Exit Function
    Timeout (.001)
    Next i
End Function

Function AC_LocateChatSkim () As Integer
AOL% = FindWindow("AOL Frame25", 0&)
BudList% = FindChildByTitle(AOL%, "Buddy List Window")
If BudList% <> 0 Then Q% = SendMessage(BudList%, WM_CLOSE, 0, 0&): DoEvents
Room1% = GetParent(FindChildByClass(AOL%, "_AOL_LISTBOX"))
Room2% = GetParent(FindChildByClass(AOL%, "_AOL_VIEW"))
Room3% = GetParent(FindChildByClass(AOL%, "_AOL_EDIT"))
If Room1% <> 0 And Room1% = Room2% And Room2% = Room3% Then AC_LocateChatSkim% = Room1% Else AC_LocateChatSkim% = 0
End Function

Function AC_MDIClient () As Integer
'Returns the value of AOL's MDICLient window.
AC_MDIClient% = FindChildByClass(AOL%, "MDIClient")
End Function

Function AC_MenuExist (ByVal MenuTitle As String) As Integer
'fixed by HackSmurf

Top_Position_Num = -1
AOL% = FindWindow("AOL Frame25", 0&)
Menu_Handle = getmenu(AOL%)
MenuCount% = GetMenuItemCount(Menu_Handle) '<----Get the menucount
For x% = 1 To MenuCount%
    Top_Position_Num = Top_Position_Num + 1
    Buffer$ = String$(255, 0)
    Look_For_Menu_String% = GetMenuString(Menu_Handle, Top_Position_Num, Buffer$, Len(MenuTitle$) + 1, MF_BYPOSITION)
    Buffer$ = TrimNull(Buffer$)
    If Buffer$ Like MenuTitle$ Then
	AC_MenuExist% = True
	Exit Function
	End If
    Next x%
AC_MenuExist% = False
End Function

Sub AC_NameFix (NewSN As String, AOLDir As String, ReplaceMe As String)
DoEvents
If AC_Online() = True Then MsgBox "You must be offline to use this feature.", 0, "Error": Exit Sub
On Error GoTo NameFixErrHandler
NewSN$ = NewSN$ & Space(10 - Len(NewSN$))
ReplaceMe$ = ReplaceMe$ & Space(10 - Len(ReplaceMe$))
Open AOLDir$ For Binary Access Read Write As #2
    FileL = LOF(2)
    FileCeiling = FileL
    FileStart = 1
While FileCeiling >= 0
If FileCeiling > 32000 Then
	Buffer = 32000
       ElseIf Wonderer = 0 Then
	Buffer = 1
       Else
	Buffer = FileCeiling
	End If
    Buf$ = String$(Buffer, " ")
    Get #2, FileStart, Buf$
    Sing! = InStr(1, Buf$, NewSN$, 1)
    If Sing! Then Mid$(Buf$, Sing!) = ReplaceMe$
    Put #2, FileStart, Buf$
    FileStart = FileStart + Buffer
    FileCeiling = FileL - FileStart
    Wend
Close #2
Exit Sub
NameFixErrHandler:
Resume ErrFix
ErrFix:
End Sub

Function AC_Online () As Integer
AOL% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(AOL%, "Welcome, ")
Com% = FindChildByClass(Wel%, "_AOL_COMBOBOX")
If Com% <> 0 Then Wel% = 0
If Wel% = 0 Then
    MsgBox "You need to be signed onto AOL to use this feature.", 0, "Error"
    AC_Online% = False
    Exit Function
    End If
If Wel% <> 0 Then AC_Online% = True
End Function

Sub AC_OwnHost (Host As String, Mesg As String)
Room% = AC_LocateChatNoLoop()
View% = FindChildByClass(Room%, "_AOL_VIEW")
Final$ = CStr(Chr(13) + Chr(10) + Host$ + ":" + Chr(9) + Mesg$)
Q% = sendmessagebystring(View%, WM_SETTEXT, 0, Final$)
DoEvents
End Sub

Sub AC_PrepIM (SN As String)
If AC_Online() = 0 Then Exit Sub
Ver% = AC_AOLVersion()
AOL% = FindWindow("AOL Frame25", 0&)
AOL% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(AOL%, "Send Instant Message")
If IMWin% = 0 Then
    AC_RunMenuByString "Send an Instant Message", "Mem&bers"
    AC_WaitForChange
    IMWin% = FindChildByTitle(AOL%, "Send Instant Message")
    End If
Q% = SetFocusAPI(IMWin%)
Ed1% = FindChildByClass(IMWin%, "_AOL_EDIT")
Q% = SetFocusAPI(IMWin%)
AC_SetText Ed1%, SN$
DoEvents
End Sub

Sub AC_RemoveMenu (MenuTitle As String)
Top_Position_Num = -1
AOL% = FindWindow("AOL Frame25", 0&)
Menu_Handle = getmenu(AOL%)
Do
    DoEvents
    Top_Position_Num = Top_Position_Num + 1
    Buffer$ = String$(255, 0)
    Look_For_Menu_String% = GetMenuString(Menu_Handle, Top_Position_Num, Buffer$, Len(MenuTitle$) + 1, MF_BYPOSITION)
    Buffer$ = TrimNull(Buffer$)
    If Buffer$ Like MenuTitle$ Then Exit Do
    Loop
Q% = DeleteMenu(Menu_Handle, Top_Position_Num, MF_BYPOSITION)
DrawMenuBar (AOL%)
DoEvents
End Sub

Sub AC_RenameHost (ByVal AOLDir As String, ByVal NewHost As String)
On Error GoTo ErrHandler
If InStr(AOLDir$, "3") <> 0 Then Version = 3 Else Version = 25
If Len(NewHost$) > 14 Then MsgBox "WTF are you tryin' to do?  Fuck up your AOL Software?", 0, "Error": Error 3110
If Version = 3 Then
    chat$ = "aolchat.aol"
    PNum = 4761
   Else
    chat$ = "chat.aol"
    PNum = 6887
    End If
Open "C:\" + AOLDir$ + "\tool\" & chat$ For Binary As #1
Seek #1, PNum
Put #1, , NewHost$
Close #1
Exit Sub
ErrHandler:
MsgBox "Renaming the host was UNSUCCESSFUL.  Please try again.", 0, "Error"
End Sub

Function AC_RoomBust (ByVal Room As String, OffCond As Timer) As Integer
'You need to place this function in a Timer.  It does not
'bust multiple times, but will if you place it in a timer.
'OffCond must be set to the timer the function is placed
'in.
If OffCond.Enabled = False Then Exit Function
AOL% = FindWindow("AOL Frame25", 0&)
If AOL% = 0 Then MsgBox "You must be signed onto AOL to use this option.", 0, "Error": Exit Function
If AC_Online() = 0 Then Exit Function
Room$ = LCase(AF_Script(Room$, " ", ""))
RoomOpen$ = AC_GetWinText(AC_LocateChatSkim())
If RoomOpen$ <> "" Then
    RoomOpen$ = LCase(AF_Script(RoomOpen$, " ", ""))
    If Room$ = RoomOpen$ Then AC_RoomBust% = AC_LocateChatSkim(): Exit Function
    End If
If OffCond.Enabled = False Then Exit Function
AC_Keyword "aol://2719:2-2-" & Room$
Do Until ErrBox% <> 0
    DoEvents
    ErrBox% = 0
    S% = 0
    suc$ = ""
    ErrBox% = FindWindow("#32770", "America Online")
    If ErrBox% <> 0 Then Exit Do
    S% = AC_LocateChatSkim()
    If S% <> 0 Then
	suc$ = LCase(AF_Script(AC_GetWinText(S%), " ", ""))
	If Room$ Like suc$ Then
	    AC_RoomBust% = S%
	    OffCond.Enabled = False
	    Exit Function
	    End If
	S% = 0
	End If
    If OffCond.Enabled = False Then Exit Function
    Timeout (.001)
    Loop
If ErrBox% <> 0 Then
    K% = FindChildByTitle(ErrBox%, "OK")
    AC_Click K%
    AC_RoomBust% = 0
    Exit Function
    End If
End Function

Function AC_RoomCheck () As Integer
Room% = AC_LocateChatNoLoop()
If Room% = 0 Then
    AC_RoomCheck = 0
    MsgBox "You need to be in a room to use this feature.", 0, "Error"
    Exit Function
    End If
AC_RoomCheck% = Room%
End Function

Function AC_RoomCount () As Variant
'Inspired by Trident.  This function returns the number of
'people in a current room.

Room% = AC_LocateChatNoLoop()
If Room% = 0 Then Exit Function
RoomList% = FindChildByClass(Room%, "_AOL_LISTBOX")
AC_RoomCount = SendMessageByNum(RoomList%, LB_GETCOUNT, 0, 0&)
End Function

Function AC_RoomName () As String
'This function returns the name of an open room.

Room% = AC_LocateChatNoLoop()
If Room% = 0 Then AC_RoomName$ = "No Room Detected": Exit Function
AC_RoomName$ = AC_GetWinText(Room%)
End Function

Sub AC_RunMenuByString (Menu_String As String, Top_Position As String)
Top_Position_Num = -1
AOL% = FindWindow("AOL Frame25", 0&)
Menu_Handle = getmenu(AOL%)
Do
    DoEvents
    Top_Position_Num = Top_Position_Num + 1
    Buffer$ = String$(255, 0)
    Look_For_Menu_String% = GetMenuString(Menu_Handle, Top_Position_Num, Buffer$, Len(Top_Position) + 1, MF_BYPOSITION)
    Trim_Buffer = TrimNull(Buffer$)
    If Trim_Buffer = Top_Position Then Exit Do
Loop
Sub_Menu_Handle = GetSubMenu(Menu_Handle, Top_Position_Num)
BY_POSITION = -1
Do
    DoEvents
    BY_POSITION = BY_POSITION + 1
    Buffer$ = String(255, 0)
    Look_For_Menu_String% = GetMenuString(Sub_Menu_Handle, BY_POSITION, Buffer$, Len(Menu_String) + 1, MF_BYPOSITION)
    Trim_Buffer = TrimNull(Buffer$)
    If Trim_Buffer = Menu_String Then Exit Do
Loop
DoEvents
Get_ID% = GetMenuItemID(Sub_Menu_Handle, BY_POSITION)
Click_Menu_Item = SendMessageByNum(AOL%, WM_COMMAND, Get_ID%, 0&)
End Sub

Function AC_RunMenuGetWin (Sub_String As String, Top_String As String, WinCaption As String, WinType As Integer) As Integer
'This function Calls AC_RunMenuByString and then loops
'until it locates the window to be found by it's caption
'"(WinCaption).  You must also specify if the window to
'be found is "_AOL_MODAL" or "AOL Child" or #32770 (MsgBox)
'(WinType).
'If its an AOL Child then WinType = 1.
'If its an _AOL_MODAL then WinType = 2.
'If its a #32770 (MsgBox) then WinType = 3.

'Example:
'IM% = AC_RunMenuGetWin("Mem&bers", "Send an Instant Message", "Send Instant Message", 1)
'Then you'd have the IM window's address in IM%.

AOL% = FindWindow("AOL Frame25", 0&)
Call AC_RunMenuByString(Sub_String$, Top_String$)
Do Until WinFind% <> 0
    DoEvents
    If WinType% = 1 Then WinFind% = FindChildByTitle(AOL%, WinCaption)
    If WinType% = 2 Then WinFind% = FindWindow("_AOL_MODAL", WinCaption)
    If WinType% = 3 Then WinFind% = FindWindow("#32770", WinCaption)
    Timeout (.001)
    Loop
AC_RunMenuGetWin% = WinFind%
End Function

Sub AC_RunToolbar (ByVal NameIndex As String)
'Applicable to AOL3.0 only
Tool% = FindChildByClass(AC_AOL(), "AOL Toolbar")
Select Case UCase(NameIndex$)
    Case "NEWMAIL"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 1)
    Exit Sub
    Case "COMPOSE"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 2)
    Exit Sub
    Case "CHANNEL"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 3)
    Exit Sub
    Case "HOT"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 4)
    Exit Sub
    Case "LOBBY"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 5)
    Exit Sub
    Case "FILESEARCH"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 6)
    Exit Sub
    Case "QUOTES", "STOCKS"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 7)
    Exit Sub
    Case "NEWS"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 8)
    Exit Sub
    Case "WWW"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 9)
    Exit Sub
    Case "SHOP"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 10)
    Exit Sub
    Case "MYAOL"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 11)
    Exit Sub
    Case "TIMER"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 12)
    Exit Sub
    Case "PRINT"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 13)
    Exit Sub
    Case "PFC"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 14)
    Exit Sub
    Case "FAVORITE"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 15)
    Exit Sub
    Case "SERVICES"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 16)
    Exit Sub
    Case "FIND"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 17)
    Exit Sub
    Case "KEYWORD"
    AC_Click AC_GetAOLWin(Tool%, "_AOL_ICON", 18)
    Exit Sub
    End Select
Debug.Print "Incorrect NameIndex String."
End Sub

Function AC_SearchForChat () As Integer
If ChatFocusMode% = 0 Then
    Room% = AC_LocateChatNoLoop()
    If Room% = 0 Then
	MsgBox "The chat room focus was lost.  I will pause until it is found again.", 0, "Error"
	ChatFocusMode% = 1
	AC_SearchForChat% = 0
	End If
    End If
If ChatFocusMode% = 1 Then
    Room% = AC_LocateChatNoLoop()
    If Room% <> 0 Then
	ChatFocusMode% = 0
	AC_SearchForChat% = 1
	End If
    End If
End Function

Sub AC_SetText (Target As Integer, ByVal What As String)
DoEvents
Q% = sendmessagebystring(Target%, WM_SETTEXT, 0, What$)
End Sub

Sub AC_SetToAOL (FRM As Form)
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(AOL%, "MDIClient")
Q% = SetParent(FRM.hWnd, MDI%)
DoEvents
End Sub

Function AC_SetToView () As Integer
'This function finds the chatroom and sets itself to the
'view.

AOL% = FindWindow("AOL Frame25", 0&)
AC_CloseBuddy
Room% = AC_LocateChatNoLoop()
If Room% = 0 Then MsgBox "You need a room open to use this feature.", 0, "Room Location Error": Exit Function
AC_SetToView% = FindChildByClass(Room%, "_AOL_VIEW")
End Function

Sub AC_Show (Target As Integer)
'This Sub will show a hidden target window
Q% = ShowWindow(Target%, SW_SHOW)
DoEvents
End Sub

Sub AC_ShowToolbar ()
AOL% = FindWindow("AOL Frame25", 0&)
Tool% = FindChildByClass(AOL%, "AOL Toolbar")
Q% = ShowWindow(AOL%, SW_HIDE)
Q% = ShowWindow(AOL%, SW_MINIMIZE)
Q% = ShowWindow(Tool%, SW_SHOW)
Q% = ShowWindow(AOL%, SW_MAXIMIZE)
Q% = ShowWindow(AOL%, SW_SHOW)
DoEvents
End Sub

Sub AC_ShowWelcome ()
AOL% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(AOL%, "Welcome, " & AC_GetSN() & "!")
Q% = ShowWindow(Wel%, SW_SHOW)
DoEvents
End Sub

Sub AC_StayOnline ()
If AC_Online() = 0 Then Exit Sub
AC_KillPalette
ModO% = FindWindow("_AOL_MODAL", 0&)
If ModO% <> 0 Then
    y% = FindChildByTitle(ModO%, "&Yes")
    AC_Click y%
    End If
Send ("·")
End Sub

Sub Send (ByVal TXT As String)
Room% = AC_LocateChatNoLoop()
Ed% = FindChildByClass(Room%, "_AOL_EDIT")
StoreTXT$ = AC_GetWinText(Ed%)
SendBttn% = GetWindow(Ed%, GW_HWNDNEXT)
Q% = sendmessagebystring(Ed%, WM_SETTEXT, 0, TXT$)
AC_Click SendBttn%
Q% = sendmessagebystring(Ed%, WM_SETTEXT, 0, StoreTXT$)
DoEvents
End Sub

Sub Send13 (ByVal TXT As String)
Room% = AC_LocateChatNoLoop()
Ed% = FindChildByClass(Room%, "_AOL_EDIT")
StoreTXT$ = AC_GetWinText(Ed%)
Q% = sendmessagebystring(Ed%, WM_SETTEXT, 0, TXT$)
Q% = SendMessageByNum(Ed%, WM_CHAR, Chr(13), 0)
Q% = sendmessagebystring(Ed%, WM_SETTEXT, 0, StoreTXT$)
DoEvents
End Sub

Function SYS_GetCPUType () As String
'Example: Label4.Caption = "Your system's CPU type is: " & sGetCPUType
lWinFlags = GetWinFlags()
If lWinFlags And WF_CPU486 Then
    SYS_GetCPUType = "486"
   ElseIf lWinFlags And WF_CPU386 Then
    SYS_GetCPUType = "386"
   ElseIf lWinFlags And WF_CPU286 Then
    SYS_GetCPUType = "286"
   Else
    SYS_GetCPUType = "Other"
    End If
End Function

Function SYS_GetDOSVersion () As String
Ver = GetVersion()
DosVer = Ver \ &H10000
SYS_GetDOSVersion$ = CStr(Format((DosVer \ 256) + ((DosVer Mod 256) / 100), "Fixed"))
End Function

Function SYS_GetFreeGDI () As String
'Example: text5.text = "Free GDI Resources: " & GetFreeGDI
SYS_GetFreeGDI = Format$(GetFreeSystemResources(GFSR_GDIRESOURCES)) & "%"
End Function

Function SYS_GetFreeSys () As String
'Example: text3.text = "Free System Resources: " & SYS_GetFreeSys
SYS_GetFreeSys = Format$(GetFreeSystemResources(GFSR_SYSTEMRESOURCES)) & "%"
End Function

Function SYS_GetFreeUser () As String
'Example: text55.text = "Free User Resources: " & sGetFreeUser
SYS_GetFreeUser = Format$(GetFreeSystemResources(GFSR_USERRESOURCES)) + "%"
End Function

Function SYS_GetKeyboardType () As String
V% = GetKeyboardType(0)
Select Case V%
    Case 1
    KB$ = "IBM PC/XT"
    Case 2
    KB$ = "Olivetti ICO"
    Case 3
    KB$ = "IBM AT"
    Case 4
    KB$ = "IBM Enhanced"
    Case 5
    KB$ = "Nokia 1050"
    Case 6
    KB$ = "Nokia 9140"
    Case 7
    KB$ = "Japanese Keyboard"
    End Select
If Len(KB$) = 0 Then KB$ = "Keyboard not detected."
SYS_GetKeyboardType$ = KB$
End Function

Function SYS_GetTimeAndDate () As String
SYS_GetTimeAndDate$ = Format$(Now, "h:mm AM/PM mm-dd-yy")
End Function

Function SYS_GetWinVersion () As String
Dim lVer As Long, iWinVer As Integer
    lVer = GetVersion()
    iWinVer = CInt(lVer And &HFFFF&)
    SYS_GetWinVersion$ = Format$(iWinVer And &HFF) + "." + Format$(CInt(iWinVer / 256))
End Function

Function T_ANUBiSCAPS (Strin As String) As String
L% = Len(Strin$)
NumSpc% = 0
Do While NumSpc% <= L%
    Let NumSpc% = NumSpc% + 1
    Let NextChr$ = Mid$(Strin$, NumSpc%, 1)
    If NextChr$ = "i" Or NextChr$ = "I" Then Final$ = Final$ & "i" Else Final$ = Final$ & UCase(NextChr$)
    Loop
T_ANUBiSCAPS$ = Final$
End Function

Function T_Backwards (StringIn As String)
Let inptxt$ = StringIn$
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
    Let NumSpc% = NumSpc% + 1
    Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
    Let Newsent$ = NextChr$ & Newsent$
    Loop
T_Backwards = Newsent$
End Function

Function T_Elite (inputtxt As String)
Let inptxt$ = inputtxt
Let lenth% = Len(inptxt$)

Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If crapp% > 0 Then GoTo dustepp2

If NextChr$ = "A" Then Let NextChr$ = "/\"
If NextChr$ = "a" Then Let NextChr$ = "å"
If NextChr$ = "B" Then Let NextChr$ = "ß"
If NextChr$ = "C" Then Let NextChr$ = "Ç"
If NextChr$ = "c" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "Ð"
If NextChr$ = "d" Then Let NextChr$ = "ð"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "è"
If NextChr$ = "f" Then Let NextChr$ = "ƒ"
If NextChr$ = "H" Then Let NextChr$ = "|-|"
If NextChr$ = "I" Then Let NextChr$ = "‡"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "K" Then Let NextChr$ = "\<"
If NextChr$ = "L" Then Let NextChr$ = "£"
If NextChr$ = "M" Then Let NextChr$ = "(\/)"
If NextChr$ = "m" Then Let NextChr$ = "^^"
If NextChr$ = "N" Then Let NextChr$ = "/\/"
If NextChr$ = "n" Then Let NextChr$ = "ñ"
If NextChr$ = "O" Then Let NextChr$ = "Ø"
If NextChr$ = "o" Then Let NextChr$ = "º"
If NextChr$ = "P" Then Let NextChr$ = "¶"
If NextChr$ = "p" Then Let NextChr$ = "Þ"
If NextChr$ = "r" Then Let NextChr$ = "®"
If NextChr$ = "S" Then Let NextChr$ = "§"
If NextChr$ = "s" Then Let NextChr$ = "š"
If NextChr$ = "t" Then Let NextChr$ = "†"
If NextChr$ = "U" Then Let NextChr$ = "Ú"
If NextChr$ = "u" Then Let NextChr$ = "µ"
If NextChr$ = "V" Then Let NextChr$ = "\/"
If NextChr$ = "W" Then Let NextChr$ = "\\'"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "x" Then Let NextChr$ = "×"
If NextChr$ = "Y" Then Let NextChr$ = "¥"
If NextChr$ = "y" Then Let NextChr$ = "ý"
If NextChr$ = "!" Then Let NextChr$ = "¡"
If NextChr$ = "?" Then Let NextChr$ = "¿"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "‚"
If NextChr$ = "1" Then Let NextChr$ = "¹"
If NextChr$ = "%" Then Let NextChr$ = "‰"
If NextChr$ = "2" Then Let NextChr$ = "²"
If NextChr$ = "3" Then Let NextChr$ = "³"
If NextChr$ = "_" Then Let NextChr$ = "¯"
If NextChr$ = "-" Then Let NextChr$ = "—"
If NextChr$ = " " Then Let NextChr$ = " "
Let Newsent$ = Newsent$ + NextChr$

dustepp2:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
T_Elite = Newsent$
End Function

Function T_Hacker (Strin As String)
Let inptxt$ = Strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "a"
If NextChr$ = "E" Then Let NextChr$ = "e"
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "O" Then Let NextChr$ = "o"
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "B"
If NextChr$ = "c" Then Let NextChr$ = "C"
If NextChr$ = "d" Then Let NextChr$ = "D"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "H"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "J"
If NextChr$ = "k" Then Let NextChr$ = "K"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "m" Then Let NextChr$ = "M"
If NextChr$ = "n" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "X"
If NextChr$ = "p" Then Let NextChr$ = "P"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "r" Then Let NextChr$ = "R"
If NextChr$ = "s" Then Let NextChr$ = "S"
If NextChr$ = "t" Then Let NextChr$ = "T"
If NextChr$ = "w" Then Let NextChr$ = "W"
If NextChr$ = "v" Then Let NextChr$ = "V"
If NextChr$ = " " Then Let NextChr$ = " "
Let Newsent$ = Newsent$ + NextChr$
Loop
T_Hacker = Newsent$
End Function

Function T_InsChr (ByVal Strin As String, ByVal InsMe As String)
'This function Inserts a Character after every character.
'
'Example:
'
'text2.text = "."
'AC_Send ("Change Me!",  text2.text)
'
'That would send "C.h.a.n.g.e. .M.e.!." to the chat room.

Let inptxt$ = Strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ <> " " Then Let NextChr$ = NextChr$ + InsMe
Let Newsent$ = Newsent$ + NextChr$
Loop
LastChr = Len(Newsent$)
T_InsChr = Left(Newsent$, LastChr - 1)
End Function

Function T_Intertwine (ByVal TXT1 As String, ByVal TXT2 As String) As String
'I wrote this function especially for Toast x 64.

If Len(TXT1$) > Len(TXT2$) Then
     Longer$ = TXT1$
     TXT$ = TXT2$
    Else
     Longer$ = TXT2$
     TXT$ = TXT1$
     End If
For i = 0 To Len(Longer$)
     Final$ = Final$ & Mid$(Longer$, i, 1)
     If Len(TXT$) >= i Then Final$ = Final$ & Mid$(TXT$, i, 1)
     Next i
T_Intertwine$ = CStr(Final$)
End Function

Function T_KTEncrypt (ByVal password, ByVal strng, force%) As String
'Example:
'temp = T_KTEncrypt("Paszwerd", CertNum, 0)
'CertNum = temp

'Set error capture routine
On Local Error GoTo ErrorHandler
'Is there Password?
If Len(password) = 0 Then Error 31100
'Is password too long?
If Len(password) > 255 Then Error 31100
'Is there a strng$ to work with?
If Len(strng) = 0 Then Error 31100
'Check if file is encrypted and not forcing
If force% = 0 Then
    'Check for encryption ID tag
    chk$ = Left$(strng, 4) + Right$(strng, 4)
    If chk$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      'Remove ID tag
      strng = Mid$(strng, 5, Len(strng) - 8)
      'String was encrypted so filter out CHR$(1) flags
      look = 1
      Do
	look = InStr(look, strng, Chr$(1))
	If look = 0 Then
	    Exit Do
	   Else
	    Addin$ = Chr$(Asc(Mid$(strng, look + 1)) - 1)
	    strng = Left$(strng, look - 1) + Addin$ + Mid$(strng, look + 2)
	    End If
	look = look + 1
	Loop
      'Since it is encrypted we want to decrypt it
      EncryptFlag% = False
     Else
      'Tag not found so flag to encrypt string
      EncryptFlag% = True
      End If
     Else
    'force% flag set, ecrypt string regardless of tag
      EncryptFlag% = True
      End If
'Set up variables
PassUp = 1
PassMax = Len(password)
'Tack on leading characters to prevent repetative recognition
password = Chr$(Asc(Left$(password, 1)) Xor PassMax) + password
password = Chr$(Asc(Mid$(password, 1, 1)) Xor Asc(Mid$(password, 2, 1))) + password
password = password + Chr$(Asc(Right$(password, 1)) Xor PassMax)
password = password + Chr$(Asc(Right$(password, 2)) Xor Asc(Right$(password, 1)))
'If Encrypting add password check tag now so it is encrypted with string
If EncryptFlag% = True Then
    strng = Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") + strng
    End If
'Loop until scanned though the whole string
For Looper = 1 To Len(strng)
    'Alter character code
    ToChange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(password, PassUp, 1))
    'Insert altered character code
    Mid$(strng, Looper, 1) = Chr$(ToChange)
    'Scroll through password string one character at a time
    PassUp = PassUp + 1
    If PassUp > PassMax + 4 Then PassUp = 1
    Next Looper
'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
If EncryptFlag% = True Then
    'First get rid of all CHR$(1) since that is what we use for our flag
    look = 1
    Do
	look = InStr(look, strng, Chr$(1))
	If look > 0 Then
	strng = Left$(strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(strng, look + 1)
	look = look + 1
	End If
	Loop While look > 0
    'Check for CHR$(0)
    Do
	look = InStr(strng, Chr$(0))
	If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(strng, look + 1)
	Loop While look > 0
    'Check for CHR$(10)
    Do
	look = InStr(strng, Chr$(10))
	If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(strng, look + 1)
	Loop While look > 0
    'Check for CHR$(13)
    Do
	look = InStr(strng, Chr$(13))
	If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(strng, look + 1)
	Loop While look > 0
    'Check for CHR$(26)
    Do
	look = InStr(strng, Chr$(26))
	If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(strng, look + 1)
	Loop While look > 0
    'Tack on encryted tag
	strng = Chr$(1) + "KT" + Chr$(1) + strng + Chr$(1) + "KT" + Chr$(1)
   Else
    'We decrypted so ensure password used was the correct one
    If Left$(strng, 9) <> Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") Then
	'Password bad cause error
	Error 31100
       Else
	'Password good, remove password check tag
	strng = Mid$(strng, 10)
	End If
    End If
'Set function equal to modified string
T_KTEncrypt$ = strng
'We're out of here
Exit Function
ErrorHandler:
'We had an error!  We're out of here.
Exit Function
End Function

Function T_Scramble (Word As String)
Rescramble:
Word = LCase(Word) & "x"
For i = 1 To Len(Word)
    cLet$ = Mid(Word, i, 1)
    If cLet$ = " " Or i = Len(Word) Then
	nWord$ = T_ScrambleArray(cWord$)
	cWord$ = ""
	Final$ = Final$ & nWord$
       Else
	cWord$ = cWord$ & cLet$
	End If
    Next i
If Final$ = Word Then GoTo Rescramble
T_Scramble = Final$
End Function

Function T_ScrambleArray (ScramMe As String) As String
ReDim Storage(0 To Len(ScramMe$)) As String
For i = 0 To Len(ScramMe$)
    Storage(i) = " "
    Next i
L = Len(ScramMe$)
If L = 1 Then T_ScrambleArray$ = ScramMe$: Exit Function
Egg$ = LCase(ScramMe$)
i = 0
Do Until Len(Egg$) = 0
i = i + 1
Nope:
    Randomize
    x = Int(L * Rnd)
    If Storage(x) <> " " Then GoTo Nope
    Letter$ = Left(Egg$, 1)
    Storage(x) = Letter$
    Egg$ = Right(Egg$, L - i)
    Loop
Egg$ = ""
For x = 0 To L
    Egg$ = T_Backwards(Egg$ & Storage(x))
    Next x
If LCase(T_Backwards(Egg$)) = LCase(ScramMe$) Then GoTo Nope
T_ScrambleArray$ = Egg$
End Function

Function T_SoundDecode (ByVal InTXT As String) As String
InTXT$ = LCase(InTXT$ & " ")
InTXT$ = AF_Script(InTXT$, "aigh ", "a")
InTXT$ = AF_Script(InTXT$, "bee ", "b")
InTXT$ = AF_Script(InTXT$, "see ", "c")
InTXT$ = AF_Script(InTXT$, "dee ", "d")
InTXT$ = AF_Script(InTXT$, "eff ", "f")
InTXT$ = AF_Script(InTXT$, "gee ", "g")
InTXT$ = AF_Script(InTXT$, "aytch ", "h")
InTXT$ = AF_Script(InTXT$, "eye ", "i")
InTXT$ = AF_Script(InTXT$, "jay ", "j")
InTXT$ = AF_Script(InTXT$, "kay ", "k")
InTXT$ = AF_Script(InTXT$, "el ", "l")
InTXT$ = AF_Script(InTXT$, "em ", "m")
InTXT$ = AF_Script(InTXT$, "en ", "n")
InTXT$ = AF_Script(InTXT$, "oh ", "o")
InTXT$ = AF_Script(InTXT$, "pee ", "p")
InTXT$ = AF_Script(InTXT$, "cue ", "q")
InTXT$ = AF_Script(InTXT$, "are ", "r")
InTXT$ = AF_Script(InTXT$, "ess ", "s")
InTXT$ = AF_Script(InTXT$, "tee ", "t")
InTXT$ = AF_Script(InTXT$, "you ", "u")
InTXT$ = AF_Script(InTXT$, "vee ", "v")
InTXT$ = AF_Script(InTXT$, "doubleyou ", "w")
InTXT$ = AF_Script(InTXT$, "ecks ", "x")
InTXT$ = AF_Script(InTXT$, "why ", "y")
InTXT$ = AF_Script(InTXT$, "zee ", "z")
InTXT$ = AF_Script(InTXT$, "ee ", "e")
T_SoundDecode$ = Trim$(InTXT$)
End Function

Function T_SoundEncode (ByVal InTXT As String) As String
For i = 1 To Len(InTXT$)
    L$ = LCase(Mid$(InTXT$, i, 1))
    Select Case L$
	Case "a"
	n$ = "aigh"
	Case "b"
	n$ = "bee"
	Case "c"
	n$ = "see"
	Case "d"
	n$ = "dee"
	Case "e"
	n$ = "ee"
	Case "f"
	n$ = "eff"
	Case "g"
	n$ = "gee"
	Case "h"
	n$ = "aytch"
	Case "i"
	n$ = "eye"
	Case "j"
	n$ = "jay"
	Case "k"
	n$ = "kay"
	Case "l"
	n$ = "el"
	Case "m"
	n$ = "em"
	Case "n"
	n$ = "en"
	Case "o"
	n$ = "oh"
	Case "p"
	n$ = "pee"
	Case "q"
	n$ = "cue"
	Case "r"
	n$ = "are"
	Case "s"
	n$ = "ess"
	Case "t"
	n$ = "tee"
	Case "u"
	n$ = "you"
	Case "v"
	n$ = "vee"
	Case "w"
	n$ = "doubleyou"
	Case "x"
	n$ = "ecks"
	Case "y"
	n$ = "why"
	Case "z"
	n$ = "zee"
	Case Else
	n$ = L$
	End Select
    Final$ = Final$ & n$ & " "
    Next i
T_SoundEncode$ = Trim$(Final$)
End Function

Function T_Spaced (Strin As String)
Let inptxt$ = Strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + " "
Let Newsent$ = Newsent$ + NextChr$
Loop
T_Spaced = Newsent$
End Function

Sub AC_AddRoom (Lst As Control)
If AC_LocateChatSkim() = 0 Then Exit Sub
ListB% = FindChildByClass(AC_LocateChatSkim(), "_AOL_LISTBOX")
Lst.Clear
UserSN$ = AC_GetSN()
For index% = 0 To 50 'Conference Room
    SN$ = String$(20, " ")
    Q% = SetFocusAPI(ListB%)
    Q% = AOLGetList(index%, SN$)
    If Len(Trim$(SN$)) <= 2 Then Exit For
    SN$ = TrimNull(SN$)
    If SN$ <> UserSN$ Then Lst.AddItem (SN$)
    Next index%
End Sub

Sub AC_AddRoomNoClear (Lst As Control)
If AC_LocateChatNoLoop() = 0 Then Exit Sub
UserSN$ = AC_GetSN()
For index% = 0 To 50 'Conference Room
    SN$ = String$(20, " ")
    Q% = AOLGetList(index%, SN$)
    If Len(Trim$(SN$)) <= 2 Then Exit For
    SN$ = TrimNull(SN$)
    If SN$ <> UserSN$ Then Lst.AddItem (SN$)
    Next index%
End Sub

Sub AC_AlterWelcomeWindow (ByVal newCaption As String, ByVal aolToday As String, ByVal staticOne As String, ByVal staticTwo As String, ByVal staticThree As String, ByVal topNewsStory As String)
'Reinspired by LoKi, original concept by Anubis.
'Worx with AOL 3.0 and 2.5.  Note that using this eliminates
'use of AC_GetSN which gets the user's screen name.
AOL% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(AOL%, "Welcome")
If Wel% = 0 Then MsgBox "You must be signed onto AOL to use this feature.": Exit Sub
aVer = AC_AOLVersion()
If aVer = 3 Then
    DoEvents
    Stat1% = FindChildByClass(Wel%, "RICHCNTL")
    Stat2% = AC_GetAOLWin(Wel%, "RICHCNTL", 2)
    Stat3% = AC_GetAOLWin(Wel%, "RICHCNTL", 3)
    Stat4% = AC_GetAOLWin(Wel%, "RICHCNTL", 4)
    Stat5% = AC_GetAOLWin(Wel%, "RICHCNTL", 5)
    AC_SetText Wel%, newCaption$
    AC_SetText Stat1%, aolToday$
    AC_SetText Stat2%, staticOne$
    AC_SetText Stat3%, staticTwo$
    AC_SetText Stat4%, staticThree$
    AC_SetText Stat5%, CStr("TOP NEWS STORY:" & AF_Enter() & topNewsStory$)
    End If
End Sub

Function AC_AOL () As Integer
'This locates AOL
AC_AOL% = FindWindow("AOL Frame25", 0&)
DoEvents
End Function

Sub AC_AOL4Free ()
'This is pretty much the code that Happy Hardcore used
'to reveal all the windows while you were in a free area
'if he was on a PC.

AOL% = FindWindow("AOL Frame25", 0&)
StartPoint% = FindChildByClass(AOL%, "AOL Child")
Q% = ShowWindow(StartPoint%, SW_SHOW): DoEvents
PrevWin% = StartPoint%
Do Until LastWin% = StartPoint%
    DoEvents
    LastWin% = GetWindow(PrevWin%, GW_HWNDNEXT)
    Q% = ShowWindow(LastWin%, SW_SHOW): DoEvents
    PrevWin% = LastWin%
    Loop
End Sub

Function AC_AOLMODAL (WinCaption As String) As Integer
'This function locates a lot of windows that are not act-
'ually part of AOL at all like the sign-off window and the
'"Edit Go To Window", "Preferences" and so on.
'
'Example:
'MailPref% = AC_AOLMODEL("Mail Preferences")

For i = 1 To 10
    CheckMe% = FindWindow("_AOL_MODAL", WinCaption$)
    DoEvents
    Timeout (.001)
    If CheckMe% <> 0 Then AC_AOLMODAL% = CheckMe%: Exit For: Exit Function
    Next i
End Function

Sub Timeout (ByVal Duration As Double)
StartTime = Timer
Do While Timer - StartTime < Duration
DoEvents
Loop
End Sub

Function TrimNull (ByVal in) As String
For x = 1 To Len(in)
    If (Mid$(in, x, 1) <> Chr$(0)) Then
    total$ = total$ + Mid$(in, x, 1)
    Else
    GoTo NullDetect
    End If
Next
NullDetect:
TrimNull = total$
End Function

