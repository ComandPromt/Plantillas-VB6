Attribute VB_Name = "WIN32API"
'
' WIN32 CONSTANTS
'
Public Const MCI_SYSINFO = &H810
Public Const MCI_SYSINFO_INSTALLNAME = &H800&
Public Const MCI_SYSINFO_NAME = &H400&
Public Const MCI_SYSINFO_OPEN = &H200&
Public Const MCI_SYSINFO_QUANTITY = &H100&
Public Const MCI_ALL_DEVICE_ID = -1
Public Const MAX_COMPUTERNAME_LENGTH = 15
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
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
Global Const VK_NUMLOCK = &H90
Global Const VK_SCROLL = &H91
Global Const WM_USER = &H400
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
Public Const COLOR_3DDKSHADOW = 21
Public Const COLOR_3DLIGHT = 22
Public Const COLOR_INFOBK = 24
Public Const COLOR_INFOTEXT = 23

Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CXVSCROLL = 2
Public Const SM_CYHSCROLL = 3
Public Const SM_CYCAPTION = 4
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYVTHUMB = 9
Public Const SM_CXHTHUMB = 10
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
Public Const SM_CXCURSOR = 13
Public Const SM_CYCURSOR = 14
Public Const SM_CYMENU = 15
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYKANJIWINDOW = 18
Public Const SM_MOUSEPRESENT = 19
Public Const SM_CYVSCROLL = 20
Public Const SM_CXHSCROLL = 21
Public Const SM_DEBUG = 22
Public Const SM_SWAPBUTTON = 23
Public Const SM_RESERVED1 = 24
Public Const SM_RESERVED2 = 25
Public Const SM_RESERVED3 = 26
Public Const SM_RESERVED4 = 27
Public Const SM_CXMIN = 28
Public Const SM_CYMIN = 29
Public Const SM_CXSIZE = 30
Public Const SM_CYSIZE = 31
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CXMINTRACK = 34
Public Const SM_CYMINTRACK = 35
Public Const SM_CXDOUBLECLK = 36
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CXICONSPACING = 38
Public Const SM_CYICONSPACING = 39
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_PENWINDOWS = 41
Public Const SM_DBCSENABLED = 42
Public Const SM_CMOUSEBUTTONS = 43
Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Public Const SM_CXSIZEFRAME = SM_CXFRAME
Public Const SM_CYSIZEFRAME = SM_CYFRAME
Public Const SM_SECURE = 44
Public Const SM_CXEDGE = 45
Public Const SM_CYEDGE = 46
Public Const SM_CXMINSPACING = 47
Public Const SM_CYMINSPACING = 48
Public Const SM_CXSMICON = 49
Public Const SM_CYSMICON = 50
Public Const SM_CYSMCAPTION = 51
Public Const SM_CXSMSIZE = 52
Public Const SM_CYSMSIZE = 53
Public Const SM_CXMENUSIZE = 54
Public Const SM_CYMENUSIZE = 55
Public Const SM_ARRANGE = 56
Public Const SM_CXMINIMIZED = 57
Public Const SM_CYMINIMIZED = 58
Public Const SM_CXMAXTRACK = 59
Public Const SM_CYMAXTRACK = 60
Public Const SM_CXMAXIMIZED = 61
Public Const SM_CYMAXIMIZED = 62
Public Const SM_NETWORK = 63
Public Const SM_CLEANBOOT = 67
Public Const SM_CXDRAG = 68
Public Const SM_CYDRAG = 69
Public Const SM_SHOWSOUNDS = 70
Public Const SM_CXMENUCHECK = 71
Public Const SM_CYMENUCHECK = 72
Public Const SM_SLOWMACHINE = 73
Public Const SM_MIDEASTENABLED = 74
Public Const SM_CMETRICS = 75

Public Const VER_PLATFORM_WIN32_NT& = 2
Public Const VER_PLATFORM_WIN32_WINDOWS& = 1
Public Const VER_PLATFORM_WIN32S& = 0

Public Const SPI_GETACCESSTIMEOUT& = 60
Public Const SPI_GETANIMATION& = 72
Public Const SPI_GETBEEP& = 1
Public Const SPI_GETBORDER& = 5
Public Const SPI_GETDEFAULTINPUTLANG& = 89
Public Const SPI_GETDRAGFULLWINDOWS& = 38
Public Const SPI_GETFASTTASKSWITCH& = 35
Public Const SPI_GETFILTERKEYS& = 50
Public Const SPI_GETFONTSMOOTHING& = 74
Public Const SPI_GETGRIDGRANULARITY& = 18
Public Const SPI_GETHIGHCONTRAST& = 66
Public Const SPI_GETICONMETRICS& = 45
Public Const SPI_GETICONTITLELOGFONT& = 31
Public Const SPI_GETICONTITLEWRAP& = 25
Public Const SPI_GETKEYBOARDDELAY& = 22
Public Const SPI_GETKEYBOARDPREF& = 68
Public Const SPI_GETKEYBOARDSPEED& = 10
Public Const SPI_GETLOWPOWERACTIVE& = 83
Public Const SPI_GETLOWPOWERTIMEOUT& = 79
Public Const SPI_GETMENUDROPALIGNMENT& = 27
Public Const SPI_GETMOUSE& = 3
Public Const SPI_GETMINIMIZEDMETRICS& = 43
Public Const SPI_GETMOUSEKEYS& = 54
Public Const SPI_GETMOUSETRAILS& = 94
Public Const SPI_GETNONCLIENTMETRICS& = 41
Public Const SPI_GETPOWEROFFACTIVE& = 84
Public Const SPI_GETPOWEROFFTIMEOUT& = 80
Public Const SPI_GETSCREENREADER& = 70
Public Const SPI_GETSCREENSAVEACTIVE& = 16
Public Const SPI_GETSCREENSAVETIMEOUT& = 14
Public Const SPI_GETSERIALKEYS& = 62
Public Const SPI_GETSHOWSOUNDS& = 56
Public Const SPI_GETSOUNDSENTRY& = 64
Public Const SPI_GETSTICKYKEYS& = 58
Public Const SPI_GETTOGGLEKEYS& = 52
Public Const SPI_GETWINDOWSEXTENSION& = 92
Public Const SPI_GETWORKAREA& = 48
Public Const SPI_ICONHORIZONTALSPACING& = 13
Public Const SPI_ICONVERTICALSPACING& = 24
Public Const SPI_LANGDRIVER& = 12
Public Const SPI_SCREENSAVERRUNNING& = 97
Public Const SPI_SETACCESSTIMEOUT& = 61
Public Const SPI_SETANIMATION& = 73
Public Const SPI_SETBEEP& = 2
Public Const SPI_SETBORDER& = 6
Public Const SPI_SETCURSORS& = 87
Public Const SPI_SETDEFAULTINPUTLANG& = 90
Public Const SPI_SETDESKPATTERN& = 21
Public Const SPI_SETDESKWALLPAPER& = 20
Public Const SPI_SETDOUBLECLICKTIME& = 32
Public Const SPI_SETDOUBLECLKHEIGHT& = 30
Public Const SPI_SETDOUBLECLKWIDTH& = 29
Public Const SPI_SETDRAGFULLWINDOWS& = 37
Public Const SPI_SETDRAGHEIGHT& = 77
Public Const SPI_SETDRAGWIDTH& = 76
Public Const SPI_SETFASTTASKSWITCH& = 36
Public Const SPI_SETFILTERKEYS& = 51
Public Const SPI_SETFONTSMOOTHING& = 75
Public Const SPI_SETGRIDGRANULARITY& = 19
Public Const SPI_SETHANDHELD& = 78
Public Const SPI_SETHIGHCONTRAST& = 67
Public Const SPI_SETICONMETRICS& = 46
Public Const SPI_SETICONS& = 88
Public Const SPI_SETICONTITLELOGFONT& = 34
Public Const SPI_SETICONTITLEWRAP& = 26
Public Const SPI_SETKEYBOARDDELAY& = 23
Public Const SPI_SETKEYBOARDPREF& = 69
Public Const SPI_SETKEYBOARDSPEED& = 11
Public Const SPI_SETLANGTOGGLE& = 91
Public Const SPI_SETLOWPOWERACTIVE& = 85
Public Const SPI_SETLOWPOWERTIMEOUT& = 81
Public Const SPI_SETMENUDROPALIGNMENT& = 28
Public Const SPI_SETMINIMIZEDMETRICS& = 44
Public Const SPI_SETMOUSE& = 4
Public Const SPI_SETMOUSEBUTTONSWAP& = 33
Public Const SPI_SETMOUSEKEYS& = 55
Public Const SPI_SETMOUSETRAILS& = 93
Public Const SPI_SETNONCLIENTMETRICS& = 42
Public Const SPI_SETPENWINDOWS& = 49
Public Const SPI_SETPOWEROFFACTIVE& = 86
Public Const SPI_SETPOWEROFFTIMEOUT& = 82
Public Const SPI_SETSCREENREADER& = 71
Public Const SPI_SETSCREENSAVEACTIVE& = 17
Public Const SPI_SETSCREENSAVETIMEOUT& = 15
Public Const SPI_SETSERIALKEYS& = 63
Public Const SPI_SETSHOWSOUNDS& = 57
Public Const SPI_SETSOUNDSENTRY& = 65
Public Const SPI_SETSTICKYKEYS& = 59
Public Const SPI_SETTOGGLEKEYS& = 53
Public Const SPI_SETWORKAREA& = 47
Public Const SPIF_UPDATEINIFILE = 1
Public Const SPIF_SENDWININICHANGE = 2

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const EWX_POWEROFF = 8

Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064

Public Const LOCALE_SYSTEM_DEFAULT = &H800
Public Const LOCALE_USER_DEFAULT = &H400
Public Const LOCALE_ILANGUAGE = &H1         '  language id
Public Const LOCALE_SLANGUAGE = &H2         '  localized name of language
Public Const LOCALE_SENGLANGUAGE = &H1001      '  English name of language
Public Const LOCALE_SABBREVLANGNAME = &H3         '  abbreviated language name
Public Const LOCALE_SNATIVELANGNAME = &H4         '  native name of language
Public Const LOCALE_ICOUNTRY = &H5         '  country code
Public Const LOCALE_SCOUNTRY = &H6         '  localized name of country
Public Const LOCALE_SENGCOUNTRY = &H1002      '  English name of country
Public Const LOCALE_SABBREVCTRYNAME = &H7         '  abbreviated country name
Public Const LOCALE_SNATIVECTRYNAME = &H8         '  native name of country
Public Const LOCALE_IDEFAULTLANGUAGE = &H9         '  default language id
Public Const LOCALE_IDEFAULTCOUNTRY = &HA         '  default country code
Public Const LOCALE_IDEFAULTCODEPAGE = &HB         '  default code page
Public Const LOCALE_SLIST = &HC         '  list item separator
Public Const LOCALE_IMEASURE = &HD         '  0 = metric, 1 = US
Public Const LOCALE_SDECIMAL = &HE         '  decimal separator
Public Const LOCALE_STHOUSAND = &HF         '  thousand separator
Public Const LOCALE_SGROUPING = &H10        '  digit grouping
Public Const LOCALE_IDIGITS = &H11        '  number of fractional digits
Public Const LOCALE_ILZERO = &H12        '  leading zeros for decimal
Public Const LOCALE_SNATIVEDIGITS = &H13        '  native ascii 0-9
Public Const LOCALE_SCURRENCY = &H14        '  local monetary symbol
Public Const LOCALE_SINTLSYMBOL = &H15        '  intl monetary symbol
Public Const LOCALE_SMONDECIMALSEP = &H16        '  monetary decimal separator
Public Const LOCALE_SMONTHOUSANDSEP = &H17        '  monetary thousand separator
Public Const LOCALE_SMONGROUPING = &H18        '  monetary grouping
Public Const LOCALE_ICURRDIGITS = &H19        '  # local monetary digits
Public Const LOCALE_IINTLCURRDIGITS = &H1A        '  # intl monetary digits
Public Const LOCALE_ICURRENCY = &H1B        '  positive currency mode
Public Const LOCALE_INEGCURR = &H1C        '  negative currency mode
Public Const LOCALE_SDATE = &H1D        '  date separator
Public Const LOCALE_STIME = &H1E        '  time separator
Public Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Public Const LOCALE_SLONGDATE = &H20        '  long date format string
Public Const LOCALE_STIMEFORMAT = &H1003      '  time format string
Public Const LOCALE_IDATE = &H21        '  short date format ordering
Public Const LOCALE_ILDATE = &H22        '  long date format ordering
Public Const LOCALE_ITIME = &H23        '  time format specifier
Public Const LOCALE_ICENTURY = &H24        '  century format specifier
Public Const LOCALE_ITLZERO = &H25        '  leading zeros in time field
Public Const LOCALE_IDAYLZERO = &H26        '  leading zeros in day field
Public Const LOCALE_IMONLZERO = &H27        '  leading zeros in month field
Public Const LOCALE_S1159 = &H28        '  AM designator
Public Const LOCALE_S2359 = &H29        '  PM designator
Public Const LOCALE_SDAYNAME1 = &H2A        '  long name for Monday
Public Const LOCALE_SDAYNAME2 = &H2B        '  long name for Tuesday
Public Const LOCALE_SDAYNAME3 = &H2C        '  long name for Wednesday
Public Const LOCALE_SDAYNAME4 = &H2D        '  long name for Thursday
Public Const LOCALE_SDAYNAME5 = &H2E        '  long name for Friday
Public Const LOCALE_SDAYNAME6 = &H2F        '  long name for Saturday
Public Const LOCALE_SDAYNAME7 = &H30        '  long name for Sunday
Public Const LOCALE_SABBREVDAYNAME1 = &H31        '  abbreviated name for Monday
Public Const LOCALE_SABBREVDAYNAME2 = &H32        '  abbreviated name for Tuesday
Public Const LOCALE_SABBREVDAYNAME3 = &H33        '  abbreviated name for Wednesday
Public Const LOCALE_SABBREVDAYNAME4 = &H34        '  abbreviated name for Thursday
Public Const LOCALE_SABBREVDAYNAME5 = &H35        '  abbreviated name for Friday
Public Const LOCALE_SABBREVDAYNAME6 = &H36        '  abbreviated name for Saturday
Public Const LOCALE_SABBREVDAYNAME7 = &H37        '  abbreviated name for Sunday
Public Const LOCALE_SMONTHNAME1 = &H38        '  long name for January
Public Const LOCALE_SMONTHNAME2 = &H39        '  long name for February
Public Const LOCALE_SMONTHNAME3 = &H3A        '  long name for March
Public Const LOCALE_SMONTHNAME4 = &H3B        '  long name for April
Public Const LOCALE_SMONTHNAME5 = &H3C        '  long name for May
Public Const LOCALE_SMONTHNAME6 = &H3D        '  long name for June
Public Const LOCALE_SMONTHNAME7 = &H3E        '  long name for July
Public Const LOCALE_SMONTHNAME8 = &H3F        '  long name for August
Public Const LOCALE_SMONTHNAME9 = &H40        '  long name for September
Public Const LOCALE_SMONTHNAME10 = &H41        '  long name for October
Public Const LOCALE_SMONTHNAME11 = &H42        '  long name for November
Public Const LOCALE_SMONTHNAME12 = &H43        '  long name for December
Public Const LOCALE_SABBREVMONTHNAME1 = &H44        '  abbreviated name for January
Public Const LOCALE_SABBREVMONTHNAME2 = &H45        '  abbreviated name for February
Public Const LOCALE_SABBREVMONTHNAME3 = &H46        '  abbreviated name for March
Public Const LOCALE_SABBREVMONTHNAME4 = &H47        '  abbreviated name for April
Public Const LOCALE_SABBREVMONTHNAME5 = &H48        '  abbreviated name for May
Public Const LOCALE_SABBREVMONTHNAME6 = &H49        '  abbreviated name for June
Public Const LOCALE_SABBREVMONTHNAME7 = &H4A        '  abbreviated name for July
Public Const LOCALE_SABBREVMONTHNAME8 = &H4B        '  abbreviated name for August
Public Const LOCALE_SABBREVMONTHNAME9 = &H4C        '  abbreviated name for September
Public Const LOCALE_SABBREVMONTHNAME10 = &H4D        '  abbreviated name for October
Public Const LOCALE_SABBREVMONTHNAME11 = &H4E        '  abbreviated name for November
Public Const LOCALE_SABBREVMONTHNAME12 = &H4F        '  abbreviated name for December
Public Const LOCALE_SABBREVMONTHNAME13 = &H100F
Public Const LOCALE_SPOSITIVESIGN = &H50        '  positive sign
Public Const LOCALE_SNEGATIVESIGN = &H51        '  negative sign
Public Const LOCALE_IPOSSIGNPOSN = &H52        '  positive sign position
Public Const LOCALE_INEGSIGNPOSN = &H53        '  negative sign position
Public Const LOCALE_IPOSSYMPRECEDES = &H54        '  mon sym precedes pos amt
Public Const LOCALE_IPOSSEPBYSPACE = &H55        '  mon sym sep by space from pos amt
Public Const LOCALE_INEGSYMPRECEDES = &H56        '  mon sym precedes neg amt
Public Const LOCALE_INEGSEPBYSPACE = &H57        '  mon sym sep by space from neg amt

Public Const HKEY_LOCAL_MACHINE = &H80000002

Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCINVERT = &H660046


'-------------------------------------------------------------------------------------
' WIN32 TYPES
'-------------------------------------------------------------------------------------

Public Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type MCI_SYSINFO_PARMS
        dwCallback As Long
        lpstrReturn As String
        dwRetSize As Long
        dwNumber As Long
        wDeviceType As Long
End Type

Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        wProcessorLevel As Integer
        wProcessorRevision As Integer
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Type FREQ_INFO
    in_cycles As Long    ' Internal clock cycles during
    ex_ticks As Long     ' Microseconds elapsed during
    raw_freq As Long     ' Raw frequency of CPU in MHz
    norm_freq As Long    ' Normalized frequency of CPU
End Type

Public Type tBitmap
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
    Width As Integer
    Height As Integer
End Type

'-------------------------------------------------------------------------------------
' WIN32 API FUNCTIONS
'-------------------------------------------------------------------------------------

' system functions
Public Declare Function ProcessorCount Lib "vbCPUInf.dll" () As Long

' graphics functions
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Long) As Integer
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function IntersectRect Lib "User" (ResultRect As tBitmap, Rect1 As tBitmap, Rect2 As tBitmap) As Integer

' multimedia functions
Public Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszNull As Long, ByVal uFlags As Long) As Long

' system information functions
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As DEVMODE) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Public Declare Function IsProcessorFeaturePresent Lib "wnaspi32" (ByVal ProcessorFeature As Long) As Boolean
Public Declare Function GetEnvironmentStrings Lib "kernel32" Alias "GetEnvironmentStringsA" () As String
Public Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

' ini file functions
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' drive information functions
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

' windows tasklist functions
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

' mouse cursor functions
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Global myVer As OSVERSIONINFO
Global Point As POINTAPI

' private ini file routines
Public Sub SaveProfile(ByVal Section$, ByVal Key$, ByVal value$, ByVal fileINI$)
    Dim ret&
    ret = WritePrivateProfileString(Section, Key, value, fileINI)
End Sub

Public Function LoadProfile(ByVal Section$, ByVal Key$, ByVal fileINI$) As String
    Dim ret&, temp$
    temp = Space$(80)
    ret = GetPrivateProfileString(Section, Key, "", temp$, 80, fileINI)
    If ret <> 0 Then
        LoadProfile = temp$
    Else
        LoadProfile = ""
    End If
End Function


