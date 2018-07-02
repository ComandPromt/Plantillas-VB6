Attribute VB_Name = "TimeSetupDeclarations"
Option Explicit

'These are the WIN32API declarations needed for setting the
'Time and Time Zone Information
Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

'   And these are the functions that use the above structures
Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Function SetLocalTime Lib "kernel32" (lpLocalTime As SYSTEMTIME) As Long
Declare Function SetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

'   This is to get the ID of the Locale Settings
Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

'   These are for Regional Settings Control Panel
Public Const LOCALE_STIMEFORMAT = &H1003       'Time Format string

'   And this is the Win32API Function
Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean

'   This is used to write to WIN.INI
Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

'   These are for the Win32API SendMesssage Function
Public Const WM_SETTINGCHANGE = &H1A       ' same as the old WM_WININICHANGE
Public Const HWND_BROADCAST = &HFFFF&      ' Special HWND value for use with
                                           ' PostMessage and SendMessage
'   And this is the Win32API Function
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
