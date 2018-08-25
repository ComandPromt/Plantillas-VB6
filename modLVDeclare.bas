Attribute VB_Name = "modLVDeclare"
Option Explicit

Public Const BIF_RETURNONLYFSDIRS = &H1

Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000  'system icon index
Public Const SHGFI_LARGEICON = &H0        'large icon
Public Const SHGFI_SMALLICON = &H1        'small icon
Public Const ILD_TRANSPARENT = &H1        'display transparent
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or _
             SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or _
             SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Const MAX_PATH = 260
Public Type SHITEMID
  cb      As Long
  abID    As Byte
End Type

Public Type ITEMIDLIST
  mkid    As SHITEMID
End Type

Public Type BROWSEINFO
  hOwner          As Long
  pidlRoot        As Long
  pszDisplayName  As String
  lpszTitle       As String
  ulFlags         As Long
  lpfn            As Long
  lParam          As Long
  iImage          As Long
End Type

Public Type SHFILEINFO
   hIcon          As Long
   iIcon          As Long
   dwAttributes   As Long
   szDisplayName  As String * MAX_PATH
   szTypeName     As String * 80
End Type

Public shinfo As SHFILEINFO

Public Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type

Public Type SYSTEMTIME
  wYear             As Integer
  wMonth            As Integer
  wDayOfWeek        As Integer
  wDay              As Integer
  wHour             As Integer
  wMinute           As Integer
  wSecond           As Integer
  wMilliseconds     As Long
End Type

Public Type WIN32_FIND_DATA
  dwFileAttributes  As Long
  ftCreationTime    As FILETIME
  ftLastAccessTime  As FILETIME
  ftLastWriteTime   As FILETIME
  nFileSizeHigh     As Long
  nFileSizeLow      As Long
  dwReserved0       As Long
  dwReserved1       As Long
  cFileName         As String * MAX_PATH
  cAlternate        As String * 14
End Type

Public Declare Function SHGetPathFromIDList Lib _
   "shell32.dll" Alias "SHGetPathFromIDListA" _
   (ByVal pidl As Long, _
    ByVal pszPath As String) As Long

Public Declare Function SHBrowseForFolder Lib _
   "shell32.dll" Alias "SHBrowseForFolderA" _
   (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
   (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
   (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Declare Function FileTimeToSystemTime Lib "kernel32" _
   (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
   
Public Declare Function UpdateWindow Lib "user32" _
      (ByVal hwnd As Long) As Long

Public Declare Function ImageList_Draw Lib "comctl32.dll" _
   (ByVal himl&, _
    ByVal i&, _
    ByVal hDCDest&, _
    ByVal x&, _
    ByVal Y&, _
    ByVal flags&) As Long

Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_GRIDLINES = &H1
Public Const LVS_EX_CHECKBOXES As Long = &H4
Public Const LVS_EX_HEADERDRAGDROP = &H10
Public Const LVS_EX_TRACKSELECT = &H8
Public Const LVS_EX_ONECLICKACTIVATE = &H40
Public Const LVS_EX_TWOCLICKACTIVATE = &H80
Public Const LVS_EX_SUBITEMIMAGES = &H2

Public Const LVIF_STATE = &H8
 
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)

Public Const LVIS_STATEIMAGEMASK As Long = &HF000

Public Type LVITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   state        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type

Public Const LVM_GETCOLUMN = (LVM_FIRST + 25)
Public Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Public Const LVCF_TEXT = &H4

Public Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText  As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type

'hHeaderFont is the handle to the font used to draw the
'header text, and must not be destroyed unless no longer
'needed (see the Unload event).
Public hHeaderFont As Long
   
'vars representing the checkbox options in
'the chkHeaderFont control array.
Public Const optBold = 0
Public Const optItalic = 1
Public Const optUnderlined = 2
Public Const optStrikeout = 3
   
'-----------------------------------------------
'APIs, constants and structures required to change the listview header font
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
      
'font weight vars
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
   
'SendMessage vars
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31

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
   lfFaceName(LF_FACESIZE) As Byte
End Type

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" _
   (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Declare Function SelectObject Lib "gdi32" _
   (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" _
   (ByVal hObject As Long) As Long

Public Declare Function CreateFontIndirect Lib "gdi32" _
    Alias "CreateFontIndirectA" _
   (lpLogFont As LOGFONT) As Long

' Flat Column Headers
Public Const HDS_BUTTONS = &H2
Public Const GWL_STYLE = (-16)
Const SWP_DRAWFRAME = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
  
Public Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hwnd As Long, _
   ByVal nIndex As Long) As Long
   
Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
   
Public Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public Const SW_NORMAL = 1
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

