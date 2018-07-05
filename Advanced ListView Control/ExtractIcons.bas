Attribute VB_Name = "MExtractIcons"
Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O

'Icon Sizes in pixels
Public Const LARGE_ICON As Integer = 32
Public Const SMALL_ICON As Integer = 16
Public Const MAX_PATH = 260

Public Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Public Const SHGFI_LARGEICON = &H0       'Large icon
Public Const SHGFI_SMALLICON = &H1       'Small icon
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400

Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type


'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Public Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal X&, ByVal Y&, ByVal flags&) As Long

Public ShInfo As SHFILEINFO



