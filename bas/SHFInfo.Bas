Attribute VB_Name = "ModSHFileInfo"
Option Explicit

Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, _
                                                           ByVal x As Long, _
                                                           ByVal y As Long, _
                                                           ByVal hIcon As Long) As Boolean

Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, _
                                                              ByVal xLeft As Long, _
                                                              ByVal yTop As Long, _
                                                              ByVal hIcon As Long, _
                                                              ByVal cxWidth As Long, _
                                                              ByVal cyWidth As Long, _
                                                              ByVal istepIfAniCur As Long, _
                                                              ByVal hbrFlickerFreeDraw As Long, _
                                                              ByVal diFlags As Long) As Boolean
' DrawIconEx() diFlags values:
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8

Declare Function SHGetFileInfo Lib "Shell32" Alias "SHGetFileInfoA" _
                              (ByVal pszPath As Any, _
                              ByVal dwFileAttributes As Long, _
                              psfi As SHFILEINFO, _
                              ByVal cbFileInfo As Long, _
                              ByVal uFlags As Long) As Long

' pszPath:
' Pointer to a buffer that contains the path and filename. Both absolute and
' relative paths are valid. If uFlags includes the SHGFI_PIDL, value pszPath
' must be the address of an ITEMIDLIST structure that contains the list of
' item identifiers that uniquely identifies the file within the shell's name space.
' This string can use either short (the 8.3 form) or long filenames.

' dwFileAttributes:
' Array of file attribute flags (FILE_ATTRIBUTE_ values). If uFlags does not
' include the SHGFI_USEFILEATTRIBUTES value, this parameter is ignored.

Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800

' psfi and cbFileInfo:
' Address and size, in bytes, of the SHFILEINFO structure that receives the file information.

' Maximun long filename path length
Public Const MAX_PATH = 260

Type SHFILEINFO   ' shfi
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

' uFlags:
' Flag that specifies the file information to retrieve. This parameter can
' be a combination of the following values:


' Modifies SHGFI_ICON, causing the function to retrieve the file's large icon.
Public Const SHGFI_LARGEICON = &H0&

' Modifies SHGFI_ICON, causing the function to retrieve the file's small icon.
Public Const SHGFI_SMALLICON = &H1&

' Modifies SHGFI_ICON, causing the function to retrieve the file's open icon.
' A container object displays an open icon to indicate that the container is open.
Public Const SHGFI_OPENICON = &H2&

' Modifies SHGFI_ICON, causing the function to retrieve a shell-sized icon.
' If this flag is not specified, the function sizes the icon according to the system metric values.
Public Const SHGFI_SHELLICONSIZE = &H4&

' Indicates that pszPath is the address of an ITEMIDLIST structure rather than a path name.
Public Const SHGFI_PIDL = &H8&

' Indicates that the function should use the dwFileAttributes parameter.
Public Const SHGFI_USEFILEATTRIBUTES = &H10&

' Retrieves the handle of the icon that represents the file and the index of the
' icon within the system image list. The handle is copied to the hIcon member
' of the structure specified by psfi, and the index is copied to the iIcon member.
' The return value is the handle of the system image list.
Public Const SHGFI_ICON = &H100&

' Retrieves the display name for the file. The name is copied to the szDisplayName
' member of the structure specified by psfi. The returned display name uses the
' long filename, if any, rather than the 8.3 form of the filename.
Public Const SHGFI_DISPLAYNAME = &H200&

' Retrieves the string that describes the file's type. The string is copied to the
' szTypeName member of the structure specified by psfi.
Public Const SHGFI_TYPENAME = &H400&

' Retrieves the file attribute flags. The flags are copied to the dwAttributes
' member of the structure specified by psfi.
Public Const SHGFI_ATTRIBUTES = &H800&

' Retrieves the name of the file that contains the icon representing the file.
' The name is copied to the szDisplayName member of the structure specified by psfi.
Public Const SHGFI_ICONLOCATION = &H1000&

' Returns the type of the executable file if pszPath identifies an executable file.
' To retrieve the executable file type, uFlags must specify only SHGFI_EXETYPE.
' The return value specifies the type of the executable file:
' 0                                                                       Nonexecutable file or an error condition.
' LOWORD = NE or PEHIWORD = 3.0, 3.5, or 4.0  Windows application
' LOWORD = MZHIWORD = 0                               MS-DOS .EXE, .COM or .BAT file
' LOWORD = PEHIWORD = 0                               Win32 console application
Public Const SHGFI_EXETYPE = &H2000&

' Retrieves the index of the icon within the system image list. The index is copied to the iIcon
' member of the structure specified by psfi. The return value is the handle of the system image list.
Public Const SHGFI_SYSICONINDEX = &H4000&

' Modifies SHGFI_ICON, causing the function to add the link overlay to the file's icon.
Public Const SHGFI_LINKOVERLAY = &H8000&

' Modifies SHGFI_ICON, causing the function to blend the file's icon with the system highlight color.
Public Const SHGFI_SELECTED = &H10000

