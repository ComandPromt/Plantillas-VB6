Attribute VB_Name = "basImageView"
Option Explicit

Private mhLargeSysIL    As Long     'Handle to System ImageList (Large Icons)
Private mhSmallSysIL    As Long     'Handle to System ImageList (Small Icons)

Private Const MAX_PATH                  As Long = &H104&        '260    - Max Path Len
Private Const GWL_STYLE                 As Long = &HFFFFFFF0    '(-16)  - Get/Set Window Style
Private Const SHGFI_SYSICONINDEX        As Long = &H4000&       '16384  - Get System Icon Index
Private Const SHGFI_LARGEICON           As Long = &H0&          '0      - Get Large Icon
Private Const SHGFI_SMALLICON           As Long = &H1&          '1      - Get Small Icon
Private Const SHGFI_DISPLAYNAME         As Long = &H200&        '512    - Get File Display Name
Private Const SHGFI_TYPENAME            As Long = &H400&        '1024   - Get File Type Name
Private Const SHGFI_ICON                As Long = &H100&        '256    - Get icon
Private Const LVIF_IMAGE                As Long = &H2&          '2      - Setting the Image
Private Const LVM_SETIMAGELIST          As Long = &H1003        '4099   - Set New Image List
Private Const LVM_SETITEM               As Long = &H1006        '4102   - Set Image List Item
Private Const LVM_SETCOLUMNWIDTH        As Long = &H101E&       '4126   - Set ListView Column Width
Private Const LVS_SHAREIMAGELISTS       As Long = &H40&         '64     - Don't Destroy Assigned Image Lists
Private Const LVSIL_NORMAL              As Long = &H0&          '0      - Large Icon
Private Const LVSIL_SMALL               As Long = &H1&          '1      - Small Icon
Private Const LVSCW_AUTOSIZE            As Long = &HFFFFFFFF    '(-1)   - Autosize ListView Column
Private Const LVSCW_AUTOSIZE_USEHEADER  As Long = &HFFFFFFFE    '(-2)   - Autosize ListView Column to Header
Private Const INVALID_HANDLE_VALUE      As Long = &HFFFFFFFF    '(-1)   - File not found

Private Const FILE_ATTRIBUTE_READONLY   As Long = &H1&          '1      - Read Only File
Private Const FILE_ATTRIBUTE_HIDDEN     As Long = &H2&          '2      - Hidden File
Private Const FILE_ATTRIBUTE_SYSTEM     As Long = &H4&          '4      - System File
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10&         '16     - Folder
Private Const FILE_ATTRIBUTE_ARCHIVE    As Long = &H20&         '32     - Archive File

Private Type SHFILEINFO
    hIcon           As Long                 'Icon handle
    iIcon           As Long                 'Icon index
    dwAttributes    As Long                 'SFGAO_flags
    szDisplayName   As String * MAX_PATH    'Display name (or path)
    szTypeName      As String * 80          'Type name
End Type

Private Type LV_ITEM
    mask        As Long
    iItem       As Long
    iSubItem    As Long
    State       As Long
    stateMask   As Long
    pszText     As String
    cchTextMax  As Long
    iImage      As Long
    lParam      As Long '(~ ItemData)
    iIndent     As Long
End Type

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMillisecs          As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal hIml As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Public Sub AssignSystemImageLists(ByVal sPath As String, lvwFiles As ListView)
    
'Retieve handles to the System ImageLists
'and assign them to the Listview.

Dim lIdx    As Long
Dim lStyle  As Long
Dim sfiFile As SHFILEINFO
    
    'Get handles to system image lists.
    'They may change, especially if the user changes display settings.
    'Safest to obtain them each time...
    mhLargeSysIL = SHGetFileInfo(sPath, 0&, sfiFile, Len(sfiFile), SHGFI_SYSICONINDEX Or SHGFI_LARGEICON)
    mhSmallSysIL = SHGetFileInfo(sPath, 0&, sfiFile, Len(sfiFile), SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
    
    'Make sure that the System ImageLists aren't destroyed when the ImageList is terminated
    lStyle = GetWindowLong(lvwFiles.hWnd, GWL_STYLE)
    Call SetWindowLong(lvwFiles.hWnd, GWL_STYLE, lStyle Or LVS_SHAREIMAGELISTS)
    
    'Assign the System ImageLists to the ListView
    Call SendMessage(lvwFiles.hWnd, LVM_SETIMAGELIST, LVSIL_NORMAL, ByVal mhLargeSysIL)
    Call SendMessage(lvwFiles.hWnd, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal mhSmallSysIL)
    
    If lvwFiles.ListItems.Count > 0 And lvwFiles.View = lvwReport Then
        'Resize the headers for report view
        For lIdx = 0 To lvwFiles.ColumnHeaders.Count - 2
            Call SendMessage(lvwFiles.hWnd, LVM_SETCOLUMNWIDTH, lIdx, LVSCW_AUTOSIZE)
        Next
    End If
    
End Sub

Public Sub FillFileList(filFiles As FileListBox, lvwFiles As ListView)

'Populate the listview with the contents of
'large or small system imagelist icons.

Dim lIdx        As Long
Dim lCnt        As Long
Dim lPos        As Long
Dim lRet        As Long
Dim hIml        As Long             'System ImageList Handle
Dim lFlags      As Long             'SHGetFileInfo flags
Dim lAttr       As Long             'File attributes
Dim hFind       As Long             'Find Handle
Dim sPath       As String           'Current Path
Dim sName       As String           'Filename
Dim sAttr       As String           'File attributes string
Dim dDateTime   As Date             'File Date/Time
Dim lvwItem     As LV_ITEM          'API ListView.ListItem
Dim oItem       As ListItem         'VB ListView.ListItem
Dim ftTime      As FILETIME         'Local file date/time
Dim stTime      As SYSTEMTIME       'System file date/time
Dim sfiFile     As SHFILEINFO       'SHGetFileInfo structure
Dim FindData    As WIN32_FIND_DATA  'FindData structure
    
    Screen.MousePointer = vbHourglass
    
    'Clear the ListView
    lvwFiles.ListItems.Clear
    
    'Assign the System ImageLists to the ListView
    Call AssignSystemImageLists(filFiles.Path, lvwFiles)
    
    'Setup the Path and Flags
    sPath = filFiles.Path & IIf(Right$(filFiles.Path, 1) <> "\", "\", "")
    lFlags = SHGFI_ICON Or SHGFI_DISPLAYNAME Or SHGFI_TYPENAME Or SHGFI_SYSICONINDEX
    
    For lIdx = 0 To filFiles.ListCount - 1
        hIml = SHGetFileInfo(sPath & filFiles.List(lIdx), &H0&, sfiFile, Len(sfiFile), lFlags)
        If hIml <> 0 Then
            'Don't need the Icon handle, so free the memory
            If sfiFile.hIcon <> 0 Then
                lRet = DestroyIcon(sfiFile.hIcon)
            End If
            'Get the display name
            lPos = InStr(sfiFile.szDisplayName, Chr$(0))
            If lPos > 1 Then
                sName = Left$(sfiFile.szDisplayName, lPos - 1)
            Else
                sName = ""
            End If
            If Len(sName) > 0 Then
                Set oItem = lvwFiles.ListItems.Add(, , sName)
                lvwItem.iItem = lCnt            'Index of ListView item (zero-based)
                lvwItem.iImage = sfiFile.iIcon  'Index in System ImageList Icon (zero-based)
                lvwItem.mask = LVIF_IMAGE       'Only setting the Image index
                'Assign the Icon's image index to the ListView Item
                Call SendMessage(lvwFiles.hWnd, LVM_SETITEM, &H0&, lvwItem)
                hFind = FindFirstFile(sPath & sName, FindData)
                If hFind <> INVALID_HANDLE_VALUE Then
                    'File Size
                    If FindData.nFileSizeLow <= 1024 Then
                        oItem.SubItems(1) = "1 KB"
                    Else
                        oItem.SubItems(1) = Format$(FindData.nFileSizeLow / 1024, "#,##0 KB")
                    End If
                    
                    'File Type
                    oItem.SubItems(2) = Left$(sfiFile.szTypeName, InStr(1, sfiFile.szTypeName, vbNullChar) - 1)
        
                    'Last Modified (Translate to local system time)
                    Call FileTimeToLocalFileTime(FindData.ftLastWriteTime, ftTime)
                    Call FileTimeToSystemTime(ftTime, stTime)
                    dDateTime = DateSerial(stTime.wYear, stTime.wMonth, stTime.wDay) + TimeSerial(stTime.wHour, stTime.wMinute, stTime.wSecond)
                    'Use default system date/time format
                    oItem.SubItems(3) = Format$(dDateTime, "Short Date") & " " & Format$(dDateTime, "Medium Time")
                    
                    'Attributes
                    lAttr = FindData.dwFileAttributes
                    sAttr = IIf((lAttr And FILE_ATTRIBUTE_READONLY) > 0, "R", "") _
                        & IIf((lAttr And FILE_ATTRIBUTE_HIDDEN) > 0, "H", "") _
                        & IIf((lAttr And FILE_ATTRIBUTE_SYSTEM) > 0, "S", "") _
                        & IIf((lAttr And FILE_ATTRIBUTE_ARCHIVE) > 0, "A", "")
                    oItem.SubItems(4) = sAttr
                    
                    'Close the Find
                    FindClose hFind
                End If
                lCnt = lCnt + 1
            End If
        End If
        If lCnt Mod 50 = 0 Then
            DoEvents
        End If
    Next

    If lvwFiles.ListItems.Count > 0 And lvwFiles.View = lvwReport Then
        'Resize the headers for report view (This activates the autosize
        'for each column in the list.
        For lIdx = 0 To lvwFiles.ColumnHeaders.Count - 2
            Call SendMessage(lvwFiles.hWnd, LVM_SETCOLUMNWIDTH, lIdx, LVSCW_AUTOSIZE)
        Next
    End If

    DoEvents
    lvwFiles.Refresh
    Screen.MousePointer = vbNormal
    
End Sub


