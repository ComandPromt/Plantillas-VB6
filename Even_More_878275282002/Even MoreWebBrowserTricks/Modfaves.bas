Attribute VB_Name = "ModFaves"
'*********Copyright PSST Software 2001**********************
'Written by MrBobo - enjoy
'Internet Explorer's Favorites Menu and Treeview
'***********************************************************

'Internet Explorer Dialog declare
Private Declare Function DoOrganizeFavDlg Lib "shdocvw.dll" (ByVal hwnd As Long, ByVal lpszRootFolder As String) As Long
'File handling API
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (Prop As SHELLEXECUTEINFO) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Public Const FO_MOVE = &H1
Public Const FO_COPY = &H2
Public Const FOF_SILENT = &H4
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_ALLOWUNDO = &H40
Public Const FO_DELETE = &H3
Public Const FO_RENAME = &H4&
Dim FOF_FLAGS As Long
Dim SHFileOp As SHFILEOPSTRUCT
Dim FO_FUNC As Long
'Browse for folders
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Type BrowseInfo
  hwndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As Long
  lpszTitle      As Long
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BFFM_INITIALIZED = 1
'Menu API to create and manage Favorites menu
Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_REMOVE = &H1000&
Private Const MF_POPUP = &H10&
Private Const MF_STRING = &H0&
Private Const GWL_WNDPROC = (-4)
Private Const WM_COMMAND = &H111
Private Const MF_BITMAP = &H4&
Private Const WM_CLOSE = &H10
'INI APIs for parsing .url files
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'APIs to locate favorites folder
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Dim ret As String
Dim Retlen As String
Dim bbfaves As String
Dim lngMenu As Long, lngNewMenu As Long, lngNewSubMenu As Long
Dim RootCount As Long
Public gOldProc As Long
Dim LinkURLColl As Collection 'Holds URLs for sublassing calls
Public BrowDlg As New SHDocVw.ShellUIHelper 'used to call dialogs
Dim TV As TreeView
Public Sub WriteINI(FileName As String, Section As String, Key As String, Text As String)
    'for writing Internet shortcuts
    WritePrivateProfileString Section, Key, Text, FileName
End Sub
Public Function ReadINI(FileName As String, Section As String, Key As String)
    'to get addresses from Internet shortcuts
    ret = Space$(255)
    Retlen = GetPrivateProfileString(Section, Key, "", ret, Len(ret), FileName)
    ret = Left$(ret, Retlen)
    ReadINI = ret
End Function
Public Sub GetFaves(mFormHwnd As Long, Optional mTV As TreeView = Nothing)
    If Not mTV Is Nothing Then Set TV = mTV
    bbfaves = SpecialFolder(6) + "\" 'User's favorites folder
    Set LinkURLColl = New Collection
    lngMenu& = GetMenu(mFormHwnd) 'handle to forms main menu
    lngNewMenu& = CreatePopupMenu 'new menu please
    'The numbers "1097" etc. below can be anything - I've set them this high
    'to avoid conflicting with existing menus in your app
    'If you have more than 1096 menu items in your app (as if)
    'then you'll need to increase these numbers accordingly !!
    Call InsertMenu(lngMenu&, 4&, MF_POPUP Or MF_STRING Or MF_BYPOSITION, lngNewMenu&, "Favorites") 'here it is
    AddMenu lngNewMenu&, 1&, 1097, "Add to Favorites..." 'add first three items
    AddMenu lngNewMenu&, 2&, 1098, "Organize Favorites..."
    Call InsertMenu(lngNewMenu&, 3&, MF_SEPARATOR Or MF_BYPOSITION, 1099, vbNullString)
    If Not TV Is Nothing Then TV.Nodes.add , , bbfaves, "Favorites", 1, 2 'initialise treeview
    ListSubDirs bbfaves, lngNewMenu&, bbfaves, True 'recurse through favorites directory - see function below
    ListFiles bbfaves, lngNewMenu&, bbfaves, RootCount 'get any files in root directory (favorites)
    
    '*******important******************
    'this is the hook on the menu
    'comment out these two lines for debugging or you'll lock up VB IDE
    gOldProc& = GetWindowLong(mFormHwnd, GWL_WNDPROC)
    Call SetWindowLong(mFormHwnd, GWL_WNDPROC, AddressOf MenuProc)
    '************************************
    If Not TV Is Nothing Then
        If TV.Nodes.Count > 0 Then
            TV.Nodes(1).Expanded = True
            TV.Nodes(1).Selected = True
        End If
    End If
End Sub
Public Function SpecialFolder(ByVal CSIDL As Long) As String
    'locate the favorites folder
    Dim R As Long
    Dim sPath As String
    Dim IDL As ITEMIDLIST
    Const NOERROR = 0
    Const MAX_LENGTH = 260
    R = SHGetSpecialFolderLocation(GetDesktopWindow, CSIDL, IDL)
    If R = NOERROR Then
        sPath = Space$(MAX_LENGTH)
        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        If R Then
            SpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
        End If
    End If
End Function
Private Function MenuProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'callback to recieve messages from the favorites menu and respond
    Dim z As Long
    Select Case wMsg&
        Case WM_CLOSE:
            Call SetWindowLong(hwnd&, GWL_WNDPROC, gOldProc&)
        Case WM_COMMAND:
            If wParam& > 1100 Then
                z = RunMenu(LinkURLColl(wParam& - 1100), hwnd) 'navigate - see sub below
            End If
            If wParam& = 1097 Then
                z = AddFaves(FormFromHwnd(hwnd))   'first menu item so show Add dialog
            End If
            If wParam& = 1098 Then
                z = OrgFaves(hwnd) 'second menu item so show Organize dialog
            End If
    End Select
    MenuProc = CallWindowProc(gOldProc&, hwnd&, wMsg&, wParam&, lParam&)
End Function
Private Sub ListSubDirs(Path As String, parent As Long, parentDir As String, Optional IsRoot As Boolean = False)
    On Error Resume Next
    If Right(parentDir, 1) <> "\" Then parentDir = parentDir + "\"
    Dim Count, D() As String, i As Long, DirName As String, nSub() As Long, nPos As Long
    DirName = Dir(Path, 16)
    Do While DirName <> ""
        If DirName <> "." And DirName <> ".." Then
            If GetAttr(Path + DirName) = 16 Then 'a folder
                If (Count Mod 10) = 0 Then
                    ReDim Preserve D(Count + 10)
                End If
                Count = Count + 1
                D(Count) = DirName
            End If
        End If
        DirName = Dir
    Loop
    If IsRoot Then 'doing first folder so allow for the first thee items
        RootCount = Count + 3
        nPos = 3
    End If
    'these will be the menu handles of subfolders
    'we need to remember these so we can add the correct links to the correct menus
    'see ListFiles below
    ReDim nSub(1 To Count)
    
    For i = 1 To Count
        nPos = nPos + 1
        nSub(i) = AddSubMenu(parent, nPos, D(i)) 'create a menu(folder)
        If Not TV Is Nothing Then TV.Nodes.add parentDir, tvwChild, Path & D(i) & "\", D(i), 1, 2 'add a node to the treeview
        ListSubDirs Path & D(i) & "\", nSub(i), Path & D(i) & "\" 'recurse any subfolders
        ListFiles Path & D(i) & "\", nSub(i), Path & D(i) & "\", nPos 'add any files held within current folder
    Next
    DoEvents
End Sub
Private Sub ListFiles(Path As String, parent As Long, parentDir As String, Optional StartCnt As Long = 1)
    On Error Resume Next
    Dim Count As Long, D(), i, DirName As String
    DirName = Dir(Path, 6)
    Count = StartCnt
    Do While DirName <> ""
        If DirName <> "." And DirName <> ".." Then
            LinkURLColl.add Path + DirName 'remember location
            AddMenu parent, Count, LinkURLColl.Count + 1100, Left(DirName, Len(DirName) - 4) 'add file to correct menu (handle=parent)
            If Not TV Is Nothing Then TV.Nodes.add parentDir, tvwChild, Path & DirName, Left(DirName, Len(DirName) - 4), 3, 3
            Count = Count + 1
        End If
        DirName = Dir
    Loop
End Sub
Private Function AddSubMenu(mParent As Long, mCount As Long, mname As String) As Long
    lngNewSubMenu = CreatePopupMenu
    If Len(mname) > 49 Then mname = Left(mname, 47) + "..." 'shorten long captions
    Call InsertMenu(mParent, mCount, MF_STRING Or MF_BYPOSITION Or MF_POPUP, lngNewSubMenu, mname)
    AddSubMenu = lngNewSubMenu
End Function
Private Sub AddMenu(mParent As Long, mCount As Long, mID As Long, mname As String)
    If Len(mname) > 49 Then mname = Left(mname, 47) + "..." 'shorten long captions
    Call InsertMenu(mParent, mCount, MF_STRING Or MF_BYPOSITION, mID, mname)
End Sub
Public Sub RefreshFaves(mFormHwnd As Long)
    Dim z As Long, ret As Long, mSmenu As Long
    Call SetWindowLong(mFormHwnd, GWL_WNDPROC, gOldProc&)
    lngMenu& = GetMenu(mFormHwnd)
    lngNewMenu = GetSubMenu(lngMenu&, 2)
    RemoveMenu lngMenu&, 4, MF_BYPOSITION Or MF_REMOVE 'kill the menu
    If Not TV Is Nothing Then
        TV.Nodes.Clear 'kill the treeview
        GetFaves mFormHwnd, TV  'reload
    Else
        GetFaves mFormHwnd
    End If
    DrawMenuBar mFormHwnd 'refresh the form's menu bar
End Sub
Public Function BrowseForFolder(owner As Long) As String
    Dim lpIDList As Long 'show the dialog
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    With tBrowseInfo
        .pIDLRoot = 6 'use favorites folder as root
        .hwndOwner = owner
        .lpszTitle = lstrcat("Move to...", "")
        .ulFlags = BIF_RETURNONLYFSDIRS
        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(260)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
    Else
        BrowseForFolder = ""
    End If
End Function
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    Select Case uMsg
      Case BFFM_INITIALIZED
            SetWindowText hwnd, "Favorites" 'put a caption on the dialog
    End Select
    BrowseCallbackProc = 0
End Function
Private Function GetAddressofFunction(add As Long) As Long
  GetAddressofFunction = add
End Function
'API file operations
Public Function MoveFave(sSource As String, sDestination As String) As Long
    On Error Resume Next
    sSource = sSource & Chr$(0) & Chr$(0)
    With SHFileOp
        .wFunc = 1
        .pFrom = sSource
        .pTo = sDestination
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
    End With
    MoveFave = SHFileOperation(SHFileOp)
End Function
Public Function RenameFave(sSource As String, sDestination As String) As Long
    On Error Resume Next
    sSource = sSource & Chr$(0) & Chr$(0)
    With SHFileOp
        .wFunc = 4
        .pFrom = sSource
        .pTo = sDestination
        .fFlags = FOF_NOCONFIRMATION Or FOF_RENAMEONCOLLISION Or FOF_SILENT
    End With
    RenameFave = SHFileOperation(SHFileOp)
End Function
Public Function DeleteFave(sSource As String) As Long
    On Error Resume Next
    sSource = sSource & Chr$(0) & Chr$(0)
    With SHFileOp
        .wFunc = 3
        .pFrom = sSource
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
    End With
    DeleteFave = SHFileOperation(SHFileOp)
End Function
Public Sub GetPropDlg(frm As Form, mfile As String)
    Dim Prop As SHELLEXECUTEINFO
    Dim R As Long
    With Prop
        .cbSize = Len(Prop)
        .fMask = &HC
        .hwnd = frm.hwnd
        .lpVerb = "properties"
        .lpFile = mfile
    End With
    R = ShellExecuteEx(Prop) 'show dialog
End Sub
Public Function RunMenu(mPath As String, mFormHwnd As Long) As Long
    Dim temp As String, z As VbFileAttribute, ParentForm As Form
    On Error GoTo woops
    Set ParentForm = FormFromHwnd(mFormHwnd)
    If FileExists(mPath) Then
        Select Case LCase(ExtOnly(mPath))
            Case "url" 'navigate
                temp = ReadINI(mPath, "InternetShortcut", "URL")
                If ParentForm.Brow.LocationURL <> temp Then ParentForm.Brow.Navigate temp
            Case "lnk" 'run
                'this will run 99% of links
                'example - fails to execute a link to my dial-up connection
                ShellExecute 0&, vbNullString, mPath, vbNullString, vbNullString, vbNormalFocus
        End Select
    End If
woops:
    RunMenu = 0
End Function

Public Function OrgFaves(mFormHwnd As Long) As Long
    On Error GoTo woops
    LockWindowUpdate mFormHwnd
    DoOrganizeFavDlg mFormHwnd, SpecialFolder(6) 'show dialog
    RefreshFaves mFormHwnd
woops:
    LockWindowUpdate 0
    OrgFaves = 0
End Function

Public Function AddFaves(ParentForm As Form)
    On Error GoTo woops
    LockWindowUpdate ParentForm.hwnd
    BrowDlg.AddFavorite ParentForm.Brow.LocationURL, ChangeExt(ParentForm.Brow.LocationName) 'show dialog
    RefreshFaves ParentForm.hwnd
woops:
    LockWindowUpdate 0
    AddFaves = 0
End Function

Public Function FormFromHwnd(mHwnd As Long) As Form
    Dim frm As Form
    For Each frm In Forms
        If frm.hwnd = mHwnd Then
            Set FormFromHwnd = frm
            Exit For
        End If
    Next
End Function
