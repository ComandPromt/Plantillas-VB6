Attribute VB_Name = "Win95_Functions"

Option Explicit
Public Type SHFILEOPSTRUCT

     hWnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAborted As Boolean
     hNameMaps As Long
     sProgress As String

End Type
Public Type BrowseInfo

     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long

End Type
Global FileDestination As String
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
     (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" _
     (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" _
     (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_CONFIRMMOUSE = &H2
Public Const FOF_FILESONLY = &H80                  '  on *.*, do only files
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Public Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_SILENT = &H4                      '  don't create progress/report
Public Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
Public Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings

Public Declare Function SHFileOperation Lib "shell32.dll" Alias _
     "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long


Public Function ShellRename(ParamArray vntFileName() As Variant) As Long


     Dim i As Integer
     Dim sFileNames As String
     Dim Dick As String
     Dim SHFileOp As SHFILEOPSTRUCT

     For i = LBound(vntFileName) To UBound(vntFileName)

          sFileNames = sFileNames & vntFileName(i) & vbNullChar

     Next
     sFileNames = sFileNames & vbNullChar
     Dick = FileDestination
     With SHFileOp

          .wFunc = &H4
          .pFrom = sFileNames
          .fFlags = FOF_ALLOWUNDO
          .pTo = Dick
     End With

     ShellRename = SHFileOperation(SHFileOp)


End Function




Public Function ShellCopy(ParamArray vntFileName() As Variant) As Long


     Dim i As Integer
     Dim sFileNames As Variant
     Dim Dick As String
     Dim SHFileOp As SHFILEOPSTRUCT

     For i = LBound(vntFileName) To UBound(vntFileName)

          sFileNames = sFileNames & vntFileName(i) & vbNullChar

     Next
     sFileNames = sFileNames & vbNullChar
     Dick = FileDestination
     With SHFileOp

          .wFunc = &H2
          .pFrom = sFileNames
          .fFlags = FOF_ALLOWUNDO
          .pTo = Dick
     End With

     ShellCopy = SHFileOperation(SHFileOp)


End Function




Public Function ShellMove(ParamArray vntFileName() As Variant) As Long


     Dim i As Integer
     Dim sFileNames As Variant
     Dim Dick As String
     Dim SHFileOp As SHFILEOPSTRUCT

     For i = LBound(vntFileName) To UBound(vntFileName)

          sFileNames = sFileNames & vntFileName(i) & vbNullChar

     Next
     sFileNames = sFileNames & vbNullChar
     Dick = FileDestination
     With SHFileOp

          .wFunc = &H1
          .pFrom = sFileNames
          .fFlags = FOF_ALLOWUNDO
          .pTo = Dick
     End With

     ShellMove = SHFileOperation(SHFileOp)


End Function



Public Function ShellDelete(ParamArray vntFileName() As Variant) As Long


     Dim i As Integer
     Dim sFileNames As String
     Dim SHFileOp As SHFILEOPSTRUCT

     For i = LBound(vntFileName) To UBound(vntFileName)

          sFileNames = sFileNames & vntFileName(i) & vbNullChar

     Next
     sFileNames = sFileNames & vbNullChar

     With SHFileOp

          .wFunc = FO_DELETE
          .pFrom = sFileNames
          .fFlags = FOF_ALLOWUNDO

     End With

     ShellDelete = SHFileOperation(SHFileOp)


End Function


Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String


     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim sPath As String
     Dim udtBI As BrowseInfo

     With udtBI

          .hwndOwner = hwndOwner
          .lpszTitle = lstrcat(sPrompt, "")
          .ulFlags = BIF_RETURNONLYFSDIRS

     End With

     lpIDList = SHBrowseForFolder(udtBI)
     If lpIDList Then

          sPath = String$(MAX_PATH, 0)
          lResult = SHGetPathFromIDList(lpIDList, sPath)
          Call CoTaskMemFree(lpIDList)
          iNull = InStr(sPath, vbNullChar)
          If iNull Then

               sPath = Left$(sPath, iNull - 1)

          End If

     End If

     BrowseForFolder = sPath


End Function
