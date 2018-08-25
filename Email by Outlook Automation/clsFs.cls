VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************************************
'FileSystemObject Wrapper Functions & OS Version Information
'Dependencies: Windows Scripting Runtime Library (scrrun.dll)
'Author(s): Matthew Hood Email: DragonWeyrDev@Yahoo.com / Unknow - www.planetsourcecode.com
'*****************************************************************************************
'*****************************************************************************************

'*****************************************************************************************
'Class Constants
'*****************************************************************************************
Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

'*****************************************************************************************
'Class Types
'*****************************************************************************************
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Private Type FILETIME
  dwLowDateTime As Long
  dwHightDateTime As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

'*****************************************************************************************
'Class API Declarations
'*****************************************************************************************
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function Searchpath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

'*****************************************************************************************
'Public Functions
'*****************************************************************************************

'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : BuildPath
' Params :
'          ByVal Path As String
'          ByVal Name As String
' Returns: String
' Desc   : The Function uses parameters ByVal Path As String and ByVal Name As String for BuildPath and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function BuildPath(ByVal Path As String, ByVal Name As String) As String
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  BuildPath = fso.BuildPath(Path, Name)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : CopyFile
' Params :
'          ByVal Source As String
'          ByVal Destination As String
'          Optional OverWriteFiles As Boolean = False
' Returns: Boolean
' Desc   : The Function uses parameters ByVal Source As String, ByVal Destination As String and Optional OverWriteFiles As Boolean = False for CopyFile and returns Boolean.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function CopyFile(ByVal Source As String, ByVal Destination As String, Optional OverWriteFiles As Boolean = False) As Boolean
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  fso.CopyFile Source, Destination, OverWriteFiles
  Set fso = Nothing

  CopyFile = True

On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : CopyFolder
' Params :
'          ByVal Source As String
'          ByVal Destination As String
'          Optional OverWriteFiles As Boolean = False
' Returns: Boolean
' Desc   : The Function uses parameters ByVal Source As String, ByVal Destination As String and Optional OverWriteFiles As Boolean = False for CopyFolder and returns Boolean.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function CopyFolder(ByVal Source As String, ByVal Destination As String, Optional OverWriteFiles As Boolean = False) As Boolean
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  fso.CopyFolder Source, Destination, OverWriteFiles
  Set fso = Nothing
  
  CopyFolder = True

On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : CreateFolder
' Params :
'          ByVal Path As String
' Returns: Folder
' Desc   : The Function uses parameters ByVal Path As String for CreateFolder and returns Folder.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function CreateFolder(ByVal Path As String) As Folder
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  Set CreateFolder = fso.CreateFolder(Path)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : CreateTextFile
' Params :
'          ByVal FileName As String
'          Optional Overwrite As Boolean = False
'          Optional Unicode As Boolean = False
' Returns: TextStream
' Desc   : The Function uses parameters ByVal FileName As String, Optional Overwrite As Boolean = False and Optional Unicode As Boolean = False for CreateTextFile and returns TextStream.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function CreateTextFile(ByVal FileName As String, Optional Overwrite As Boolean = False, Optional Unicode As Boolean = False) As TextStream
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  Set CreateTextFile = fso.CreateTextFile(FileName, Overwrite, Unicode)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : DeleteFile
' Params :
'          ByVal FileSpec As String
'          Optional Force As Boolean = False
' Returns: Boolean
' Desc   : The Function uses parameters ByVal FileSpec As String and Optional Force As Boolean = False for DeleteFile and returns Boolean.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function DeleteFile(ByVal FileSpec As String, Optional Force As Boolean = False) As Boolean
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  fso.DeleteFile FileSpec, Force
  Set fso = Nothing
  
  DeleteFile = True

On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : DeleteFolder
' Params :
'          ByVal FolderSpec As String
'          Optional Force As Boolean = False
' Returns: Boolean
' Desc   : The Function uses parameters ByVal FolderSpec As String and Optional Force As Boolean = False for DeleteFolder and returns Boolean.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function DeleteFolder(ByVal FolderSpec As String, Optional Force As Boolean = False) As Boolean
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  fso.DeleteFolder FolderSpec, Force
  Set fso = Nothing

  DeleteFolder = True

On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : DriveExists
' Params :
'          ByVal DriveSpec As String
' Returns: Boolean
' Desc   : The Function uses parameters ByVal DriveSpec As String for DriveExists and returns Boolean.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function DriveExists(ByVal DriveSpec As String) As Boolean
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  DriveExists = fso.DriveExists(DriveSpec)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Property Get
' Name   : Drives
' Params :
' Returns: Drives
' Desc   : The Property Get uses parameters  for Drives and returns Drives.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Property Get Drives() As Drives
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  Set Drives = fso.Drives
  Set fso = Nothing
On_exit:
  Exit Property
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Property


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : FileExists
' Params :
'          ByVal FileSpec As String
' Returns: Boolean
' Desc   : The Function uses parameters ByVal FileSpec As String for FileExists and returns Boolean.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function FileExists(ByVal FileSpec As String) As Boolean
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  FileExists = fso.FileExists(FileSpec)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : FileFind
' Params :
'          ByVal FileName As String
'          ByVal Path As String
' Returns: String
' Desc   : The Function uses parameters ByVal FileName As String and ByVal Path As String for FileFind and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function FileFind(ByVal FileName As String, ByVal Path As String) As String
On Error GoTo On_Error
  Dim hFile As Long
  Dim ts As String
  Dim WFD As WIN32_FIND_DATA
  Dim result As Long
  Dim sAttempt As String
  Dim szPath As String
  Dim szPath2 As String
  Dim szFileName As String
  Dim dwBufferLen As Long
  Dim szBuffer As String
  Dim lpFilePart As String
  
  szPath = Me.BuildPath(Path, "*.*" & Chr$(0))
  szPath2 = Path & Chr$(0)
  szFileName = FileName & Chr$(0)
  szBuffer = String$(MAX_PATH, 0)
  dwBufferLen = Len(szBuffer)
  result = Searchpath(szPath2, szFileName, vbNullString, dwBufferLen, szBuffer, lpFilePart)

  If result Then
    FileFind = TrimNull(szBuffer)
    Exit Function
  End If
  
  hFile = FindFirstFile(szPath, WFD)
  
  Do
    If WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
      ts = TrimNull(WFD.cFileName)
      If Not (ts = "." Or ts = "..") Then
        If Not (WFD.dwFileAttributes And (FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM)) Then
          sAttempt = FileFind(FileName, Me.BuildPath(Path, ts))
          If sAttempt <> "" Then
            FileFind = sAttempt
            Exit Do
          End If
        End If
      End If
    End If
    WFD.cFileName = ""
    result = FindNextFile(hFile, WFD)
  Loop Until result = 0
  FindClose hFile

On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : FolderExists
' Params :
'          ByVal FolderSpec As String
' Returns: Boolean
' Desc   : The Function uses parameters ByVal FolderSpec As String for FolderExists and returns Boolean.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function FolderExists(ByVal FolderSpec As String) As Boolean
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  FolderExists = fso.FolderExists(FolderSpec)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : FileSearch
' Params :
'          ByVal FileName As String
' Returns: String
' Desc   : The Function uses parameters ByVal FileName As String for FileSearch and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function FileSearch(ByVal FileName As String) As String
On Error GoTo On_Error
  Dim drv As Drive
  Dim sPath As String
  
  For Each drv In Me.Drives
    sPath = Me.FileFind(FileName, drv.Path & "\")
    If sPath <> vbNullString Then Exit For
  Next drv
  
  FileSearch = sPath
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetAbsolutePathName
' Params :
'          ByVal Path As String
' Returns: String
' Desc   : The Function uses parameters ByVal Path As String for GetAbsolutePathName and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetAbsolutePathName(ByVal Path As String) As String
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  GetAbsolutePathName = fso.GetAbsolutePathName(Path)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetBaseName
' Params :
'          ByVal Path As String
' Returns: String
' Desc   : The Function uses parameters ByVal Path As String for GetBaseName and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetBaseName(ByVal Path As String) As String
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  GetBaseName = fso.GetBaseName(Path)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetDrive
' Params :
'          ByVal DriveSpec As String
' Returns: Drive
' Desc   : The Function uses parameters ByVal DriveSpec As String for GetDrive and returns Drive.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetDrive(ByVal DriveSpec As String) As Drive
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  Set GetDrive = fso.GetDrive(DriveSpec)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetDriveName
' Params :
'          ByVal Path As String
' Returns: String
' Desc   : The Function uses parameters ByVal Path As String for GetDriveName and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetDriveName(ByVal Path As String) As String
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  GetDriveName = fso.GetDriveName(Path)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetExtensionName
' Params :
'          ByVal Path As String
' Returns: String
' Desc   : The Function uses parameters ByVal Path As String for GetExtensionName and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetExtensionName(ByVal Path As String) As String
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  GetExtensionName = fso.GetExtensionName(Path)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetFile
' Params :
'          ByVal FilePath As String
' Returns: File
' Desc   : The Function uses parameters ByVal FilePath As String for GetFile and returns File.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetFile(ByVal FilePath As String) As File
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  Set GetFile = fso.GetFile(FilePath)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetFileName
' Params :
'          ByVal Path As String
' Returns: String
' Desc   : The Function uses parameters ByVal Path As String for GetFileName and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetFileName(ByVal Path As String) As String
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  GetFileName = fso.GetFileName(Path)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetFileVersion
' Params :
'          ByVal FileName As String
' Returns: String
' Desc   : The Function uses parameters ByVal FileName As String for GetFileVersion and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetFileVersion(ByVal FileName As String) As String
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  GetFileVersion = fso.GetFileVersion(FileName)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetFolder
' Params :
'          ByVal FolderPath As String
' Returns: Folder
' Desc   : The Function uses parameters ByVal FolderPath As String for GetFolder and returns Folder.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetFolder(ByVal FolderPath As String) As Folder
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  Set GetFolder = fso.GetFolder(FolderPath)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetParentFolderName
' Params :
'          ByVal Path As String
' Returns: String
' Desc   : The Function uses parameters ByVal Path As String for GetParentFolderName and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetParentFolderName(ByVal Path As String) As String
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  GetParentFolderName = fso.GetParentFolderName(Path)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetSpecialFolder
' Params :
'          ByVal SpecialFolder As SpecialFolderConst
' Returns: Folder
' Desc   : The Function uses parameters ByVal SpecialFolder As SpecialFolderConst for GetSpecialFolder and returns Folder.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetSpecialFolder(ByVal SpecialFolder As SpecialFolderConst) As Folder
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  Set GetSpecialFolder = fso.GetSpecialFolder(SpecialFolder)
  
  
  
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetStandardStream
' Params :
'          ByVal StandardStreamType As StandardStreamTypes
'          Optional Unicode As Boolean = False
' Returns: TextStream
' Desc   : The Function uses parameters ByVal StandardStreamType As StandardStreamTypes and Optional Unicode As Boolean = False for GetStandardStream and returns TextStream.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetStandardStream(ByVal StandardStreamType As StandardStreamTypes, Optional Unicode As Boolean = False) As TextStream
    On Error GoTo On_Error
    
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    Set GetStandardStream = fso.GetStandardStream(StandardStreamType, Unicode)
    Set fso = Nothing

On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : GetTempName
' Params :
' Returns: String
' Desc   : The Function uses parameters  for GetTempName and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function GetTempName() As String
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  GetTempName = fso.GetTempName
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : MoveFile
' Params :
'          ByVal Source As String
'          ByVal Destination As String
' Returns: Boolean
' Desc   : The Function uses parameters ByVal Source As String and ByVal Destination As String for MoveFile and returns Boolean.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function MoveFile(ByVal Source As String, ByVal Destination As String) As Boolean
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  fso.MoveFile Source, Destination
  Set fso = Nothing

  MoveFile = True

On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : MoveFolder
' Params :
'          ByVal Source As String
'          ByVal Destination As String
' Returns: Boolean
' Desc   : The Function uses parameters ByVal Source As String and ByVal Destination As String for MoveFolder and returns Boolean.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function MoveFolder(ByVal Source As String, ByVal Destination As String) As Boolean
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  fso.MoveFolder Source, Destination
  Set fso = Nothing

  MoveFolder = True

On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function


'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : OpenTextFile
' Params :
'          ByVal FileName As String
'          Optional IOMode As IOMode = ForReading
'          Optional Create As Boolean = False
'          Optional Format As Tristate = TristateFalse
' Returns: TextStream
' Desc   : The Function uses parameters ByVal FileName As String, Optional IOMode As IOMode = ForReading, Optional Create As Boolean = False and Optional Format As Tristate = TristateFalse for OpenTextFile and returns TextStream.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function OpenTextFile(ByVal FileName As String, Optional IOMode As IOMode = ForReading, Optional Create As Boolean = False, Optional Format As Tristate = TristateFalse) As TextStream
On Error GoTo On_Error
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  Set OpenTextFile = fso.OpenTextFile(FileName, IOMode, Create, Format)
  Set fso = Nothing
On_exit:
  Exit Function
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Function

'*****************************************************************************************
'Operating System Information
'****************************************************************************************

'__________________________________________________
' Scope  : Public
' Type   : Property Get
' Name   : OS_BuildNumber
' Params :
' Returns: Long
' Desc   : The Property Get uses parameters  for OS_BuildNumber and returns Long.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Property Get OS_BuildNumber() As Long
On Error GoTo On_Error
  Dim osi As OSVERSIONINFO
  Dim ret As Long

  osi.dwOSVersionInfoSize = 148
  osi.szCSDVersion = Space$(128)
  ret = GetVersionExA(osi)
  
  OS_BuildNumber = osi.dwBuildNumber
On_exit:
  Exit Property
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Property


'__________________________________________________
' Scope  : Public
' Type   : Property Get
' Name   : OS_MajorVersion
' Params :
' Returns: Long
' Desc   : The Property Get uses parameters  for OS_MajorVersion and returns Long.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Property Get OS_MajorVersion() As Long
On Error GoTo On_Error
  Dim osi As OSVERSIONINFO
  Dim ret As Long

  osi.dwOSVersionInfoSize = 148
  osi.szCSDVersion = Space$(128)
  ret = GetVersionExA(osi)
  
  OS_MajorVersion = osi.dwMajorVersion
On_exit:
  Exit Property
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Property


'__________________________________________________
' Scope  : Public
' Type   : Property Get
' Name   : OS_MinorVersion
' Params :
' Returns: Long
' Desc   : The Property Get uses parameters  for OS_MinorVersion and returns Long.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Property Get OS_MinorVersion() As Long
On Error GoTo On_Error
  Dim osi As OSVERSIONINFO
  Dim ret As Long

  osi.dwOSVersionInfoSize = 148
  osi.szCSDVersion = Space$(128)
  ret = GetVersionExA(osi)
  
  OS_MinorVersion = osi.dwMinorVersion
On_exit:
  Exit Property
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Property


'__________________________________________________
' Scope  : Public
' Type   : Property Get
' Name   : OS_PlatformId
' Params :
' Returns: Long
' Desc   : The Property Get uses parameters  for OS_PlatformId and returns Long.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Property Get OS_PlatformId() As Long
On Error GoTo On_Error
  Dim osi As OSVERSIONINFO
  Dim ret As Long

  osi.dwOSVersionInfoSize = 148
  osi.szCSDVersion = Space$(128)
  ret = GetVersionExA(osi)
  
  OS_PlatformId = osi.dwPlatformId
On_exit:
  Exit Property
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Property


'__________________________________________________
' Scope  : Public
' Type   : Property Get
' Name   : OS_MajorVersionName
' Params :
' Returns: String
' Desc   : The Property Get uses parameters  for OS_MajorVersionName and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Property Get OS_MajorVersionName() As String
On Error GoTo On_Error
  Select Case Me.OS_PlatformId
    Case 1 'Win 9x
      Select Case Me.OS_MinorVersion
        Case 0 '95
          OS_MajorVersionName = "Windows 95"
        Case 10 '98
          OS_MajorVersionName = "Windows 98"
      End Select
    Case 2 'Win NT
      Select Case Me.OS_MajorVersion
        Case 3 'NT 3
          OS_MajorVersionName = "Windows NT 3.51"
        Case 4 'NT 4
          OS_MajorVersionName = "Windows NT 4.0"
        Case 5 '2000
          OS_MajorVersionName = "Windows 2000"
      End Select
  End Select
On_exit:
  Exit Property
On_Error:
  MsgBox Err.Description, vbExclamation, "Program Error: " & Err.Number
  Resume On_exit
End Property



'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : TrimNull
' Params :
'          vVal As Variant
' Returns: Variant
' Desc   : The Function uses parameters vVal As Variant for TrimNull and returns Variant.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function TrimNull(vVal As Variant) As Variant
    On Error GoTo Proc_Err
    Const csProcName As String = "TrimNull"


    TrimNull = RTrim(LTrim(vVal))


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Function

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "clsFs->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Function
 
End Function


