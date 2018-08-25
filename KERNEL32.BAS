Attribute VB_Name = "basKernel32"
' Module Name               basKernel32
' Module Description        Module which contains ALL Declarations
'                           for API calls in the Kernel32 Library
'
Option Explicit

Declare Function AddAtom _
    Lib "kernel32" _
    Alias "AddAtomA" ( _
        ByVal lpString As String) _
    As Integer

Declare Function AllocConsole _
    Lib "kernel32" ( _
        ) _
    As Long

Declare Function BackupRead _
    Lib "kernel32" ( _
        ByVal hFile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, _
        lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, _
        lpContext As Any) _
    As Long

Declare Function BackupSeek _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, _
        lpdwLowByteSeeked As Long, lpdwHighByteSeeked As Long, _
        lpContext As Long) _
    As Long

Declare Function BackupWrite _
    Lib "kernel32" ( _
        ByVal hFile As Long, lpBuffer As Byte, _
        ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, _
        ByVal bAbort As Long, ByVal bProcessSecurity As Long, _
        lpContext As Long) _
    As Long

Declare Function Beep _
    Lib "kernel32" ( _
        ByVal dwFreq As Long, _
        ByVal dwDuration As Long) _
    As Long

Declare Function BeginUpdateResource _
    Lib "kernel32" _
    Alias "BeginUpdateResourceA" ( _
        ByVal pFileName As String, _
        ByVal bDeleteExistingResources As Long) _
    As Long

Declare Function BuildCommDCB _
    Lib "kernel32" _
    Alias "BuildCommDCBA" ( _
        ByVal lpDef As String, _
        lpDCB As DCB) _
    As Long

Declare Function BuildCommDCBAndTimeouts _
    Lib "kernel32" _
    Alias "BuildCommDCBAndTimeoutsA" ( _
        ByVal lpDef As String, _
        lpDCB As DCB, _
        lpCommTimeouts As COMMTIMEOUTS) _
    As Long

Declare Function ClearCommBreak _
    Lib "kernel32" ( _
        ByVal nCid As Long) _
    As Long

Declare Function ClearCommError _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpErrors As Long, _
        lpStat As COMSTAT) _
    As Long

Declare Function CloseHandle _
    Lib "kernel32" ( _
        ByVal hObject As Long) _
    As Long

Declare Function CommConfigDialog _
    Lib "kernel32" _
    Alias "CommConfigDialogA" ( _
        ByVal lpszName As String, _
        ByVal hWnd As Long, _
        lpCC As COMMCONFIG) _
    As Boolean

Declare Function CompareFileTime _
    Lib "kernel32" ( _
        lpFileTime1 As FILETIME, _
        lpFileTime2 As FILETIME) _
    As Long

Declare Function CompareString _
    Lib "kernel32" _
    Alias "CompareStringA" ( _
        ByVal Locale As Long, ByVal dwCmpFlags As Long, _
        ByVal lpString1 As String, ByVal cchCount1 As Long, _
        ByVal lpString2 As String, ByVal cchCount2 As Long) _
    As Long

Declare Function ConnectNamedPipe _
    Lib "kernel32" ( _
        ByVal hNamedPipe As Long, _
        lpOverlapped As OVERLAPPED) _
    As Long

Declare Function ContinueDebugEvent _
    Lib "kernel32" ( _
        ByVal dwProcessId As Long, _
        ByVal dwThreadId As Long, _
        ByVal dwContinueStatus As Long) _
    As Long

Declare Function ConvertDefaultLocale _
    Lib "kernel32" ( _
        ByVal Locale As Long) _
    As Long

Declare Function CopyFile _
    Lib "kernel32" _
    Alias "CopyFileA" ( _
        ByVal lpExistingFileName As String, _
        ByVal lpNewFileName As String, _
        ByVal bFailIfExists As Long) _
    As Long

Declare Function CreateDirectory _
    Lib "kernel32" _
    Alias "CreateDirectoryA" ( _
        ByVal lpPathName As String, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES) _
    As Long

Declare Function CreateDirectoryEx _
    Lib "kernel32" _
    Alias "CreateDirectoryExA" ( _
        ByVal lpTemplateDirectory As String, _
        ByVal lpNewDirectory As String, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES) _
    As Long

Declare Function CreateEvent _
    Lib "kernel32" _
    Alias "CreateEventA" ( _
        lpEventAttributes As SECURITY_ATTRIBUTES, _
        ByVal bManualReset As Long, ByVal bInitialState As Long, _
        ByVal lpName As String) _
    As Long

Declare Function CreateFile _
    Lib "kernel32" _
    Alias "CreateFileA" ( _
        ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) _
    As Long

Declare Function CreateFileMapping _
    Lib "kernel32" _
    Alias "CreateFileMappingA" ( _
        ByVal hFile As Long, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, _
        ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) _
    As Long

Declare Function CreateIoCompletionPort _
    Lib "kernel32" ( _
        ByVal FileHandle As Long, _
        ByVal ExistingCompletionPort As Long, _
        ByVal CompletionKey As Long, _
        ByVal NumberOfConcurrentThreads As Long) _
    As Long

Declare Function CreateMailslot _
    Lib "kernel32" _
    Alias "CreateMailslotA" ( _
        ByVal lpName As String, ByVal nMaxMessageSize As Long, _
        ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) _
    As Long

Declare Function CreateMutex _
    Lib "kernel32" _
    Alias "CreateMutexA" ( _
        lpMutexAttributes As SECURITY_ATTRIBUTES, _
        ByVal bInitialOwner As Long, _
        ByVal lpName As String) _
    As Long

Declare Function CreateNamedPipe _
    Lib "kernel32" _
    Alias "CreateNamedPipeA" ( _
        ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, _
        ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, _
        ByVal nDefaultTimeOut As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) _
    As Long

Declare Function CreatePipe _
    Lib "kernel32" ( _
        phReadPipe As Long, _
        phWritePipe As Long, _
        lpPipeAttributes As SECURITY_ATTRIBUTES, _
        ByVal nSize As Long) _
    As Long

Declare Function CreateProcess _
    Lib "kernel32" _
    Alias "CreateProcessA" ( _
        ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, _
        lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
        lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) _
    As Long

Declare Function CreateProcessAsUser _
    Lib "kernel32" _
    Alias "CreateProcessAsUserA" ( _
        ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As SECURITY_ATTRIBUTES, _
        ByVal lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, ByVal lpStartupInfo As STARTUPINFO, ByVal lpProcessInformation As PROCESS_INFORMATION) _
    As Long

Declare Function CreateRemoteThread _
    Lib "kernel32" ( _
        ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, _
        lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) _
    As Long

Declare Function CreateSemaphore _
    Lib "kernel32" _
    Alias "CreateSemaphoreA" ( _
        lpSemaphoreAttributes As SECURITY_ATTRIBUTES, _
        ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) _
    As Long

Declare Function CreateTapePartition _
    Lib "kernel32" ( _
        ByVal hDevice As Long, _
        ByVal dwPartitionMethod As Long, _
        ByVal dwCount As Long, _
        ByVal dwSize As Long) _
    As Long

Declare Function CreateThread _
    Lib "kernel32" ( _
        lpThreadAttributes As SECURITY_ATTRIBUTES, _
        ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, _
        ByVal dwCreationFlags As Long, lpThreadId As Long) _
    As Long

Declare Function DefineDosDevice _
    Lib "kernel32" _
    Alias "DefineDosDeviceA" ( _
        ByVal dwFlags As Long, _
        ByVal lpDeviceName As String, _
        ByVal lpTargetPath As String) _
    As Long

Declare Function DeleteAtom _
    Lib "kernel32" ( _
        ByVal nAtom As Integer) _
    As Integer

Declare Sub DeleteCriticalSection _
    Lib "kernel32" ( _
        lpCriticalSection As CRITICAL_SECTION)

Declare Function DeleteFile _
    Lib "kernel32" _
    Alias "DeleteFileA" ( _
        ByVal lpFileName As String) _
    As Long

Declare Function DeviceIoControl _
    Lib "kernel32" ( _
        ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, _
        ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, _
        lpBytesReturned As Long, lpOverlapped As OVERLAPPED) _
    As Long

Declare Function DisableThreadLibraryCalls _
    Lib "kernel32" ( _
        ByVal hLibModule As Long) _
    As Boolean

Declare Function DisconnectNamedPipe _
    Lib "kernel32" ( _
        ByVal hNamedPipe As Long) _
    As Long

Declare Function DosDateTimeToFileTime _
    Lib "kernel32" ( _
        ByVal wFatDate As Long, _
        ByVal wFatTime As Long, _
        lpFileTime As FILETIME) _
    As Long

Declare Function DuplicateHandle _
    Lib "kernel32" ( _
        ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, _
        ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, _
        ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
        ByVal dwOptions As Long) _
    As Long

Declare Function EndUpdateResource _
    Lib "kernel32" _
    Alias "EndUpdateResourceA" ( _
        ByVal hUpdate As Long, _
        ByVal fDiscard As Long) _
    As Long

Declare Sub EnterCriticalSection _
    Lib "kernel32" ( _
        lpCriticalSection As CRITICAL_SECTION)

Declare Function EnumCalendarInfo _
    Lib "kernel32" _
    Alias "EnumCalendarInfoA" ( _
        ByVal lpCalInfoEnumProc As Long, ByVal Locale As Long, _
        ByVal Calendar As Long, ByVal CalType As Long) _
    As Boolean

Declare Function EnumDateFormats _
    Lib "kernel32" ( _
        ByVal lpDateFmtEnumProc As Long, _
        ByVal Locale As Long, _
        ByVal dwFlags As Long) _
    As Long

Declare Function EnumResourceLanguages _
    Lib "kernel32" _
    Alias "EnumResourceLanguagesA" ( _
        ByVal hModule As Long, ByVal lpType As String, _
        ByVal lpName As String, ByVal lpEnumFunc As Long, _
        ByVal lParam As Long) _
    As Boolean

Declare Function EnumResourceNames _
    Lib "kernel32" _
    Alias "EnumResourceNamesA" ( _
        ByVal hModule As Long, ByVal lpType As String, _
        ByVal lpEnumFunc As Long, ByVal lParam As Long) _
    As Boolean

Declare Function EnumResourceTypes _
    Lib "kernel32" _
    Alias "EnumResourceTypesA" ( _
        ByVal hModule As Long, _
        ByVal lpEnumFunc As Long, _
        ByVal lParam As Long) _
    As Boolean

Declare Function EnumSystemCodePages _
    Lib "kernel32" ( _
        ByVal lpCodePageEnumProc As Long, _
        ByVal dwFlags As Long) _
    As Long

Declare Function EnumSystemLocales _
    Lib "kernel32" ( _
        ByVal lpLocaleEnumProc As Long, _
        ByVal dwFlags As Long) _
    As Long

Declare Function EnumTimeFormats _
    Lib "kernel32" ( _
        ByVal lpTimeFmtEnumProc As Long, _
        ByVal Locale As Long, _
        ByVal dwFlags As Long) _
    As Long

Declare Function EraseTape _
    Lib "kernel32" ( _
        ByVal hDevice As Long, _
        ByVal dwEraseType As Long, _
        ByVal bimmediate As Long) _
    As Long

Declare Function EscapeCommFunction _
    Lib "kernel32" ( _
        ByVal nCid As Long, _
        ByVal nFunc As Long) _
    As Long

Declare Sub ExitProcess _
    Lib "kernel32" ( _
        ByVal uExitCode As Long)

Declare Sub ExitThread _
    Lib "kernel32" ( _
        ByVal dwExitCode As Long)

Declare Function ExpandEnvironmentStrings _
    Lib "kernel32" _
    Alias "ExpandEnvironmentStringsA" ( _
        ByVal lpSrc As String, _
        ByVal lpDst As String, _
        ByVal nSize As Long) _
    As Long

Declare Sub FatalAppExit _
    Lib "kernel32" _
    Alias "FatalAppExitA" ( _
        ByVal uAction As Long, _
        ByVal lpMessageText As String)

Declare Sub FatalExit _
    Lib "kernel32" ( _
        ByVal code As Long)

Declare Function FileTimeToDosDateTime _
    Lib "kernel32" ( _
        lpFileTime As FILETIME, _
        ByVal lpFatDate As Long, _
        ByVal lpFatTime As Long) _
    As Long

Declare Function FileTimeToLocalFileTime _
    Lib "kernel32" ( _
        lpFileTime As FILETIME, _
        lpLocalFileTime As FILETIME) _
    As Long

Declare Function FileTimeToSystemTime _
    Lib "kernel32" ( _
        lpFileTime As FILETIME, _
        lpSystemTime As SYSTEMTIME) _
    As Long

Declare Function FillConsoleOutputAttribute _
    Lib "kernel32" ( _
        ByVal hConsoleOutput As Long, _
        ByVal wAttribute As Long, ByVal nLength As Long, _
        dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) _
    As Long

Declare Function FillConsoleOutputCharacter _
    Lib "kernel32" _
    Alias "FillConsoleOutputCharacterA" ( _
        ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, _
        ByVal nLength As Long, dwWriteCoord As COORD, _
        lpNumberOfCharsWritten As Long) _
    As Long

Declare Function FindAtom _
    Lib "kernel32" _
    Alias "FindAtomA" ( _
        ByVal lpString As String) _
    As Integer

Declare Function FindClose _
    Lib "kernel32" ( _
        ByVal hFindFile As Long) _
    As Long

Declare Function FindCloseChangeNotification _
    Lib "kernel32" ( _
        ByVal hChangeHandle As Long) _
    As Long

Declare Function FindFirstChangeNotification _
    Lib "kernel32" _
    Alias "FindFirstChangeNotificationA" ( _
        ByVal lpPathName As String, _
        ByVal bWatchSubtree As Long, _
        ByVal dwNotifyFilter As Long) _
    As Long

Declare Function FindFirstFile _
    Lib "kernel32" _
    Alias "FindFirstFileA" ( _
        ByVal lpFileName As String, _
        lpFindFileData As WIN32_FIND_DATA) _
    As Long

Declare Function FindNextChangeNotification _
    Lib "kernel32" ( _
        ByVal hChangeHandle As Long) _
    As Long

Declare Function FindNextFile _
    Lib "kernel32" _
    Alias "FindNextFileA" ( _
        ByVal hFindFile As Long, _
        lpFindFileData As WIN32_FIND_DATA) _
    As Long

Declare Function FindResource _
    Lib "kernel32" _
    Alias "FindResourceA" ( _
        ByVal hInstance As Long, _
        ByVal lpName As String, _
        ByVal lpType As String) _
    As Long

Declare Function FindResourceEx _
    Lib "kernel32" _
    Alias "FindResourceExA" ( _
        ByVal hModule As Long, ByVal lpType As String, _
        ByVal lpName As String, ByVal wLanguage As Long) _
    As Long

Declare Function FreeConsole _
    Lib "kernel32" () _
    As Long

Declare Function FreeEnvironmentStrings _
    Lib "kernel32" _
    Alias "FreeEnvironmentStringsA" ( _
        ByVal lpsz As String) _
    As Boolean

Declare Function FreeLibrary _
    Lib "kernel32" ( _
        ByVal hLibModule As Long) _
    As Long

Declare Sub FreeLibraryAndExitThread _
    Lib "kernel32" ( _
        ByVal hLibModule As Long, _
        ByVal dwExitCode As Long)

Declare Function FreeResource _
    Lib "kernel32" ( _
        ByVal hResData As Long) _
    As Boolean

Declare Function GenerateConsoleCtrlEvent _
    Lib "kernel32" ( _
        ByVal dwCtrlEvent As Long, _
        ByVal dwProcessGroupId As Long) _
    As Long

Declare Function GetACP _
    Lib "kernel32" () _
    As Long

Declare Function GetAtomName _
    Lib "kernel32" _
    Alias "GetAtomNameA" ( _
        ByVal nAtom As Integer, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long) _
    As Long

Declare Function GetBinaryType _
    Lib "kernel32" _
    Alias "GetBinaryTypeA" ( _
        ByVal lpApplicationName As String, _
        lpBinaryType As Long) _
    As Long

Declare Function GetCommandLine _
    Lib "kernel32" _
    Alias "GetCommandLineA" () _
    As String

Declare Function GetCommConfig _
    Lib "kernel32" ( _
        ByVal hCommDev As Long, _
        lpCC As COMMCONFIG, _
        lpdwSize As Long) _
    As Boolean

Declare Function GetCommMask _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpEvtMask As Long) _
    As Long

Declare Function GetCommModemStatus _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpModemStat As Long) _
    As Long

Declare Function GetCommProperties _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpCommProp As COMMPROP) _
    As Long

Declare Function GetCommState _
    Lib "kernel32" ( _
        ByVal nCid As Long, _
        lpDCB As DCB) _
    As Long

Declare Function GetCommTimeouts _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpCommTimeouts As COMMTIMEOUTS) _
    As Long

Declare Function GetCompressedFileSize _
    Lib "kernel32" _
    Alias "GetCompressedFileSizeA" ( _
        ByVal lpFileName As String, _
        lpFileSizeHigh As Long) _
    As Long

Declare Function GetComputerName _
    Lib "kernel32" _
    Alias "GetComputerNameA" ( _
        ByVal lpBuffer As String, _
        nSize As Long) _
    As Long

Declare Function GetConsoleCP _
    Lib "kernel32" () _
    As Long

Declare Function GetConsoleCursorInfo _
    Lib "kernel32" ( _
        ByVal hConsoleOutput As Long, _
        lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) _
    As Long

Declare Function GetConsoleMode _
    Lib "kernel32" ( _
        ByVal hConsoleHandle As Long, _
        lpMode As Long) _
    As Long

Declare Function GetConsoleOutputCP _
    Lib "kernel32" () _
    As Long

Declare Function GetConsoleScreenBufferInfo _
    Lib "kernel32" ( _
        ByVal hConsoleOutput As Long, _
        lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) _
    As Long

Declare Function GetDateFormat _
    Lib "kernel32" _
    Alias "GetDateFormatA" ( _
        ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, _
        ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) _
    As Long

Declare Function GetDefaultCommConfig _
    Lib "kernel32" _
    Alias "GetDefaultCommConfigA" ( _
        ByVal lpszName As String, lpCC As COMMCONFIG, _
        lpdwSize As Long) _
    As Boolean

Declare Function GetDiskFreeSpace _
    Lib "kernel32" _
    Alias "GetDiskFreeSpaceA" ( _
        ByVal lpRootPathName As String, lpSectorsPerCluster As Long, _
        lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, _
        lpTtoalNumberOfClusters As Long) _
    As Long

Declare Function GetDriveType _
    Lib "kernel32" _
    Alias "GetDriveTypeA" ( _
        ByVal nDrive As String) _
    As Long

Declare Function GetEnvironmentStrings _
    Lib "kernel32" _
    Alias "GetEnvironmentStringsA" ( _
        ) _
    As String

Declare Function GetEnvironmentVariable _
    Lib "kernel32" _
    Alias "GetEnvironmentVariableA" ( _
        ByVal lpName As String, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long) _
    As Long

Declare Function GetExitCodeProcess _
    Lib "kernel32" ( _
        ByVal hProcess As Long, _
        lpExitCode As Long) _
    As Long

Declare Function GetExitCodeThread _
    Lib "kernel32" ( _
        ByVal hThread As Long, _
        lpExitCode As Long) _
    As Long

Declare Function GetFileAttributes _
    Lib "kernel32" _
    Alias "GetFileAttributesA" ( _
        ByVal lpFileName As String) _
    As Long

Declare Function GetFileInformationByHandle _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpFileInformation As BY_HANDLE_FILE_INFORMATION) _
    As Long

Declare Function GetFileSize _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpFileSizeHigh As Long) _
    As Long

Declare Function GetFileTime _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpCreationTime As FILETIME, _
        lpLastAccessTime As FILETIME, _
        lpLastWriteTime As FILETIME) _
    As Long

Declare Function GetFileType _
    Lib "kernel32" ( _
        ByVal hFile As Long) _
    As Long

Declare Function GetFullPathName _
    Lib "kernel32" _
    Alias "GetFullPathNameA" ( _
        ByVal lpFileName As String, ByVal nBufferLength As Long, _
        ByVal lpBuffer As String, ByVal lpFilePart As String) _
    As Long

Declare Function GetHandleInformation _
    Lib "kernel32" ( _
        ByVal hObject As Long, _
        lpdwFlags As Long) _
    As Boolean

Declare Function GetLargestConsoleWindowSize _
    Lib "kernel32" ( _
        ByVal hConsoleOutput As Long) _
    As COORD

Declare Function GetLastError _
    Lib "kernel32" ( _
        ) _
    As Long

Declare Function GetLocaleInfo _
    Lib "kernel32" _
    Alias "GetLocaleInfoA" ( _
        ByVal Locale As Long, ByVal LCType As Long, _
        ByVal lpLCData As String, ByVal cchData As Long) _
    As Long

Declare Sub GetLocalTime _
    Lib "kernel32" ( _
        lpSystemTime As SYSTEMTIME)

Declare Function GetLogicalDrives _
    Lib "kernel32" ( _
        ) _
    As Long

Declare Function GetLogicalDriveStrings _
    Lib "kernel32" _
    Alias "GetLogicalDriveStringsA" ( _
        ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) _
    As Long

Declare Function GetMailslotInfo _
    Lib "kernel32" ( _
        ByVal hMailslot As Long, lpMaxMessageSize As Long, _
        lpNextSize As Long, lpMessageCount As Long, _
        lpReadTimeout As Long) _
    As Long

Declare Function GetModuleFileName _
    Lib "kernel32" _
    Alias "GetModuleFileNameA" ( _
        ByVal hModule As Long, _
        ByVal lpFileName As String, _
        ByVal nSize As Long) _
    As Long

Declare Function GetModuleHandle _
    Lib "kernel32" _
    Alias "GetModuleHandleA" ( _
        ByVal lpModuleName As String) _
    As Long

Declare Function GetOEMCP _
    Lib "kernel32" ( _
        ) _
    As Long

Declare Function GetOverlappedResult _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpOverlapped As OVERLAPPED, _
        lpNumberOfBytesTransferred As Long, _
        ByVal bWait As Long) _
    As Long

Declare Function GetPriorityClass _
    Lib "kernel32" ( _
        ByVal hProcess As Long) _
    As Long

Declare Function GetPrivateProfileInt _
    Lib "kernel32" _
    Alias "GetPrivateProfileIntA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As String, _
        ByVal nDefault As Long, ByVal lpFileName As String) _
    As Long

Declare Function GetPrivateProfileSection _
    Lib "kernel32" _
    Alias "GetPrivateProfileSectionA" ( _
        ByVal lpAppName As String, ByVal lpReturnedString As String, _
        ByVal nSize As Long, ByVal lpFileName As String) _
    As Long

Declare Function GetPrivateProfileString _
    Lib "kernel32" _
    Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
        ByVal lpDefault As String, ByVal lpReturnedString As String, _
        ByVal nSize As Long, ByVal lpFileName As String) _
    As Long

Declare Function GetProcAddress _
    Lib "kernel32" ( _
        ByVal hModule As Long, _
        ByVal lpProcName As String) _
    As Long

Declare Function GetProcessAffinityMask _
    Lib "kernel32" ( _
        ByVal hProcess As Long, _
        lpProcessAffinityMask As Long, _
        SystemAffinityMask As Long) _
    As Long

Declare Function GetProcessHeap _
    Lib "kernel32" () _
    As Long

Declare Function GetProcessHeaps _
    Lib "kernel32" ( _
        ByVal NumberOfHeaps As Long, _
        ProcessHeaps As Long) _
    As Long

Declare Function GetProcessShutdownParameters _
    Lib "kernel32" ( _
        lpdwLevel As Long, _
        lpdwFlags As Long) _
    As Long

Declare Function GetProcessTimes _
    Lib "kernel32" ( _
        ByVal hProcess As Long, _
        lpCreationTime As FILETIME, lpExitTime As FILETIME, _
        lpKernelTime As FILETIME, lpUserTime As FILETIME) _
    As Long

Declare Function GetProcessWorkingSetSize _
    Lib "kernel32" ( _
        ByVal hProcess As Long, _
        lpMinimumWorkingSetSize As Long, _
        lpMaximumWorkingSetSize As Long) _
    As Boolean

Declare Function GetProfileInt _
    Lib "kernel32" _
    Alias "GetProfileIntA" ( _
        ByVal lpAppName As String, _
        ByVal lpKeyName As String, _
        ByVal nDefault As Long) _
    As Long

Declare Function GetProfileSection _
    Lib "kernel32" _
    Alias "GetProfileSectionA" ( _
        ByVal lpAppName As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long) _
    As Long

Declare Function GetProfileString _
    Lib "kernel32" _
    Alias "GetProfileStringA" ( _
        ByVal lpAppName As String, ByVal lpKeyName As String, _
        ByVal lpDefault As String, ByVal lpReturnedString As String, _
        ByVal nSize As Long) _
    As Long

Declare Function GetQueuedCompletionStatus _
    Lib "kernel32" ( _
        ByVal CompletionPort As Long, lpNumberOfBytesTransferred As Long, _
        lpCompletionKey As Long, lpOverlapped As Long, _
        ByVal dwMilliseconds As Long) _
    As Boolean

Declare Function GetShortPathName _
    Lib "kernel32" ( _
        ByVal lpszLongPath As String, _
        ByVal lpszShortPath As String, _
        ByVal cchBuffer As Long) _
    As Long

Declare Sub GetStartupInfo _
    Lib "kernel32" _
    Alias "GetStartupInfoA" ( _
        lpStartupInfo As STARTUPINFO)

Declare Function GetStdHandle _
    Lib "kernel32" ( _
        ByVal nStdHandle As Long) _
    As Long

Declare Function GetStringTypeA _
    Lib "kernel32" ( _
        ByVal lcid As Long, ByVal dwInfoType As Long, _
        ByVal lpSrcStr As String, ByVal cchSrc As Long, _
        lpCharType As Long) _
    As Long

Declare Function GetStringTypeEx _
    Lib "kernel32" _
    Alias "GetStringTypeExA" ( _
        ByVal Locale As Long, ByVal dwInfoType As Long, _
        ByVal lpSrcStr As String, ByVal cchSrc As Long, _
        lpCharType As Integer) _
    As Boolean

Declare Function GetStringTypeW _
    Lib "kernel32" ( _
        ByVal dwInfoType As Long, _
        ByVal lpSrcStr As String, _
        ByVal cchSrc As Long, _
        lpCharType As Integer) _
    As Boolean

Declare Function GetSystemPowerStatus _
    Lib "kernel32" ( _
        lpSystemPowerStatus As SYSTEM_POWER_STATUS) _
    As Long

Declare Sub GetSystemTime _
    Lib "kernel32" ( _
        lpSystemTime As SYSTEMTIME)

Declare Function GetSystemTimeAdjustment _
    Lib "kernel32" ( _
        lpTimeAdjustment As Long, _
        lpTimeIncrement As Long, _
        lpTimeAdjustmentDisabled As Boolean) _
    As Boolean

Declare Function GetTapeParameters _
    Lib "kernel32" ( _
        ByVal hDevice As Long, _
        ByVal dwOperation As Long, _
        lpdwSize As Long, _
        lpTapeInformation As Any) _
    As Long

Declare Function GetTapePosition _
    Lib "kernel32" ( _
        ByVal hDevice As Long, ByVal dwPositionType As Long, _
        lpdwPartition As Long, lpdwOffsetLow As Long, _
        lpdwOffsetHigh As Long) _
    As Long

Declare Function GetTapeStatus _
    Lib "kernel32" ( _
        ByVal hDevice As Long) _
    As Long

Declare Function GetTempFileName _
    Lib "kernel32" _
    Alias "GetTempFileNameA" ( _
        ByVal lpszPath As String, ByVal lpPrefixString As String, _
        ByVal wUnique As Long, ByVal lpTempFileName As String) _
    As Long

Declare Function GetTempPath _
    Lib "kernel32" _
    Alias "GetTempPathA" ( _
        ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) _
    As Long

Declare Function GetThreadContext _
    Lib "kernel32" ( _
        ByVal hThread As Long, _
        lpContext As CONTEXT) _
    As Long

Declare Function GetThreadLocale _
    Lib "kernel32" ( _
        ) _
    As Long

Declare Function GetThreadPriority _
    Lib "kernel32" ( _
        ByVal hThread As Long) _
    As Long

Declare Function GetThreadSelectorEntry _
    Lib "kernel32" ( _
        ByVal hThread As Long, _
        ByVal dwSelector As Long, _
        lpSelectorEntry As LDT_ENTRY) _
    As Long

Declare Function GetThreadTimes _
    Lib "kernel32" ( _
        ByVal hThread As Long, _
        lpCreationTime As FILETIME, lpExitTime As FILETIME, _
        lpKernelTime As FILETIME, lpUserTime As FILETIME) _
    As Long

Declare Function GetTickCount _
    Lib "kernel32" () _
    As Long

Declare Function GetTimeFormat _
    Lib "kernel32" _
    Alias "GetTimeFormatA" ( _
        ByVal Locale As Long, ByVal dwFlags As Long, _
        lpTime As SYSTEMTIME, ByVal lpFormat As String, _
        ByVal lpTimeStr As String, ByVal cchTime As Long) _
    As Long

Declare Function GetTimeZoneInformation _
    Lib "kernel32" ( _
        lpTimeZoneInformation As TIME_ZONE_INFORMATION) _
    As Long

Declare Function GetUserDefaultLangID _
    Lib "kernel32" () _
    As Integer

Declare Function GetUserDefaultLCID _
    Lib "kernel32" () _
    As Long

Declare Function GetVersion _
    Lib "kernel32" () _
    As Long

Declare Function GetVersionEx _
    Lib "kernel32" _
    Alias "GetVersionExA" ( _
        ByVal lpVersionInformation As OSVERSIONINFO) _
    As Long

Declare Function GetVolumeInformation _
    Lib "kernel32" _
    Alias "GetVolumeInformationA" ( _
        ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
        lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
        ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) _
    As Long

Declare Function GetWindowsDirectory _
    Lib "kernel32" _
    Alias "GetWindowsDirectoryA" ( _
        ByVal lpBuffer As String, _
        ByVal nSize As Long) _
    As Long

Declare Function GlobalAddAtom _
    Lib "kernel32" _
    Alias "GlobalAddAtomA" ( _
        ByVal lpString As String) _
    As Integer

Declare Function GlobalAlloc _
    Lib "kernel32" ( _
        ByVal wFlags As Long, _
        ByVal dwBytes As Long) _
    As Long

Declare Function GlobalCompact _
    Lib "kernel32" ( _
        ByVal dwMinFree As Long) _
    As Long

Declare Function GlobalDeleteAtom _
    Lib "kernel32" ( _
        ByVal nAtom As Integer) _
    As Integer

Declare Function GlobalFindAtom _
    Lib "kernel32" _
    Alias "GlobalFindAtomA" ( _
        ByVal lpString As String) _
    As Integer

Declare Sub GlobalFix _
    Lib "kernel32" ( _
        ByVal hMem As Long)

Declare Function GlobalFlags _
    Lib "kernel32" ( _
        ByVal hMem As Long) _
    As Long

Declare Function GlobalFree _
    Lib "kernel32" ( _
        ByVal hMem As Long) _
    As Long

Declare Function GlobalGetAtomName _
    Lib "kernel32" _
    Alias "GlobalGetAtomNameA" ( _
        ByVal nAtom As Integer, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long) _
    As Long

Declare Function HeapAlloc _
    Lib "kernel32" ( _
        ByVal hHeap As Long, _
        ByVal dwFlags As Long, _
        ByVal dwBytes As Long) _
    As Long

Declare Function HeapCompact _
    Lib "kernel32" ( _
        ByVal hHeap As Long, _
        ByVal dwFlags As Long) _
    As Long

Declare Function HeapCreate _
    Lib "kernel32" ( _
        ByVal flOptions As Long, _
        ByVal dwInitialSize As Long, _
        ByVal dwMaximumSize As Long) _
    As Long

Declare Function HeapDestroy _
    Lib "kernel32" ( _
        ByVal hHeap As Long) _
    As Long

Declare Function HeapFree _
    Lib "kernel32" ( _
        ByVal hHeap As Long, _
        ByVal dwFlags As Long, _
        lpMem As Any) _
    As Long

Declare Function HeapLock _
    Lib "kernel32" ( _
        ByVal hHeap As Long) _
    As Long

Declare Function HeapReAlloc _
    Lib "kernel32" ( _
        ByVal hHeap As Long, _
        ByVal dwFlags As Long, _
        lpMem As Any, _
        ByVal dwBytes As Long) _
    As Long

Declare Function HeapSize _
    Lib "kernel32" ( _
        ByVal hHeap As Long, _
        ByVal dwFlags As Long, _
        lpMem As Any) _
    As Long

Declare Function HeapUnlock _
    Lib "kernel32" ( _
        ByVal hHeap As Long) _
    As Long

Declare Function HeapValidate _
    Lib "kernel32" ( _
        ByVal hHeap As Long, _
        ByVal dwFlags As Long, _
        lpMem As Any) _
    As Long

Declare Function hread _
    Lib "kernel32" _
    Alias "_hread" ( _
        ByVal hFile As Long, _
        lpBuffer As Any, _
        ByVal lBytes As Long) _
    As Long

Declare Function hwrite _
    Lib "kernel32" _
    Alias "_hwrite" ( _
        ByVal hFile As Long, _
        ByVal lpBuffer As String, _
        ByVal lBytes As Long) _
    As Long

Declare Function ImpersonateLoggedOnUser _
    Lib "kernel32" ( _
        ByVal hToken As Long) _
    As Long

Declare Function InitAtomTable _
    Lib "kernel32" ( _
        ByVal nSize As Long) _
    As Long

Declare Sub InitializeCriticalSection _
    Lib "kernel32" ( _
        lpCriticalSection As CRITICAL_SECTION)

Declare Function InterlockedDecrement _
    Lib "kernel32" ( _
        lpAddend As Long) _
    As Long

Declare Function InterlockedExchange _
    Lib "kernel32" ( _
        Target As Long, _
        ByVal Value As Long) _
    As Long

Declare Function InterlockedIncrement _
    Lib "kernel32" ( _
        lpAddend As Long) _
    As Long

Declare Function IsBadCodePtr _
    Lib "kernel32" ( _
        ByVal lpfn As Long) _
    As Boolean

Declare Function IsBadHugeReadPtr _
    Lib "kernel32" ( _
        lp As Any, _
        ByVal ucb As Long) _
    As Long

Declare Function IsBadHugeWritePtr _
    Lib "kernel32" ( _
        lp As Any, _
        ByVal ucb As Long) _
    As Long

Declare Function IsBadReadPtr _
    Lib "kernel32" ( _
        lp As Any, _
        ByVal ucb As Long) _
    As Long

Declare Function IsBadStringPtr _
    Lib "kernel32" _
    Alias "IsBadStringPtrA" ( _
        ByVal lpsz As String, _
        ByVal ucchMax As Long) _
    As Long

Declare Function IsBadWritePtr _
    Lib "kernel32" ( _
        lp As Any, _
        ByVal ucb As Long) _
    As Long

Declare Function IsDBCSLeadByte _
    Lib "kernel32" ( _
        ByVal bTestChar As Byte) _
    As Long

Declare Function lclose _
    Lib "kernel32" _
    Alias "_lclose" ( _
        ByVal hFile As Long) _
    As Long

Declare Function LCMapString _
    Lib "kernel32" _
    Alias "LCMapStringA" ( _
        ByVal Locale As Long, ByVal dwMapFlags As Long, _
        ByVal lpSrcStr As String, ByVal cchSrc As Long, _
        ByVal lpDestStr As String, ByVal cchDest As Long) _
    As Long

Declare Function lcreat _
    Lib "kernel32" _
    Alias "_lcreat" ( _
        ByVal lpPathName As String, _
        ByVal iAttribute As Long) _
    As Long

Declare Sub LeaveCriticalSection _
    Lib "kernel32" ( _
        lpCriticalSection As CRITICAL_SECTION)

Declare Function llseek _
    Lib "kernel32" _
    Alias "_llseek" ( _
        ByVal hFile As Long, _
        ByVal lOffset As Long, _
        ByVal iOrigin As Long) _
    As Long

Declare Function LoadLibrary _
    Lib "kernel32" _
    Alias "LoadLibraryA" ( _
        ByVal lpLibFileName As String) _
    As Long

Declare Function LoadLibraryEx _
    Lib "kernel32" _
    Alias "LoadLibraryExA" ( _
        ByVal lpLibFileName As String, _
        ByVal hFile As Long, _
        ByVal dwFlags As Long) _
    As Long

Declare Function LoadModule _
    Lib "kernel32" ( _
        ByVal lpModuleName As String, _
        lpParameterBlock As Any) _
    As Long

Declare Function LoadResource _
    Lib "kernel32" ( _
        ByVal hInstance As Long, _
        ByVal hResInfo As Long) _
    As Long

Declare Function LocalAlloc _
    Lib "kernel32" ( _
        ByVal wFlags As Long, _
        ByVal wBytes As Long) _
    As Long

Declare Function LocalCompact _
    Lib "kernel32" ( _
        ByVal uMinFree As Long) _
    As Long

Declare Function LocalFileTimeToFileTime _
    Lib "kernel32" ( _
        lpLocalFileTime As FILETIME, _
        lpFileTime As FILETIME) _
    As Long

Declare Function LocalFlags _
    Lib "kernel32" ( _
        ByVal hMem As Long) _
    As Long

Declare Function LocalFree _
    Lib "kernel32" ( _
        ByVal hMem As Long) _
    As Long

Declare Function LocalHandle _
    Lib "kernel32" ( _
        wMem As Any) _
    As Long

Declare Function LocalLock _
    Lib "kernel32" ( _
        ByVal hMem As Long) _
    As Long

Declare Function LocalReAlloc _
    Lib "kernel32" ( _
        ByVal hMem As Long, _
        ByVal wBytes As Long, _
        ByVal wFlags As Long) _
    As Long

Declare Function LocalShrink _
    Lib "kernel32" ( _
        ByVal hMem As Long, _
        ByVal cbNewSize As Long) _
    As Long

Declare Function LocalSize _
    Lib "kernel32" ( _
        ByVal hMem As Long) _
    As Long

Declare Function LocalUnlock _
    Lib "kernel32" ( _
        ByVal hMem As Long) _
    As Long

Declare Function LockFile _
    Lib "kernel32" ( _
        ByVal hFile As Long, ByVal dwFileOffsetLow As Long, _
        ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, _
        ByVal nNumberOfBytesToLockHigh As Long) _
    As Long

Declare Function LockFileEx _
    Lib "kernel32" ( _
        ByVal hFile As Long, ByVal dwFlags As Long, ByVal dwReserved As Long, _
        ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long, _
        lpOverlapped As OVERLAPPED) _
    As Long

Declare Function LockResource _
    Lib "kernel32" ( _
        ByVal hResData As Long) _
    As Long

Declare Function LogonUser _
    Lib "kernel32" _
    Alias " LogonUserA" ( _
        ByVal lpszUsername As String, ByVal lpszDomain As String, _
        ByVal lpszPassword As String, ByVal dwLogonType As Long, _
        ByVal dwLogonProvider As Long, phToken As Long) _
    As Long

Declare Function lopen _
    Lib "kernel32" _
    Alias "_lopen" ( _
        ByVal lpPathName As String, _
        ByVal iReadWrite As Long) _
    As Long

Declare Function lread _
    Lib "kernel32" _
    Alias "_lread" ( _
        ByVal hFile As Long, _
        lpBuffer As Any, _
        ByVal wBytes As Long) _
    As Long

Declare Function lstrcat _
    Lib "kernel32" _
    Alias "lstrcatA" ( _
        ByVal lpString1 As String, _
        ByVal lpString2 As String) _
    As Long

Declare Function lstrcmp _
    Lib "kernel32" _
    Alias "lstrcmpA" ( _
        ByVal lpString1 As String, _
        ByVal lpString2 As String) _
    As Long

Declare Function lstrcmpi _
    Lib "kernel32" _
    Alias "lstrcmpiA" ( _
        ByVal lpString1 As String, _
        ByVal lpString2 As String) _
    As Long

Declare Function lstrcpy _
    Lib "kernel32" _
    Alias "lstrcpyA" ( _
        ByVal lpString1 As String, _
        ByVal lpString2 As String) _
    As Long

Declare Function lstrcpyn _
    Lib "kernel32" _
    Alias "lstrcpynA" ( _
        ByVal lpString1 As String, _
        ByVal lpString2 As String, _
        ByVal iMaxLength As Long) _
    As Long

Declare Function lstrlen _
    Lib "kernel32" _
    Alias "lstrlenA" ( _
        ByVal lpString As String) _
    As Long

Declare Function lwrite _
    Lib "kernel32" _
    Alias "_lwrite" ( _
        ByVal hFile As Long, _
        ByVal lpBuffer As String, _
        ByVal wBytes As Long) _
    As Long

Declare Function MapViewOfFile _
    Lib "kernel32" ( _
        ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, _
        ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, _
        ByVal dwNumberOfBytesToMap As Long) _
    As Long

Declare Function MapViewOfFileEx _
    Lib "kernel32" ( _
        ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, _
        ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, _
        ByVal dwNumberOfBytesToMap As Long, lpBaseAddress As Any) _
    As Long

Declare Function MoveFile _
    Lib "kernel32" _
    Alias "MoveFileA" ( _
        ByVal lpExistingFileName As String, _
        ByVal lpNewFileName As String) _
    As Long

Declare Function MoveFileEx _
    Lib "kernel32" _
    Alias "MoveFileExA" ( _
        ByVal lpExistingFileName As String, _
        ByVal lpNewFileName As String, _
        ByVal dwFlags As Long) _
    As Long

Declare Function MulDiv _
    Lib "kernel32" ( _
        ByVal nNumber As Long, _
        ByVal nNumerator As Long, _
        ByVal nDenominator As Long) _
    As Long

Declare Function MultiByteToWideChar _
    Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, _
        ByVal lpWideCharStr As String, ByVal cchWideChar As Long) _
    As Long

Declare Function ObjectOpenAuditAlarm _
    Lib "kernel32" _
    Alias "ObjectOpenAuditAlarmA" ( _
        ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, _
        pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As Long, ByVal DesiredAccess As Long, _
        ByVal GrantedAccess As Long, Privileges As PRIVILEGE_SET, ByVal ObjectCreation As Long, _
        ByVal AccessGranted As Long, ByVal GenerateOnClose As Long) _
    As Long

Declare Function OpenEvent _
    Lib "kernel32" _
    Alias "OpenEventA" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal lpName As String) _
    As Long

Declare Function OpenFile _
    Lib "kernel32" ( _
        ByVal lpFileName As String, _
        lpReOpenBuff As OFSTRUCT, _
        ByVal wStyle As Long) _
    As Long

Declare Function OpenFileMapping _
    Lib "kernel32" _
    Alias "OpenFileMappingA" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal lpName As String) _
    As Long

Declare Function OpenMutex _
    Lib "kernel32" _
    Alias "OpenMutexA" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal lpName As String) _
    As Long

Declare Function OpenProcess _
    Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) _
    As Long

Declare Function OpenSemaphore _
    Lib "kernel32" _
    Alias "OpenSemaphoreA" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal lpName As String) _
    As Long

Declare Sub OutputDebugString _
    Lib "kernel32" _
    Alias "OutputDebugStringA" ( _
        ByVal lpOutputString As String)

Declare Function PeekNamedPipe _
    Lib "kernel32" ( _
        ByVal hNamedPipe As Long, lpBuffer As Any, _
        ByVal nBufferSize As Long, lpBytesRead As Long, _
        lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) _
    As Long

Declare Function PrepareTape _
    Lib "kernel32" ( _
        ByVal hDevice As Long, _
        ByVal dwOperation As Long, _
        ByVal bimmediate As Long) _
    As Long

Declare Function PulseEvent _
    Lib "kernel32" ( _
        ByVal hEvent As Long) _
    As Long

Declare Function ReadConsoleOutputAttribute _
    Lib "kernel32" ( _
        ByVal hConsoleOutput As Long, lpAttribute As Long, _
        ByVal nLength As Long, dwReadCoord As COORD, _
        lpNumberOfAttrsRead As Long) _
    As Long

Declare Function ReadConsoleOutputCharacter _
    Lib "kernel32" _
    Alias "ReadConsoleOutputCharacterA" ( _
        ByVal hConsoleOutput As Long, ByVal lpCharacter As String, _
        ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfCharsRead As Long) _
    As Long

Declare Function ReadFile _
    Lib "kernel32" ( _
        ByVal hFile As Long, lpBuffer As Any, _
        ByVal nNumberOfBytesToRead As Long, _
        lpNumberOfBytesRead As Long, _
        lpOverlapped As OVERLAPPED) _
    As Long

Declare Function ReadFileEx _
    Lib "kernel32" ( _
        ByVal hFile As Long, lpBuffer As Any, _
        ByVal nNumberOfBytesToRead As Long, _
        lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) _
    As Boolean

Declare Function ReadProcessMemory _
    Lib "kernel32" ( _
        ByVal hProcess As Long, lpBaseAddress As Any, _
        lpBuffer As Any, ByVal nSize As Long, _
        lpNumberOfBytesWritten As Long) _
    As Long

Declare Function ReleaseMutex _
    Lib "kernel32" ( _
        ByVal hMutex As Long) _
    As Long

Declare Function ReleaseSemaphore _
    Lib "kernel32" ( _
        ByVal hSemaphore As Long, _
        ByVal lReleaseCount As Long, _
        lpPreviousCount As Long) _
    As Long

Declare Function RemoveDirectory _
    Lib "kernel32" _
    Alias "RemoveDirectoryA" ( _
        ByVal lpPathName As String) _
    As Long

Declare Function ResetEvent _
    Lib "kernel32" ( _
        ByVal hEvent As Long) _
    As Long

Declare Function ResumeThread _
    Lib "kernel32" ( _
        ByVal hThread As Long) _
    As Long

Declare Function SetCommBreak _
    Lib "kernel32" ( _
        ByVal nCid As Long) _
    As Long

Declare Function SetCommConfig _
    Lib "kernel32" ( _
        ByVal hCommDev As Long, _
        lpCC As COMMCONFIG, _
        ByVal dwSize As Long) _
    As Boolean

Declare Function SetCommMask _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal dwEvtMask As Long) _
    As Long

Declare Function SetCommState _
    Lib "kernel32" ( _
        ByVal hCommDev As Long, _
        lpDCB As DCB) _
    As Long

Declare Function SetCommTimeouts _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpCommTimeouts As COMMTIMEOUTS) _
    As Long

Declare Function SetComputerName _
    Lib "kernel32" _
    Alias "SetComputerNameA" ( _
        ByVal lpComputerName As String) _
    As Long

Declare Function SetConsoleActiveScreenBuffer _
    Lib "kernel32" ( _
        ByVal hConsoleOutput As Long) _
    As Long

Declare Function SetConsoleCP _
    Lib "kernel32" ( _
        ByVal wCodePageID As Long) _
    As Long

Declare Function SetConsoleCtrlHandler _
    Lib "kernel32" ( _
        ByVal HandlerRoutine As Long, _
        ByVal Add As Long) _
    As Long

Declare Function SetConsoleCursorInfo _
    Lib "kernel32" ( _
        ByVal hConsoleOutput As Long, _
        lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) _
    As Long

Declare Function SetConsoleCursorPosition _
    Lib "kernel32" ( _
        ByVal hConsoleOutput As Long, _
        dwCursorPosition As COORD) _
    As Long

Declare Function SetConsoleMode _
    Lib "kernel32" ( _
        ByVal hConsoleHandle As Long, _
        ByVal dwMode As Long) _
    As Long

Declare Function SetConsoleOutputCP _
    Lib "kernel32" ( _
        ByVal wCodePageID As Long) _
    As Long

Declare Function SetConsoleScreenBufferSize _
    Lib "kernel32" ( _
        ByVal hConsoleOutput As Long, _
        dwSize As COORD) _
    As Long

Declare Function SetConsoleTextAttribute _
    Lib "kernel32" ( _
        ByVal hConsoleOutput As Long, _
        ByVal wAttributes As Long) _
    As Long

Declare Function SetConsoleTitle _
    Lib "kernel32" _
    Alias "SetConsoleTitleA" ( _
        ByVal lpConsoleTitle As String) _
    As Long

Declare Function SetConsoleWindowInfo _
    Lib "kernel32" ( _
        ByVal hConsoleOutput As Long, _
        ByVal bAbsolute As Long, _
        lpConsoleWindow As SMALL_RECT) _
    As Long

Declare Function SetCurrentDirectory _
    Lib "kernel32" _
    Alias "SetCurrentDirectoryA" ( _
        ByVal lpPathName As String) _
    As Long

Declare Function SetDefaultCommConfig _
    Lib "kernel32" _
    Alias "SetDefaultCommConfigA" ( _
        ByVal lpszName As String, _
        lpCC As COMMCONFIG, _
        ByVal dwSize As Long) _
    As Boolean

Declare Function SetEndOfFile _
    Lib "kernel32" ( _
        ByVal hFile As Long) _
    As Long

Declare Function SetEnvironmentVariable _
    Lib "kernel32" _
    Alias "SetEnvironmentVariableA" ( _
        ByVal lpName As String, _
        ByVal lpValue As String) _
    As Long

Declare Function SetErrorMode _
    Lib "kernel32" ( _
        ByVal wMode As Long) _
    As Long

Declare Function SetEvent _
    Lib "kernel32" ( _
        ByVal hEvent As Long) _
    As Long

Declare Sub SetFileApisToANSI _
    Lib "kernel32" ( _
        )

Declare Sub SetFileApisToOEM _
    Lib "kernel32" ( _
        )

Declare Function SetFileAttributes _
    Lib "kernel32" _
    Alias "SetFileAttributesA" ( _
        ByVal lpFileName As String, _
        ByVal dwFileAttributes As Long) _
    As Long

Declare Function SetFilePointer _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal lDistanceToMove As Long, _
        lpDistanceToMoveHigh As Long, _
        ByVal dwMoveMethod As Long) _
    As Long

Declare Function SetFileTime _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpCreationTime As FILETIME, _
        lpLastAccessTime As FILETIME, _
        lpLastWriteTime As FILETIME) _
    As Long

Declare Function SetLocaleInfo _
    Lib "kernel32" _
    Alias "SetLocaleInfoA" ( _
        ByVal Locale As Long, _
        ByVal LCType As Long, _
        ByVal lpLCData As String) _
    As Boolean

Declare Function SetLocalTime _
    Lib "kernel32" ( _
        lpSystemTime As SYSTEMTIME) _
    As Long

Declare Function SetMailslotInfo _
    Lib "kernel32" ( _
        ByVal hMailslot As Long, _
        ByVal lReadTimeout As Long) _
    As Long

Declare Function SetNamedPipeHandleState _
    Lib "kernel32" ( _
        ByVal hNamedPipe As Long, _
        lpMode As Long, _
        lpMaxCollectionCount As Long, _
        lpCollectDataTimeout As Long) _
    As Long

Declare Function SetPriorityClass _
    Lib "kernel32" ( _
        ByVal hProcess As Long, _
        ByVal dwPriorityClass As Long) _
    As Long

Declare Function SetProcessShutdownParameters _
    Lib "kernel32" ( _
        ByVal dwLevel As Long, _
        ByVal dwFlags As Long) _
    As Long

Declare Function SetProcessWorkingSetSize _
    Lib "kernel32" ( _
        ByVal hProcess As Long, _
        ByVal dwMinimumWorkingSetSize As Long, _
        ByVal dwMaximumWorkingSetSize As Long) _
    As Boolean

Declare Function SetStdHandle _
    Lib "kernel32" ( _
        ByVal nStdHandle As Long, _
        ByVal nHandle As Long) _
    As Long

Declare Function SetSystemPowerState _
    Lib "kernel32" ( _
        ByVal fSuspend As Long, _
        ByVal fForce As Long) _
    As Long

Declare Function SetSystemTime _
    Lib "kernel32" ( _
        lpSystemTime As SYSTEMTIME) _
    As Long

Declare Function SetSystemTimeAdjustment _
    Lib "kernel32" ( _
        ByVal dwTimeAdjustment As Long, _
        ByVal bTimeAdjustmentDisabled As Boolean) _
    As Boolean

Declare Function SetTapeParameters _
    Lib "kernel32" ( _
        ByVal hDevice As Long, _
        ByVal dwOperation As Long, _
        lpTapeInformation As Any) _
    As Long

Declare Function SetTapePosition _
    Lib "kernel32" ( _
        ByVal hDevice As Long, ByVal dwPositionMethod As Long, _
        ByVal dwPartition As Long, ByVal dwOffsetLow As Long, _
        ByVal dwOffsetHigh As Long, ByVal bimmediate As Long) _
    As Long

Declare Function SetThreadAffinityMask _
    Lib "kernel32" ( _
        ByVal hThread As Long, _
        ByVal dwThreadAffinityMask As Long) _
    As Long

Declare Function SetThreadContext _
    Lib "kernel32" ( _
        ByVal hThread As Long, _
        lpContext As CONTEXT) _
    As Long

Declare Function SetThreadLocale _
    Lib "kernel32" ( _
        ByVal Locale As Long) _
    As Long

Declare Function SetThreadPriority _
    Lib "kernel32" ( _
        ByVal hThread As Long, _
        ByVal nPriority As Long) _
    As Long

Declare Function SetTimeZoneInformation _
    Lib "kernel32" ( _
        lpTimeZoneInformation As TIME_ZONE_INFORMATION) _
    As Long

Declare Function SetUnhandledExceptionFilter _
    Lib "kernel32" ( _
        ByVal lpTopLevelExceptionFilter As Long) _
    As Long

Declare Function SetupComm _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal dwInQueue As Long, _
        ByVal dwOutQueue As Long) _
    As Long

Declare Function SetVolumeLabel _
    Lib "kernel32" _
    Alias "SetVolumeLabelA" ( _
        ByVal lpRootPathName As String, _
        ByVal lpVolumeName As String) _
    As Long

Declare Function SuspendThread _
    Lib "kernel32" ( _
        ByVal hThread As Long) _
    As Long

Declare Function SystemTimeToFileTime _
    Lib "kernel32" ( _
        lpSystemTime As SYSTEMTIME, _
        lpFileTime As FILETIME) _
    As Long

Declare Function SystemTimeToTzSpecificLocalTime _
    Lib "kernel32" ( _
        lpTimeZoneInformation As TIME_ZONE_INFORMATION, _
        lpUniversalTime As SYSTEMTIME, _
        lpLocalTime As SYSTEMTIME) _
    As Boolean

Declare Function TerminateProcess _
    Lib "kernel32" ( _
        ByVal hProcess As Long, _
        ByVal uExitCode As Long) _
    As Long

Declare Function TerminateThread _
    Lib "kernel32" ( _
        ByVal hThread As Long, _
        ByVal dwExitCode As Long) _
    As Long

Declare Function TlsAlloc _
    Lib "kernel32" () _
    As Long

Declare Function TlsFree _
    Lib "kernel32" ( _
        ByVal dwTlsIndex As Long) _
    As Long

Declare Function TlsGetValue _
    Lib "kernel32" ( _
        ByVal dwTlsIndex As Long) _
    As Long

Declare Function TlsSetValue _
    Lib "kernel32" ( _
        ByVal dwTlsIndex As Long, _
        lpTlsValue As Any) _
    As Long

Declare Function TransactNamedPipe _
    Lib "kernel32" ( _
        ByVal hNamedPipe As Long, lpInBuffer As Any, _
        ByVal nInBufferSize As Long, lpOutBuffer As Any, _
        ByVal nOutBufferSize As Long, lpBytesRead As Long, _
        lpOverlapped As OVERLAPPED) _
    As Long

Declare Function TransmitCommChar _
    Lib "kernel32" ( _
        ByVal nCid As Long, _
        ByVal cChar As Byte) _
    As Long

Declare Function UnhandledExceptionFilter _
    Lib "kernel32" ( _
        ExceptionInfo As EXCEPTION_POINTERS) _
    As Long

Declare Function UnlockFile _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, _
        ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) _
    As Long

Declare Function UnlockFileEx _
    Lib "kernel32" ( _
        ByVal hFile As Long, ByVal dwReserved As Long, _
        ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long, _
        lpOverlapped As OVERLAPPED) _
    As Long

Declare Function UnmapViewOfFile _
    Lib "kernel32" ( _
        lpBaseAddress As Any) _
    As Long

Declare Function UpdateResource _
    Lib "kernel32" _
    Alias "UpdateResourceA" ( _
        ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As String, _
        ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) _
    As Long

Declare Function VerLanguageName _
    Lib "kernel32" _
    Alias "VerLanguageNameA" ( _
        ByVal wLang As Long, _
        ByVal szLang As String, _
        ByVal nSize As Long) _
    As Long

Declare Function WaitCommEvent _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpEvtMask As Long, _
        lpOverlapped As OVERLAPPED) _
    As Long

Declare Function WaitForMultipleObjects _
    Lib "kernel32" ( _
        ByVal nCount As Long, _
        lpHandles As Long, _
        ByVal bWaitAll As Long, _
        ByVal dwMilliseconds As Long) _
    As Long

Declare Function WaitForMultipleObjectsEx _
    Lib "kernel32" ( _
        ByVal nCount As Long, _
        lpHandles As Long, _
        ByVal bWaitAll As Long, _
        ByVal dwMilliseconds As Long, _
        ByVal bAlertable As Long) _
    As Long

Declare Function WaitForSingleObject _
    Lib "kernel32" ( _
        ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) _
    As Long

Declare Function WaitForSingleObjectEx _
    Lib "kernel32" ( _
        ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long, _
        ByVal bAlertable As Long) _
    As Long

Declare Function WaitNamedPipe _
    Lib "kernel32" _
    Alias "WaitNamedPipeA" ( _
        ByVal lpNamedPipeName As String, _
        ByVal nTimeOut As Long) _
    As Long

Declare Function WideCharToMultiByte _
    Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpWideCharStr As String, ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, _
        ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) _
    As Long

Declare Function WinExec _
    Lib "kernel32" ( _
        ByVal lpCmdLine As String, _
        ByVal nCmdShow As Long) _
    As Long

Declare Function WriteConsole _
    Lib "kernel32" _
    Alias "WriteConsoleA" ( _
        ByVal hConsoleOutput As Long, lpBuffer As Any, _
        ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, _
        lpReserved As Any) _
    As Long

Declare Function WriteTapemark _
    Lib "kernel32" ( _
        ByVal hDevice As Long, _
        ByVal dwTapemarkType As Long, _
        ByVal dwTapemarkCount As Long, _
        ByVal bimmediate As Long) _
    As Long


