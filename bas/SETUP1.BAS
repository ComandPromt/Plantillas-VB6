Attribute VB_Name = "basSetup1"
Option Explicit
Option Compare Text

'
' Global Constants
'

Public Enum OverwriteReturnVal
    owYes
    owNo
    owNoToAll
End Enum

'Return values for setup toolkit functions
Global Const gintRET_CONT% = 1
Global Const gintRET_CANCEL% = 2
Global Const gintRET_EXIT% = 3
Global Const gintRET_ABORT% = 4
Global Const gintRET_FATAL% = 5
Global Const gintRET_FINISHEDSUCCESS% = 6 'Used only as parameter to ExitSetup at end of successful install

'Error levels for GetAppRemovalCmdLine()
Global Const APPREMERR_NONE = 0 'no error
Global Const APPREMERR_FATAL = 1 'fatal error
Global Const APPREMERR_NONFATAL = 2 'non-fatal error, user chose to abort
Global Const APPREMERR_USERCANCEL = 3 'user chose to cancel (no error)

'Flag for Path Dialog specifying Source or Dest directory needed
Global Const gstrDIR_SRC$ = "S"
Global Const gstrDIR_DEST$ = "D"

'Beginning of lines in [Files], [Bootstrap], and [Licenses] sections of SETUP.LST
Global Const gstrINI_FILE$ = "File"
Global Const gstrINI_REMOTE$ = "Remote"
Global Const gstrINI_LICENSE$ = "License"
'
' Command line constants
'
Global Const gstrSILENTSWITCH = "s"
Global Const gstrSMSSWITCH = "q"
'
' Icon Information
'
Global Const gsGROUP As String = "Group"
Global Const gsICON As String = "Icon"
Global Const gsTITLE As String = "Title"
Global Const gsICONGROUP As String = "IconGroups"

Global Const gstrINI_BOOTFILES$ = "Bootstrap Files"

'Font info
Global Const gsEXT_FONTTTF As String = "TTF"
Global Const gsEXT_FONTFON As String = "FON"
Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFilename As String) As Long
 
'Registry files (execute them based on .reg extension)
Global Const gsREGEDIT As String = "regedit /s "
Global Const gsEXT_REG As String = "reg"
'
'Type Definitions
'
Type FILEINFO                                               ' Setup information file line format
    intDiskNum As Integer                                   ' disk number
    fSplit As Integer                                       ' split flag
    strSrcName As String                                    ' name of source file
    strDestName As String                                   ' name of destination file
    strDestDir As String                                    ' destination directory
    strRegister As String                                   ' registration info
    fShared As Boolean                                      ' whether the file is shared or private
    fSystem As Boolean                                      ' whether the file is a system file (i.e. should be installed but never removed)
    varDate As Date                                         ' file date
    lFileSize As Long                                       ' file size
    sVerInfo As VERINFO                                     ' file version number
    strReserved As String                                   ' Reserved. Leave empty, or error.
    strProgramIconTitle As String                                ' Caption for icon in program group
    strProgramIconCmdLine As String                         ' Command Line for icon in program group
End Type

Type DISKINFO                                               ' Disk drive information
    lAvail As Long                                          ' Bytes available on drive
    lReq As Long                                            ' Bytes required for setup
    lMinAlloc As Long                                       ' minimum allocation unit
End Type

Type DESTINFO                                               ' save dest dir for certain files
    strAppDir As String
    strAUTMGR32 As String
    strRACMGR32 As String
End Type

Type REGINFO                                                ' save registration info for files
    strFilename As String
    strRegister As String
    
    'The following are used only for remote server registration
    strNetworkAddress As String
    strNetworkProtocol As String
    intAuthentication As Integer
    fDCOM As Boolean      ' True if DCOM, otherwise False
End Type

'
'Global Variables
'
Global gstrSETMSG As String
Global gfRetVal As Integer                                  'return value for form based functions
Global gstrAppName As String                                'name of app being installed
Global gintCabs As Long
Global gstrTitle As String                                  '"setup" name of app being installed
Public gstrDefGroup As String                               'Default name for group -- from setup.lst
Global gstrDestDir As String                                'dest dir for application files
Global gstrAppExe As String                                 'name of app .EXE being installed
Public gstrAppToUninstall As String                         ' Name of app exe/ocx/dll to be uninstalled.  Should be the same as gstrAppExe in most cases.
Global gstrSrcPath As String                                'path of source files
Global gstrSetupInfoFile As String                          'pathname of SETUP.LST file
Global gstrWinDir As String                                 'windows directory
Global gstrFontDir As String                                'windows\font directory
Global gstrWinSysDir As String                              'windows\system directory
Global gsDiskSpace() As DISKINFO                            'disk space for target drives
Global gstrDrivesUsed As String                             'dest drives used by setup
Global glTotalCopied As Long                                'total bytes copied so far
Global gintCurrentDisk As Integer                           'current disk number being installed
Global gsDest As DESTINFO                                   'dest dirs for certain files
Global gstrAppRemovalLog As String                           'name of the app removal logfile
Global gstrAppRemovalEXE As String                           'name of the app removal executable
Global gfAppRemovalFilesMoved As Boolean                     'whether or not the app removal files have been moved to the application directory
Global gfForceUseDefDest As Boolean                         'If set to true, then the user will not be prompted for the destination directory
Global fMainGroupWasCreated As Boolean                     'Whether or not a main folder/group has been created
Public gfRegDAO As Boolean                                 ' If this gets set to true in the code, then
                                                           ' we need to add some registration info for DAO
                                                           ' to the registry.

Global gsCABNAME As String
Global gsTEMPDIR As String

Global Const gsINI_CABNAME As String = "Cab"
Global Const gsINI_TEMPDIR As String = "TmpDir"
'
'Form/Module Constants
'

'SetFileTime junk
Public Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_READ = &H80000000
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4

Public Declare Function LocalFileTimeToFileTime Lib "Kernel32" (lpFileTime As FileTime, lpLocalFileTime As FileTime) As Long
Public Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" (ByVal lpFilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function SetFileTime Lib "Kernel32" (ByVal hFile As Long, lpCreationTime As FileTime, lpLastAccessTime As FileTime, lpLastWriteTime As FileTime) As Long
Public Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Declare Function SystemTimeToFileTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FileTime) As Long
Public Declare Function VariantChangeTypeEx Lib "oleaut32.dll" (ByVal pvArgDest As Long, ByVal pvArgSrc As Long, ByVal LCID As Long, ByVal wFlags As Integer, ByVal VarType As Integer) As Long
Public Declare Function VariantTimeToSystemTime Lib "oleaut32.dll" (ByVal vtime As Date, lpSystemTime As SYSTEMTIME) As Long

'Possible ProgMan actions
Const mintDDE_ITEMADD% = 1                                  'AddProgManItem flag
Const mintDDE_GRPADD% = 2                                   'AddProgManGroup flag

'Special file names
Const mstrFILE_APPREMOVALLOGBASE$ = "ST6UNST"               'Base name of the app removal logfile
Const mstrFILE_APPREMOVALLOGEXT$ = ".LOG"                   'Default extension for the app removal logfile
Const mstrFILE_AUTMGR32 = "AUTMGR32.EXE"
Const mstrFILE_RACMGR32 = "RACMGR32.EXE"
Const mstrFILE_RICHED32$ = "RICHED32.DLL"

'Name of temporary file used for concatenation of split files
Const mstrCONCATFILE$ = "VB5STTMP.CCT"

'setup information file registration macros
Const mstrDLLSELFREGISTER$ = "$(DLLSELFREGISTER)"
Const mstrEXESELFREGISTER$ = "$(EXESELFREGISTER)"
Const mstrTLBREGISTER$ = "$(TLBREGISTER)"
Const mstrREMOTEREGISTER$ = "$(REMOTE)"
Const mstrVBLREGISTER$ = "$(VBLREGISTER)"  ' Bug 5-8039

'
'Form/Module Variables
'
Private msRegInfo() As REGINFO                                  'files to be registered
Private mlTotalToCopy As Long                                   'total bytes to copy
Private mintConcatFile As Integer                               'handle of dest file for concatenation
Private mlSpaceForConcat As Long                                'extra space required for concatenation
Private mstrConcatDrive As String                               'drive to use for concatenation
Private mstrVerTmpName As String                                'temp file name for VerInstallFile API

' Hkey cache (used for logging purposes)
Private Type HKEY_CACHE
    hKey As Long
    strHkey As String
End Type

Private hkeyCache() As HKEY_CACHE

' Registry manipulation API's (32-bit)
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Const ERROR_SUCCESS = 0&
Const ERROR_NO_MORE_ITEMS = 259&

Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4


Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function OSRegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function OSRegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String) As Long
Declare Function OSRegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal iSubKey As Long, ByVal lpszName As String, ByVal cchName As Long) As Long
Declare Function OSRegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function OSRegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, lpdwType As Long, lpbData As Any, cbData As Long) As Long
Declare Function OSRegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function OSRegSetValueNumEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Declare Sub lstrcpyn Lib "Kernel32" (ByVal strDest As String, ByVal strSrc As Any, ByVal lBytes As Long)
Declare Function GetCurrentProcessId Lib "Kernel32" () As Long
Declare Function ExtractFileFromCab Lib "vb6stkit.dll" (ByVal Cab As String, ByVal File As String, ByVal Dest As String, ByVal iCab As Long, ByVal sSrc As String) As Long
'Reboot info
Public Const ANYSIZE_ARRAY = 1

Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Type LUID_AND_ATTRIBUTES
        pLuid As LARGE_INTEGER
        Attributes As Long
End Type

Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Public Const TOKEN_ADJUST_PRIVILEGES = 32
Public Const TOKEN_QUERY = 8
Public Const SE_PRIVILEGE_ENABLED As Long = 2

Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Declare Function GetCurrentProcess Lib "Kernel32" () As Long
Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
'Exit the program and return an error code
Private Declare Sub ExitProcess Lib "Kernel32" (ByVal uExitCode As Long)
'-----------------------------------------------------------
' SUB: AddPerAppPath
'
' Adds an application's full pathname and per-app path to the
'   system registry (this is currently only meaningful to
'   Windows 95).
'
' IN: [strAppExe] - app EXE name, not including path
'     [strAppDir] - full path of EXE, not including filename
'     [strAppPath] - per-app path for this application
'       (semicolon-separated list of directory path names)
'       If this is the empty string (""), no per-app path
'       is registered, but the full pathname of the
'       exe IS still registered.
'
' OUT:
'   Example registry entries:
'     HKEY_LOCAL_MACHINE\[strPathsBaseKeyName]\MyApp.Exe
'       [Default]=C:\Program Files\MyApp\MyApp.Exe
'       [Path]=C:\Program Files\MyApp;C:\Program Files\MyApp\System
'
'-----------------------------------------------------------
'
Sub AddPerAppPath(ByVal strAppExe As String, ByVal strAppDir As String, ByVal strPerAppPath As String)
    If Not TreatAsWin95() Then
        Exit Sub
    End If
    
    Dim strPathsBaseKeyName As String
    Const strAppPaths$ = "App Paths"
    Const strAppPathKeyName = "Path"
    Dim fOk As Boolean
    Dim hKey As Long
    
    AddDirSep strAppDir
    
    ' Create the new key, whose name is based on the app's name
   If Not RegCreateKey(HKEY_LOCAL_MACHINE, RegPathWinCurrentVersion(), strAppPaths & gstrSEP_DIR & strAppExe, hKey) Then
        GoTo Err
    End If
    
    fOk = True
    
    ' Default value indicates full EXE pathname
    fOk = fOk And RegSetStringValue(hKey, "", strAppDir & strAppExe)
    
    ' [Path] value indicates the per-app path
    If strPerAppPath <> "" Then
        fOk = fOk And RegSetStringValue(hKey, strAppPathKeyName, strPerAppPath)
    End If
    
    If Not fOk Then
        GoTo Err
    End If
    
    RegCloseKey hKey
    
    Exit Sub
    
Err:
    MsgError ResolveResString(resERR_REG), vbExclamation Or vbOKOnly, gstrTitle
    '
    ' If we are running an SMS install, we can't continue.
    '
    If gfSMS Then
        ExitSetup frmSetup1, gintRET_FATAL
    End If
End Sub

'-----------------------------------------------------------
' FUNCTION: AddQuotesToFN
'
' Given a pathname (directory and/or filename), returns
'   that pathname surrounded by double quotes if the
'   path contains spaces or commas.  This is required for
'   setting up an icon correctly, since otherwise such paths
'   would be interpreted as a pathname plus arguments.
'-----------------------------------------------------------
'
Function AddQuotesToFN(ByVal strFilename) As String
    If InStr(strFilename, " ") Or InStr(strFilename, ",") Then
        AddQuotesToFN = """" & strFilename & """"
    Else
        AddQuotesToFN = strFilename
    End If
End Function

'-----------------------------------------------------------
' SUB: CalcDiskSpace
'
' Calculates disk space required for installing the files
' listed in the specified section of the setup information
' file (SETUP.LST)
'-----------------------------------------------------------
'
Sub CalcDiskSpace(ByVal strsection As String)
    Static fSplitFile As Integer
    Static lDestFileSpace As Long

    Dim intIdx As Integer
    Dim intDrvIdx As Integer
    Dim sFile As FILEINFO
    Dim strDrive As String
    Dim lThisFileSpace As Long

    intIdx = 1

    On Error GoTo CalcDSError

    '
    'For each file in the specified section, read info from the setup info file
    '
    Do While ReadSetupFileLine(strsection, intIdx, sFile) = True
        '
        'if the file isn't split or if this is the first section of a split file
        '
        If sFile.strDestDir <> vbNullString Then
            fSplitFile = sFile.fSplit

            '
            'Get the dest drive used for this file.  If this is the first file using
            'the drive for a destination, add the drive to the drives used 'table',
            'allocate an array element for the holding the drive info, and get
            'available disk space and minimum allocation unit
            '
            strDrive = Left$(sFile.strDestDir, 1)
        
            intDrvIdx = InStr(gstrDrivesUsed, strDrive)
            If intDrvIdx = 0 Then
                gstrDrivesUsed = gstrDrivesUsed & strDrive
                intDrvIdx = Len(gstrDrivesUsed)

                ReDim Preserve gsDiskSpace(intDrvIdx)
                gsDiskSpace(intDrvIdx).lAvail = GetDiskSpaceFree(strDrive)

                gsDiskSpace(intDrvIdx).lMinAlloc = GetDrivesAllocUnit(strDrive)
            End If

            '
            'Calculate size of the dest final (file size + minimum allocation for drive)
            '
            lThisFileSpace = CalcFinalSize(sFile.lFileSize, strDrive)
            mlTotalToCopy = mlTotalToCopy + lThisFileSpace

            '
            'If the file already exists, then if we copy it at all, we'll be
            'replacing it.  So, we get the size of the existing dest file so
            'that we can subtract it from the amount needed later.
            '
            If FileExists(sFile.strDestDir & sFile.strDestName) Then
                lDestFileSpace = FileLen(sFile.strDestDir & sFile.strDestName)
            Else
                lDestFileSpace = 0
            End If
        End If

        '
        'If file not split, or if the last section of a split file
        '
        If sFile.fSplit = False Then
            '
            'If this is the last section of a split file, then if it's the *largest*
            'split file, set the extra space needed for concatenation to this size
            '
            If fSplitFile = True And lThisFileSpace > mlSpaceForConcat Then
                mlSpaceForConcat = lThisFileSpace
            End If

            '
            'Subtract size of existing dest file, if applicable and then accumulate
            'space required
            '
            lThisFileSpace = lThisFileSpace - lDestFileSpace
            If lThisFileSpace < 0 Then
                lThisFileSpace = 0
            End If

            gsDiskSpace(intDrvIdx).lReq = gsDiskSpace(intDrvIdx).lReq + lThisFileSpace
        End If

        intIdx = intIdx + 1
    Loop

    Exit Sub

CalcDSError:
    MsgError Error$ & vbLf & vbLf & ResolveResString(resCALCSPACE), vbCritical, gstrSETMSG
    ExitSetup frmMessage, gintRET_FATAL
End Sub

'-----------------------------------------------------------
' SUB: CalcFinalSize
'
' Computes the space required for a file of the size
' specified on the given dest path.  This includes the
' file size plus a padding to ensure that the final size
' is a multiple of the minimum allocation unit for the
' dest drive
'-----------------------------------------------------------
'
Function CalcFinalSize(lBaseFileSize As Long, strDestPath As String) As Long
    Dim lMinAlloc As Long
    Dim intPadSize As Long

    lMinAlloc = gsDiskSpace(InStr(gstrDrivesUsed, Left$(strDestPath, 1))).lMinAlloc
    intPadSize = lMinAlloc - (lBaseFileSize Mod lMinAlloc)
    If intPadSize = lMinAlloc Then
        intPadSize = 0
    End If

    CalcFinalSize = lBaseFileSize + intPadSize
End Function

'-----------------------------------------------------------
' SUB: CenterForm
'
' Centers the passed form just above center on the screen
'-----------------------------------------------------------
'
Sub CenterForm(frm As Form)
    SetMousePtr vbHourglass

    frm.Top = (Screen.Height * 0.85) \ 2 - frm.Height \ 2
    frm.Left = Screen.Width \ 2 - frm.Width \ 2

    SetMousePtr gintMOUSE_DEFAULT
End Sub
'-----------------------------------------------------------
' SUB: UpdateDateTime
'
' Updates the date/time for bootstrap files
'-----------------------------------------------------------
'
Sub UpdateDateTime()
    Dim intIdx As Integer
    Dim sFile As FILEINFO
    Dim lTime As FileTime
    Dim hFile As Long
    '
    'For each file in the specified section, read info from the setup info file
    '
    intIdx = 1
    Do While ReadSetupFileLine(gstrINI_BOOTFILES, intIdx, sFile) = True
        Dim sCurDate As String, sFileDate As String
        
        sFileDate = Format(FileDateTime(sFile.strDestDir & sFile.strDestName), "m/d/yyyy h:m")
        sCurDate = Format(Now, "m/d/yyyy h:m")
        
        If sFileDate = sCurDate Then
            lTime = GetFileTime(sFile.varDate)
            hFile = CreateFile(sFile.strDestDir & sFile.strDestName, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
            Call SetFileTime(hFile, lTime, lTime, lTime)
            DoEvents
            CloseHandle hFile
        End If
        intIdx = intIdx + 1
    Loop
    
End Sub

'-----------------------------------------------------------
' FUNCTION: CheckDiskSpace
'
' Reads from the space required array generated by calling
' the 'CalcDiskSpace' function and determines whether there
' is sufficient free space on all of the drives used for
' installation
'
' Returns: True if there is enough space, False otherwise
'-----------------------------------------------------------
'
Function CheckDiskSpace() As Integer
    Static fDontAskOnSpaceErr As Integer

    Dim intIdx As Integer
    Dim intTmpDrvIdx As Integer
    Dim lDiskSpaceLeft As Long
    Dim lMostSpaceLeft As Long
                                             
    '
    'Default to True (enough space on all drives)
    '
    CheckDiskSpace = True

    '
    'For each drive that is the destination for one or more files, compare
    'the space available to the space required.
    '
    For intIdx = 1 To Len(gstrDrivesUsed)
        lDiskSpaceLeft = gsDiskSpace(intIdx).lAvail - gsDiskSpace(intIdx).lReq
        If lDiskSpaceLeft < 0 Then
            GoSub CheckDSAskSpace
        Else
            '
            'If no "TMP" drive was found, or if the "TMP" drive wasn't ready,
            'save the index of the drive and the amount of space on the drive
            'which will have the most free space.  If no "TMP" drive was
            'found in InitDiskInfo(), then this drive will be used as a
            'temporary drive for concatenating split files
            '
            If mstrConcatDrive = vbNullString Then
                If lDiskSpaceLeft > lMostSpaceLeft Then
                    lMostSpaceLeft = lDiskSpaceLeft
                    intTmpDrvIdx = intIdx
                End If
            Else
                '
                '"TMP" drive was specified, so we'll use that
                '
                If Left$(mstrConcatDrive, 1) = Mid$(gstrDrivesUsed, intIdx, 1) Then
                    intTmpDrvIdx = intIdx
                End If
            End If
        End If
    Next

    '
    'If at least one drive was specified as a destination (if there was at least
    'one CalcDiskSpace call in Form_Load of SETUP1.FRM), then subtract the extra
    'space needed for concatenation from either:
    '   The "TMP" drive if available  - OR -
    '   The drive with the most space remaining
    '
    If intTmpDrvIdx > 0 Then
        gsDiskSpace(intTmpDrvIdx).lReq = gsDiskSpace(intTmpDrvIdx).lReq + mlSpaceForConcat
        If gsDiskSpace(intTmpDrvIdx).lAvail < gsDiskSpace(intTmpDrvIdx).lReq Then
            GoSub CheckDSAskSpace
        End If

        '
        'If a "TMP" drive was found, we use it regardless, otherwise we use the drive
        'with the most free space
        '
        If mstrConcatDrive = vbNullString Then
            mstrConcatDrive = Mid$(gstrDrivesUsed, intTmpDrvIdx, 1) & gstrCOLON & gstrSEP_DIR
            AddDirSep mstrConcatDrive
        End If
    End If

    Exit Function

CheckDSAskSpace:
    '
    'if the user hasn't been prompted before in the event of not enough free space,
    'then display table of drive space and allow them to (basically) abort, retry,
    'or ignore.
    '
    If fDontAskOnSpaceErr = False Then
        If gfNoUserInput Then
            If gfSilent = True Then
                LogSilentMsg ResolveResString(resLBLNOSPACE)
            End If
            If gfSMS = True Then
                LogSMSMsg ResolveResString(resLBLNOSPACE)
            End If
            ExitSetup frmSetup1, gintRET_FATAL
        Else
            frmDskSpace.Show vbModal
        End If
        
        If gfRetVal <> gintRET_CONT Then
            CheckDiskSpace = False
            Exit Function
        Else
            fDontAskOnSpaceErr = True
        End If
    End If

    Return
End Function

'-----------------------------------------------------------
' FUNCTION: CheckDrive
'
' Check to see if the specified drive is ready to be read
' from.  In the case of a drive that holds removable media,
' this would mean that formatted media was in the drive and
' that the drive door was closed.
'
' IN: [strDrive] - drive to check
'     [strCaption] - caption if the drive isn't ready
'
' Returns: True if the drive is ready, False otherwise
'-----------------------------------------------------------
'
Function CheckDrive(ByVal strDrive As String, ByVal strCaption As String) As Integer
    Dim strDir As String
    Dim strMsg As String
    Dim fIsUNC As Boolean

    On Error Resume Next

    SetMousePtr vbHourglass

    Do
        Err = 0
        fIsUNC = False
        '
        'Attempt to read the current directory of the specified drive.  If
        'an error occurs, we assume that the drive is not ready
        '
        If IsUNCName(strDrive) Then
            fIsUNC = True
            strDir = Dir$(GetUNCShareName(strDrive))
        Else
            strDir = Dir$(Left$(strDrive, 2))
        End If

        If Err > 0 Then
            If fIsUNC Then
                strMsg = Error$ & vbLf & vbLf & ResolveResString(resCANTREADUNC, "|1", strDrive) & vbLf & vbLf & ResolveResString(resCHECKUNC)
            Else
                strMsg = Error$ & vbLf & vbLf & ResolveResString(resDRVREAD) & strDrive & vbLf & vbLf & ResolveResString(resDRVCHK)
            End If
            If MsgError(strMsg, vbExclamation Or vbRetryCancel, strCaption) = vbCancel Then
                CheckDrive = False
                Err = 0
            End If
        Else
            CheckDrive = True
        End If
        
        If Err And gfNoUserInput = True Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    Loop While Err

    SetMousePtr gintMOUSE_DEFAULT
End Function

'-----------------------------------------------------------
' FUNCTION: CheckOverwritePrivateFile
'
' Checks if a private file that we are about to install
' already exists in the destination directory.  If it
' does, it asks if they want to overwrite the file
'
' IN: [strFN] - Full path of the private file that is
'               about to be installed.
'
'-----------------------------------------------------------
'
Public Function CheckOverwritePrivateFile(ByVal strFN As String) As Boolean
    Static fNoToAll As Boolean
    
    If fNoToAll Then 'They've already said no to all, don't ask again
        CheckOverwritePrivateFile = False
        Exit Function
    End If
    If FileExists(strFN) Then
        Do
            Select Case MsgFunc(ResolveResString(resOVERWRITEPRIVATE) & vbLf & vbLf & ResolveResString(resCANCELSETUP), vbYesNo Or vbDefaultButton1 Or vbExclamation, gstrTitle)
                Case vbYes
                    'The user chose to cancel.  (This is best.)
                    gfDontLogSMS = True  ' Don't log this message if SMS because we already logged the previous one and we can only use 255 characters.
                    MsgError ResolveResString(resCHOOSENEWDEST), vbOKOnly, gstrTitle
                    ExitSetup frmCopy, gintRET_FATAL
                Case Else
                    'One more level of warning to let them know that we highly
                    '  recommend cancelling setup at this point
                    Select Case MsgFunc(ResolveResString(resOVERWRITEPRIVATE2) & vbLf & vbLf & ResolveResString(resVERIFYCONTINUE), vbYesNo Or vbDefaultButton2 Or vbExclamation, gstrTitle)
                        Case vbNo
                            'User chose "no, don't continue"
                            'Repeat the first-level warning
                        Case Else
                            'They decided to continue anyway
                            Exit Do
                        'End Case
                    End Select
                'End Case
            End Select
        Loop
    Else
        CheckOverwritePrivateFile = True
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: CopyFile
'
' Uses the Windows VerInstallFile API to copy a file from
' the specified source location/name to the destination
' location/name.  Split files should be combined via the
' '...Concat...' file routines before calling this
' function.
' If the file is successfully updated and the file is a
' shared file (fShared = True), then the
' files reference count is updated (32-bits only)
'
' IN: [strSrcDir] - directory where source file is located
'     [strDestDir] - destination directory for file
'     [strSrcName] - name of source file
'     [strDestName] - name of destination file
'
' PRECONDITION: NewAction() must have already been called
'               for this file copy (of type either
'               gstrKEY_SHAREDFILE or gstrKEY_PRIVATE --
'               see CopySection for an example of how
'               this works).  See NewAction() and related
'               functions in LOGGING.BAS for comments on
'               using the logging function.
'               Either CommitAction() or AbortAction() will
'               allows be called by this procedure, and
'               should not be done by the caller.
'
' Returns: True if copy was successful, False otherwise
'
' POSTCONDITION: The current action will be either committed or
'                aborted.
'-----------------------------------------------------------
'
Function CopyFile(ByVal strSrcDir As String, ByVal strDestDir As String, ByVal strSrcName As String, ByVal strDestName As String, ByVal fShared As Boolean, ByVal fSystem As Boolean, Optional ByVal fOverWrite As Boolean = False) As Boolean
    Const intUNKNOWN% = 0
    Const intCOPIED% = 1
    Const intNOCOPY% = 2
    Const intFILEUPTODATE% = 3

    '
    'VerInstallFile() Flags
    '
    Const VIFF_FORCEINSTALL% = &H1
    Const VIF_TEMPFILE& = &H1
    Const VIF_SRCOLD& = &H4
    Const VIF_DIFFLANG& = &H8
    Const VIF_DIFFCODEPG& = &H10
    Const VIF_DIFFTYPE& = &H20
    Const VIF_WRITEPROT& = &H40
    Const VIF_FILEINUSE& = &H80
    Const VIF_OUTOFSPACE& = &H100
    Const VIF_ACCESSVIOLATION& = &H200
    Const VIF_SHARINGVIOLATION = &H400
    Const VIF_CANNOTCREATE = &H800
    Const VIF_CANNOTDELETE = &H1000
    Const VIF_CANNOTRENAME = &H2000
    Const VIF_OUTOFMEMORY = &H8000&
    Const VIF_CANNOTREADSRC = &H10000
    Const VIF_CANNOTREADDST = &H20000
    Const VIF_BUFFTOOSMALL = &H40000

    Static fIgnoreWarn As Integer             'user warned about ignoring error?

    Dim strMsg As String
    Dim lRC As Long
    Dim lpTmpNameLen As Long
    Dim intFlags As Integer
    Dim intRESULT As Integer
    Dim fFileAlreadyExisted

    On Error Resume Next

    CopyFile = False

    '
    'Ensure that the source file is available for copying
    '
    If DetectFile(strSrcDir & strSrcName) = vbIgnore Then
        AbortAction
        Exit Function
    End If
    
    '
    ' Make sure that the Destination path (including path, filename, commandline args, etc.
    ' is not longer than the max allowed.
    '
    If Not fCheckFNLength(strDestDir & strDestName) Then
        AbortAction
        strMsg = ResolveResString(resCANTCOPYPATHTOOLONG) & vbLf & vbLf & ResolveResString(resCHOOSENEWDEST) & vbLf & vbLf & strDestDir & strDestName
        Call MsgError(strMsg, vbOKOnly, gstrSETMSG)
        ExitSetup frmCopy, gintRET_FATAL
        Exit Function
    End If
    '
    'Make the destination directory, prompt the user to retry if there is an error
    '
    If Not MakePath(strDestDir) Then
        AbortAction ' Abort file copy
        Exit Function
    End If

    '
    'Make sure we have the LFN (long filename) of the destination directory
    '
    strDestDir = GetLongPathName(strDestDir)
    
    '
    'Setup for VerInstallFile call
    '
    lpTmpNameLen = gintMAX_SIZE
    mstrVerTmpName = String$(lpTmpNameLen, 0)
    intFlags = 0
    If fOverWrite Then intFlags = VIFF_FORCEINSTALL
    fFileAlreadyExisted = FileExists(strDestDir & strDestName)

    intRESULT = intUNKNOWN

    Do While intRESULT = intUNKNOWN
        'VerInstallFile under Windows 95 does not handle
        '  long filenames, so we must give it the short versions
        '  (32-bit only).
        Dim strShortSrcName As String
        Dim strShortDestName As String
        Dim strShortSrcDir As String
        Dim strShortDestDir As String
        
        strShortSrcName = strSrcName
        strShortSrcDir = strSrcDir
        strShortDestName = strDestName
        strShortDestDir = strDestDir
        If Not FileExists(strDestDir & strDestName) Then
            'If the destination file does not already
            '  exist, we create a dummy with the correct
            '  (long) filename so that we can get its
            '  short filename for VerInstallFile.
            Open strDestDir & strDestName For Output Access Write As #1
            Close #1
        End If
    
        On Error GoTo UnexpectedErr
        If Not IsWindowsNT() Then
            Dim strTemp As String
            'This conversion is not necessary under Windows NT
            strShortSrcDir = GetShortPathName(strSrcDir)
            If GetFileName(strSrcName) = strSrcName Then
                strShortSrcName = GetFileName(GetShortPathName(strSrcDir & strSrcName))
            Else
                strTemp = GetShortPathName(strSrcDir & strSrcName)
                strShortSrcName = Mid$(strTemp, Len(strShortSrcDir) + 1)
            End If
            strShortDestDir = GetShortPathName(strDestDir)
            strShortDestName = GetFileName(GetShortPathName(strDestDir & strDestName))
        End If
        On Error Resume Next
            
        lRC = VerInstallFile(intFlags, strShortSrcName, strShortDestName, strShortSrcDir, strShortDestDir, 0&, mstrVerTmpName, lpTmpNameLen)
        If Err <> 0 Then
            '
            'If the version or file expansion DLLs couldn't be found, then abort setup
            '
            ExitSetup frmCopy, gintRET_FATAL
        End If

        If lRC = 0 Then
            '
            'File was successfully installed, increment reference count if needed
            '
            
            'One more kludge for long filenames: VerInstallFile may have renamed
            'the file to its short version if it went through with the copy.
            'Therefore we simply rename it back to what it should be.
            Name strDestDir & strShortDestName As strDestDir & strDestName
            intRESULT = intCOPIED
        ElseIf lRC And VIF_SRCOLD Then
            '
            'Source file was older, so not copied, the existing version of the file
            'will be used.  Increment reference count if needed
            '
            intRESULT = intFILEUPTODATE
        ElseIf lRC And (VIF_DIFFLANG Or VIF_DIFFCODEPG Or VIF_DIFFTYPE) Then
            '
            'We retry and force installation for these cases.  You can modify the code
            'here to prompt the user about what to do.
            '
            intFlags = VIFF_FORCEINSTALL
        ElseIf lRC And VIF_WRITEPROT Then
            strMsg = ResolveResString(resWRITEPROT)
            GoSub CFMsg
        ElseIf lRC And VIF_FILEINUSE Then
            strMsg = ResolveResString(resINUSE)
            GoSub CFMsg
        ElseIf lRC And VIF_OUTOFSPACE Then
            strMsg = ResolveResString(resOUTOFSPACE) & Left$(strDestDir, 2)
            GoSub CFMsg
        ElseIf lRC And VIF_ACCESSVIOLATION Then
            strMsg = ResolveResString(resACCESSVIOLATION)
            GoSub CFMsg
        ElseIf lRC And VIF_SHARINGVIOLATION Then
            strMsg = ResolveResString(resSHARINGVIOLATION)
            GoSub CFMsg
        ElseIf lRC And VIF_OUTOFMEMORY Then
            strMsg = ResolveResString(resOUTOFMEMORY)
            GoSub CFMsg
        Else
            '
            ' For these cases, we generically report the error and do not install the file
            ' unless this is an SMS install; in which case we abort.
            '
            If lRC And VIF_CANNOTCREATE Then
                strMsg = ResolveResString(resCANNOTCREATE)
            ElseIf lRC And VIF_CANNOTDELETE Then
                strMsg = ResolveResString(resCANNOTDELETE)
            ElseIf lRC And VIF_CANNOTRENAME Then
                strMsg = ResolveResString(resCANNOTRENAME)
            ElseIf lRC And VIF_CANNOTREADSRC Then
                strMsg = ResolveResString(resCANNOTREADSRC)
            ElseIf lRC And VIF_CANNOTREADDST Then
                strMsg = ResolveResString(resCANNOTREADDST)
            ElseIf lRC And VIF_BUFFTOOSMALL Then
                strMsg = ResolveResString(resBUFFTOOSMALL)
            End If

            strMsg = strMsg & ResolveResString(resNOINSTALL)
            MsgError strMsg, vbOKOnly Or vbExclamation, gstrTitle
            If gfSMS Then
                ExitSetup frmSetup1, gintRET_FATAL
            End If
            intRESULT = intNOCOPY
        End If
    Loop

    '
    'If there was a temp file left over from VerInstallFile, remove it
    '
    If lRC And VIF_TEMPFILE Then
        Kill mstrVerTmpName
    End If

    'Abort or commit the current Action, and do reference counting
    Select Case intRESULT
        Case intNOCOPY
            AbortAction
        Case intCOPIED
            DecideIncrementRefCount strDestDir & strDestName, fShared, fSystem, fFileAlreadyExisted
            If (Extension(strDestName) = gsEXT_FONTFON) Or (Extension(strDestName) = gsEXT_FONTTTF) Then
                'do nothing
            Else
                AddActionNote ResolveResString(resLOG_FILECOPIED)
                CommitAction
            End If
            CopyFile = True
        Case intFILEUPTODATE
            DecideIncrementRefCount strDestDir & strDestName, fShared, fSystem, fFileAlreadyExisted
            If (Extension(strDestName) = gsEXT_FONTFON) Or (Extension(strDestName) = gsEXT_FONTTTF) Then
                'do nothing
            Else
                AddActionNote ResolveResString(resLOG_FILECOPIED)
                CommitAction
            End If
            CopyFile = True
        Case Else
            AbortAction ' Defensive - this shouldn't be reached
        'End Case
    End Select

    Exit Function

UnexpectedErr:
    MsgError Error$ & vbLf & vbLf & ResolveResString(resUNEXPECTED), vbOKOnly Or vbExclamation, gstrTitle
    ExitSetup frmCopy, gintRET_FATAL
    
CFMsg: '(Subroutine)
    Dim intMsgRet As Integer
    strMsg = strDestDir & strDestName & vbLf & vbLf & strMsg
    intMsgRet = MsgError(strMsg, vbAbortRetryIgnore Or vbExclamation Or vbDefaultButton2, gstrTitle)
    If gfNoUserInput Then intMsgRet = vbAbort
    Select Case intMsgRet
        Case vbAbort
            ExitSetup frmCopy, gintRET_ABORT
        Case vbIgnore
            If fIgnoreWarn = True Then
                intRESULT = intNOCOPY
            Else
                fIgnoreWarn = True
                strMsg = strMsg & vbLf & vbLf & ResolveResString(resWARNIGNORE)
                If MsgError(strMsg, vbYesNo Or vbQuestion Or vbDefaultButton2, gstrTitle) = vbYes Then
                    intRESULT = intNOCOPY
                Else
                    'Will retry
                End If
            End If
        'End Case
    End Select

    Return
End Function

'-----------------------------------------------------------
' SUB: CopySection
'
' Attempts to copy the files that need to be copied from
' the named section of the setup info file (SETUP.LST)
'
' IN: [strSection] - name of section to copy files from
'
'-----------------------------------------------------------
'
Sub CopySection(ByVal strsection As String)
    Dim intIdx As Integer
    Dim fSplit As Integer
    Dim fSrcVer As Integer
    Dim sFile As FILEINFO
    Dim strLastFile As String
    Dim intRC As Integer
    Dim lThisFileSize As Long
    Dim strSrcDir As String
    Dim strDestDir As String
    Dim strSrcName As String
    Dim strDestName As String
    Dim strRegister As String
    Dim sSrcVerInfo As VERINFO
    Dim sDestVerInfo As VERINFO
    Dim fFileWasUpToDate As Boolean
    Dim strMultDirBaseName As String
    Dim strMsg As String
    Dim strDetectPath As String
    Dim fRemoteReg As Boolean
    Dim fOverWrite As Boolean
    Dim frm As frmOverwrite
    Static fOverwriteAll As Boolean
    
    On Error Resume Next

    UpdateDateTime
    strMultDirBaseName = ResolveResString(resCOMMON_MULTDIRBASENAME)
    intIdx = 1

    If Not FileExists(gsTEMPDIR) Then
        MkDir gsTEMPDIR
    End If
    '
    'For each file in the specified section, read info from the setup info file
    '
    Do While ReadSetupFileLine(strsection, intIdx, sFile) = True
        fFileWasUpToDate = False
        
        '
        'If last result was IGNORE, and if this is an extent of a split file,
        'then no need to process this chunk of the file either
        '
        
        If sFile.strSrcName = gstrSEP_AMPERSAND & gstrFILE_MDAG Then
            'We don't need to extract mdac_typ twice
            GoTo CSContinue
        End If
        ExtractFileFromCab GetShortPathName(gsCABNAME), sFile.strSrcName, gsTEMPDIR & sFile.strDestName, gintCabs, gstrSrcPath
        If FileExists(gsTEMPDIR & sFile.strDestName) Then
            sFile.strSrcName = gsTEMPDIR & sFile.strDestName
            sFile.intDiskNum = gintCurrentDisk
        End If
        If intRC = vbIgnore And sFile.strDestName = strDestName Then
            GoTo CSContinue
        End If
        intRC = 0

        '
        ' If a new disk is called for, or if for some reason we can't find the
        ' source path (user removed the install floppy, for instance) then
        ' prompt for the next disk.  The PromptForNextDisk function won't
        ' actually prompt the user unless it determines that the source drive
        ' contains removeable media or is a network connection.  Also, we don't
        ' prompt if this is a silent install.  It will fail later on a silent
        ' install when it can't find the file.
        '
        If gfNoUserInput = False And (sFile.intDiskNum <> gintCurrentDisk Or DirExists(gstrSrcPath) = False) Then
            PromptForNextDisk sFile.intDiskNum, sFile.strSrcName
        End If

        strSrcName = sFile.strSrcName
        '
        ' The file could exist in either the main source directory or
        ' in a subdirectory named DISK1, DISK2, etc.  Set the appropriate
        ' path.  If it's in neither place, it is an error and will be
        ' handled later.
        '
        If FileExists(strSrcName) = True Then
            strSrcDir = gsTEMPDIR
        'ElseIf FileExists(gstrSrcPath & ".." & gstrSEP_DIR & strMultDirBaseName & Format(sFile.intDiskNum) & gstrSEP_DIR & strSrcName) = True Then
            'strSrcDir = ResolveDir(gstrSrcPath & ".." & gstrSEP_DIR & strMultDirBaseName & Format(sFile.intDiskNum) & gstrSEP_DIR, False, False)
            'gstrSrcPath = strSrcDir
        Else
            '
            ' Can't find the file.
            '
            If DirExists(gstrSrcPath & strMultDirBaseName & Format(sFile.intDiskNum)) = True Then
                strDetectPath = gstrSrcPath & strMultDirBaseName & Format(sFile.intDiskNum)
            Else
                strDetectPath = gstrSrcPath
            End If
            strMsg = ResolveResString(resCOMMON_CANTFINDSRCFILE, "|1", strDetectPath & gstrSEP_DIR & strSrcName)
            MsgError strMsg, vbExclamation Or vbOKOnly, gstrTitle
            ExitSetup frmCopy, gintRET_FATAL
        End If

        '
        'if the file isn't split, or if this is the first section of a split file
        '
        If sFile.strDestDir <> vbNullString Then
            fSplit = sFile.fSplit

            strDestDir = sFile.strDestDir
            strDestName = sFile.strDestName
            
            'We need to go ahead and create the destination directory, or else
            'GetLongPathName() may fail
            If Not MakePath(strDestDir) Then
                intRC = vbIgnore
            End If
            
            If intRC <> vbIgnore Then
                Err = 0
                strDestDir = GetLongPathName(strDestDir)

                frmCopy.lblDestFile.Caption = strDestDir & sFile.strDestName
                frmCopy.lblDestFile.Refresh

                If UCase(strDestName) = gstrFILE_AXDIST Then
                    '
                    ' AXDIST.EXE is installed temporarily.  We'll be
                    ' deleting it at the end of setup.  Set gfAXDist = True
                    ' so we know we need to delete it later.
                    '
                    NewAction gstrKEY_TEMPFILE, """" & strDestDir & strDestName & """"
                    gfAXDist = True
                    gstrAXDISTInstallPath = strDestDir & strDestName
                ElseIf UCase(strDestName) = gstrFILE_MDAG Then
                    '
                    ' mdac_typ.EXE is installed temporarily.  We'll be
                    ' deleting it at the end of setup.  Set mdag = True
                    ' so we know we need to delete it later.
                    '
                    NewAction gstrKEY_TEMPFILE, """" & strDestDir & strDestName & """"
                    gfMDag = True
                    gstrMDagInstallPath = strDestDir & strDestName
                ElseIf UCase(strDestName) = gstrFILE_WINT351 Then
                    '
                    ' WINt351.EXE is installed temporarily.  We'll be
                    ' deleting it at the end of setup.  Set WINt351 = True
                    ' so we know we need to delete it later.  (Note, this file
                    ' is only installed if the target is nt3.51.  This is dealt
                    ' with below in this same routine.  )
                    '
                    NewAction gstrKEY_TEMPFILE, """" & strDestDir & strDestName & """"
                    gfWINt351 = True
                    gstrWINt351InstallPath = strDestDir & strDestName
                ElseIf (Extension(sFile.strDestName) = gsEXT_FONTTTF) Then
                    'No new actions for fonts
                ElseIf (Extension(sFile.strDestName) = gsEXT_FONTFON) Then
                    'No new actions for fonts
                ElseIf sFile.fShared Then
                    NewAction gstrKEY_SHAREDFILE, """" & strDestDir & strDestName & """"
                ElseIf sFile.fSystem Then
                    NewAction gstrKEY_SYSTEMFILE, """" & strDestDir & strDestName & """"
                ElseIf (Extension(sFile.strDestName) = gsEXT_REG) Then
                    If Extension(sFile.strRegister) = gsEXT_REG Then
                        'No new actions for registration files.
                    Else
                        NewAction gstrKEY_PRIVATEFILE, """" & strDestDir & strDestName & """"
                    End If
                Else
                    NewAction gstrKEY_PRIVATEFILE, """" & strDestDir & strDestName & """"
                End If
            End If
            
            '
            'If the file info just read from SETUP.LST is the application .EXE
            '(i.e.; it's the value of the AppExe Key in the [Setup] section,
            'then save it's full pathname for later use
            '
            If strDestName = gstrAppExe Then
                '
                'Used for creating a program manager icon in Form_Load of SETUP1.FRM
                'and for registering the per-app path
                '
                gsDest.strAppDir = strDestDir
            End If

            'Special case for RICHED32.DLL
            '-- we only install this file under Windows 95, not under Windows NT (3.51 or 4.0)
            If strDestName = mstrFILE_RICHED32 Then
                If Not IsWindows95() Then
                    'We're not running under Win95 - do not install this file.
                    intRC = vbIgnore
                    LogNote ResolveResString(resCOMMON_RICHED32NOTCOPIED, "|1", strDestName)
                    AbortAction
                End If
            End If
            '
            ' Special case for AXDIST.EXE
            ' If this is Win95 or NT4 and AXDIST.EXE is in the setup list, we need
            ' to execute it when setup1 is complete.  AXDIST.EXE is a self-extracting
            ' exe that installs special files needed for internet functionality.
            '
            If UCase(strDestName) = gstrFILE_AXDIST Then
                '
                ' Don't do anything here if this is not Win95 or NT4.
                '
                If Not TreatAsWin95() Then
                    'We're not running under Win95 or NT4- do not install this file.
                    intRC = vbIgnore
                    LogNote ResolveResString(resCOMMON_AXDISTNOTCOPIED, "|1", strDestName)
                    AbortAction
                    gfAXDist = False
                End If
            End If
            '
            ' Special case for WINt351.EXE
            ' If this is NT3.51 and WINt351.EXE is in the setup list, we need
            ' to execute it when setup1 is complete.  WINt351.EXE is a self-extracting
            ' exe that installs special files needed for internet functionality.
            '
            If UCase(strDestName) = gstrFILE_WINT351 Then
                '
                ' Don't do anything here if this is not NT3.51.
                '
                If TreatAsWin95() Then
                    'We're not running under NT3.51- do not install this file.
                    intRC = vbIgnore
                    LogNote ResolveResString(resCOMMON_WINT351NOTCOPIED, "|1", strDestName)
                    AbortAction
                    gfWINt351 = False
                End If
            End If
            
            strRegister = sFile.strRegister

            lThisFileSize = CalcFinalSize(sFile.lFileSize, sFile.strDestDir)

            '
            'The stuff below trys to save some time by pre-checking whether a file
            'should be installed before a split file is concatenated or before
            'VerInstallFile does its think which involves a full file read (for
            'a compress file) at the minimum.  Basically, if both files have
            'version numbers, they are compared.  If one file has a version number
            'and the other doesn't, the one with the version number is deemed
            '"Newer".  If neither file has a version number, we compare date.
            '
            'Always attempt to get the source file version number.  If the setup
            'info file did not contain a version number (sSrcVerInfo.nMSHi =
            'gintNOVERINFO), we attempt to read the version number from the source
            'file.  Reading the version number from a split file will always fail.
            'That's why it's a good idea to include the version number for a file
            '(especially split ones) in the setup info file (SETUP.LST)
            '
            fSrcVer = True
            sSrcVerInfo = sFile.sVerInfo
            If sSrcVerInfo.FileVerPart1 = gintNOVERINFO Then
                fSrcVer = GetFileVerStruct(strSrcDir & strSrcName, sSrcVerInfo)
            End If

            '
            'If there is an existing destination file with version information, then
            'compare its version number to the source file version number.
            '
            fOverWrite = True
            If intRC <> vbIgnore Then
                fRemoteReg = (sFile.strRegister = mstrREMOTEREGISTER)
                If GetFileVerStruct(strDestDir & strDestName, sDestVerInfo, fRemoteReg) = True Then
                    If fSrcVer = True Then
                        If IsNewerVer(sSrcVerInfo, sDestVerInfo) = False Then
                            '
                            'Existing file is newer than the one we want to install;
                            'prompt user for what to do
                            '
                            
                            If Not fOverwriteAll Then
                                Set frm = New frmOverwrite
                                frm.FileName = strDestDir & strDestName
                                With sDestVerInfo
                                    frm.Version = CStr(.FileVerPart1) & "." & CStr(.FileVerPart2) & "." & _
                                        CStr(.FileVerPart3) & "." & CStr(.FileVerPart4)
                                End With
                                frm.Description = GetFileDescription(strDestDir & strDestName)
                                frm.Show vbModal, frmSetup1
                                If frm.ReturnVal = owNo Then 'overwrite the file
                                    fOverWrite = True
                                ElseIf frm.ReturnVal = owYes Then 'Keep this file
                                    fOverWrite = False
                                ElseIf frm.ReturnVal = owNoToAll Then 'Overwrite all files
                                    fOverWrite = True
                                    fOverwriteAll = True
                                End If
                            End If
                            If Not fOverWrite Then
                                intRC = vbIgnore
                                fFileWasUpToDate = True
                                DecideIncrementRefCount strDestDir & strDestName, sFile.fShared, sFile.fSystem, True
                                If (Extension(sFile.strDestName) = gsEXT_FONTFON) Or (Extension(sFile.strDestName) = gsEXT_FONTTTF) Then
                                    'do nothing
                                Else
                                    AddActionNote ResolveResString(resLOG_FILEUPTODATE)
                                    CommitAction
                                End If
                            End If
                        End If
                    End If
                Else
                    '
                    'If the destination file has no version info, then we'll copy the
                    'source file if it *does* have a version.  If neither file has a
                    'version number, then we compare date.
                    '
                    If sFile.varDate <= FileDateTime(strDestDir & strDestName) Then
                        If Err = 0 Then
                            '
                            'Although neither the source nor the existing file contain version
                            'information, the existing file has a newer date so we'll use it.
                            '
                            If Not fOverwriteAll Then
                                Set frm = New frmOverwrite
                                frm.FileName = strDestDir & strDestName
                                frm.Version = vbNullString
                                frm.Description = GetFileDescription(strDestDir & strDestName)
                                frm.Show vbModal, frmSetup1
                                If frm.ReturnVal = owNo Then 'overwrite the file
                                    fOverWrite = True
                                ElseIf frm.ReturnVal = owYes Then 'Keep this file
                                    fOverWrite = False
                                ElseIf frm.ReturnVal = owNoToAll Then 'Overwrite all files
                                    fOverWrite = True
                                    fOverwriteAll = True
                                End If
                            End If
                            If Not fOverWrite Then
                                intRC = vbIgnore
                                fFileWasUpToDate = True
                                DecideIncrementRefCount strDestDir & strDestName, sFile.fShared, sFile.fSystem, True
                                If (Extension(sFile.strDestName) = gsEXT_FONTFON) Or (Extension(sFile.strDestName) = gsEXT_FONTTTF) Then
                                    'do nothing
                                Else
                                    AddActionNote ResolveResString(resLOG_FILEUPTODATE)
                                    CommitAction
                                End If
                            End If
                        Else
                            Err = 0
                        End If
                    End If
                End If
            End If
            
        End If
        If fOverwriteAll Then fOverWrite = True
        '
        'If the file wasn't split, or if this is the last extent of a split file
        '
        If fSplit = False Then

            '
            'After all of this, if we're still ready to copy, then give it a whirl!
            '
            If intRC <> vbIgnore Then
                ' CopyFile will increment the reference count for us, and will either
                ' commit or abort the current Action.
                'Turn off READONLY flag in case we copy.
                SetAttr strDestDir & strDestName, vbNormal
                If Extension(sFile.strRegister) <> gsEXT_REG Then
                    intRC = IIf(CopyFile(strSrcDir, strDestDir, strDestName, strDestName, sFile.fShared, sFile.fSystem, fOverWrite), 0, vbIgnore)
                End If
            End If

            '
            'Save the paths of certain files for later use, if they were
            'successfully installed or were already on the system
            '
            If (Extension(strDestDir & strDestName) = gsEXT_FONTTTF) Or (Extension(strDestDir & strDestName) = gsEXT_FONTFON) Then
                If AddFontResource(strDestDir & strDestName) <> 0 Then
                    'Success
                Else
                    'Failure
                End If
            End If
            If (intRC = 0 Or fFileWasUpToDate) Then
                Select Case strDestName
                    Case mstrFILE_AUTMGR32
                        '
                        'Used for creating an icon if installed
                        '
                        gsDest.strAUTMGR32 = strDestDir & mstrFILE_AUTMGR32
                    Case mstrFILE_RACMGR32
                        '
                        'Used for creating an icon if installed
                        '
                        gsDest.strRACMGR32 = strDestDir & mstrFILE_RACMGR32
                    'End Case
                End Select
            
                '
                'If we successfully copied the file, and if registration information was
                'specified in the setup info file, save the registration info into an
                'array so that we can register all files requiring it in one fell swoop
                'after all the files have been copied.
                '
                If strRegister <> vbNullString Then
                    Err = 0
                    ReDim Preserve msRegInfo(UBound(msRegInfo) + 1)
    
                    If Err > 0 Then
                        ReDim msRegInfo(0)
                    End If
    
                    msRegInfo(UBound(msRegInfo)).strFilename = strDestDir & strDestName
    
                    Select Case strRegister
                        Case mstrDLLSELFREGISTER, mstrEXESELFREGISTER, mstrTLBREGISTER, mstrVBLREGISTER
                            'Nothing in particular to do
                        Case mstrREMOTEREGISTER
                            'We need to look for and parse the corresponding "RemoteX=..." line
                            If Not ReadSetupRemoteLine(strsection, intIdx, msRegInfo(UBound(msRegInfo))) = True Then
                                MsgError ResolveResString(resREMOTELINENOTFOUND, "|1", strDestName, "|2", gstrINI_REMOTE & Format$(intIdx)), vbExclamation Or vbOKOnly, gstrTitle
                                ExitSetup frmSetup1, gintRET_FATAL
                            End If
                        Case Else
                            '
                            'If the registration info specified the name of a file with
                            'registration info (which we assume if a registration macro
                            'was not specified), then we also assume that, if no path
                            'information is available, this reginfo file is in the same
                            'directory as the file it registers
                            '
                            strRegister = ResolveDestDirs(strRegister)
                            If InStr(strRegister, gstrSEP_DIR) = 0 Then
                                strRegister = strSrcDir & strRegister
                            End If
                        'End Case
                    End Select
    
                    If Extension(strRegister) = gsEXT_REG Then
                        SyncShell gsREGEDIT & strQuoteString(strRegister), INFINITE
                    End If
                    msRegInfo(UBound(msRegInfo)).strRegister = strRegister
                End If
            
            End If
        End If

        strLastFile = sFile.strDestName

CSContinue:
        '
        'If the file wasn't split, or if this was the last extent of a split file, then
        'update the copy status bar.  We need to do the update regardless of whether a
        'file was actually copied or not.
        '
        If sFile.fSplit = False Then
            glTotalCopied = glTotalCopied + lThisFileSize
            UpdateStatus frmCopy.picStatus, glTotalCopied / mlTotalToCopy
        End If

        Dim sCurDate As String, sFileDate As String
        
        sFileDate = Format(FileDateTime(sFile.strDestDir & sFile.strDestName), "m/d/yyyy h:m")
        sCurDate = Format(Now, "m/d/yyyy h:m")
        
        If sFileDate = sCurDate Then
            Dim lTime As FileTime
            Dim hFile As Long
            
            lTime = GetFileTime(sFile.varDate)
            hFile = CreateFile(sFile.strDestDir & sFile.strDestName, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
            Call SetFileTime(hFile, lTime, lTime, lTime)
            DoEvents
            CloseHandle hFile
        Else
            '
            'Give a chance for the 'Cancel' button command to be processed if it was pressed
            '
            DoEvents
        End If
        'Delete the files after copy...
        SetAttr gsTEMPDIR & sFile.strDestName, vbNormal
        Kill gsTEMPDIR & sFile.strDestName
        intIdx = intIdx + 1
    Loop

    Err = 0
End Sub
'-----------------------------------------------------------
' SUB: CreateOSProgramGroup
'
' Calls CreateProgManGroup under Windows NT or
' fCreateShellGroup under Windows 95
'-----------------------------------------------------------
'
Function fCreateOSProgramGroup(frm As Form, ByVal strFolderName As String, ByVal fRetOnErr As Boolean, Optional ByVal fLog As Boolean = True, Optional ByVal fPriv As Boolean = True, Optional ByVal fStartMenu As Boolean = False) As Boolean
    If TreatAsWin95() Then
        fCreateOSProgramGroup = fCreateShellGroup(strFolderName, fRetOnErr, fLog, fPriv, fStartMenu)
    Else
        CreateProgManGroup frm, strFolderName, fRetOnErr, fLog
        fCreateOSProgramGroup = True
    End If
End Function

'-----------------------------------------------------------
' SUB: CreateOSLink
'
' Calls CreateProgManItem under Windows NT or
' CreateFolderLink under Windows 95.
'
' If fLog is missing, the default is True.
'-----------------------------------------------------------
'
Sub CreateOSLink(frm As Form, ByVal strGroupName As String, ByVal strLinkPath As String, ByVal strLinkArguments As String, ByVal strLinkName As String, ByVal fPrivate As Boolean, sParent As String, Optional ByVal fLog As Boolean = True)
    If TreatAsWin95() Then
        CreateShellLink strLinkPath, strGroupName, strLinkArguments, strLinkName, fPrivate, sParent, fLog
    Else
        '
        ' DDE will not work properly if you try to send NT the long filename.  If it is
        ' in quotes, then the parameters get ignored.  If there are no parameters, the
        ' long filename can be used and the following line could be skipped.
        '
        strLinkPath = GetShortPathName(strUnQuoteString(strLinkPath))
        CreateProgManItem frm, strGroupName, strLinkPath & " " & strLinkArguments, strLinkName, fLog
    End If
End Sub

'-----------------------------------------------------------
' SUB: CreateProgManGroup
'
' Creates a new group in the Windows program manager if
' the specified groupname doesn't already exist
'
' IN: [frm] - form containing a label named 'lblDDE'
'     [strGroupName] - text name of the group
'     [fRetOnErr]    - ignored
'     [fLog] - Whether or not to write to the logfile (default
'                is true if missing)
'-----------------------------------------------------------
'
Sub CreateProgManGroup(frm As Form, ByVal strGroupName As String, ByVal fRetOnErr As Boolean, Optional ByVal fLog As Boolean = True)
    '
    'Call generic progman DDE function with flag to add a group
    '
    
    'Perform the DDE to create the group
    PerformDDE frm, strGroupName, vbNullString, vbNullString, mintDDE_GRPADD, fLog
End Sub

'-----------------------------------------------------------
' SUB: CreateProgManItem
'
' Creates (or replaces) a program manager icon in the active
' program manager group
'
' IN: [frm] - form containing a label named 'lblDDE'
'     [strGroupName] - Caption of group in which icon will go.
'     [strCmdLine] - command line for the item/icon,
'                    Ex: 'c:\myapp\myapp.exe'
'                    Note:  If this path contains spaces
'                      or commas, it should be enclosed
'                      with quotes so that it is properly
'                      interpreted by Windows (see AddQuotesToFN)
'     [strIconTitle] - text caption for the icon
'     [fLog] - Whether or not to write to the logfile (default
'                is true if missing)
'
' PRECONDITION: CreateProgManGroup has already been called.  The
'               new icon will be created in the group last created.
'-----------------------------------------------------------
'
Sub CreateProgManItem(frm As Form, ByVal strGroupName As String, ByVal strCmdLine As String, ByVal strIconTitle As String, Optional ByVal fLog As Boolean = True)
    '
    'Call generic progman DDE function with flag to add an item
    '
    PerformDDE frm, strGroupName, strCmdLine, strIconTitle, mintDDE_ITEMADD, fLog
End Sub

'-----------------------------------------------------------
' SUB: fCreateShellGroup
'
' Creates a new program group off of Start>Programs in the
' Windows 95 shell if the specified folder doesn't already exist.
'
' IN: [strFolderName] - text name of the folder.
'                      This parameter may not contain
'                      backslashes.
'                      ex: "My Application" - this creates
'                        the folder Start>Programs>My Application
'     [fRetOnerr] - Whether or not this routine should return if
'                   there is an error creating the group.  If false,
'                   setup aborts and does not return.  Set this to
'                   true if the user can do something to correct the
'                   error.  E.g., they entered a group name in the
'                   Choose Program Group dialog as opposed to calling
'                   this routine when creating the Remote Automation
'                   group in which the user had no control.
'     [fLog] - Whether or not to write to the logfile (default
'                is true if missing)
'-----------------------------------------------------------
'
Function fCreateShellGroup(ByVal strFolderName As String, fRetOnErr As Boolean, Optional ByVal fLog As Boolean = True, Optional ByVal fPriv As Boolean = True, Optional ByVal fStartMenu As Boolean = False) As Boolean
    Dim oMalloc As IVBMalloc

    ReplaceDoubleQuotes strFolderName
    
    If strFolderName = "" Then
        Exit Function
    End If

Retry:
    
    Dim fSuccess As Boolean
    Dim sPath As String
    Dim IDL As Long
    Dim lPrograms As SpecialFolderIDs
    
    If IsWindows95() Then
        If fStartMenu Then
            lPrograms = sfidSTARTMENU
        Else
            lPrograms = sfidPROGRAMS
        End If
    Else
        If fPriv Then
            If fStartMenu Then
                lPrograms = sfidSTARTMENU
            Else
                lPrograms = sfidPROGRAMS
            End If
        Else
            If fStartMenu Then
                lPrograms = sfidCOMMON_STARTMENU
            Else
                lPrograms = sfidCOMMON_PROGRAMS
            End If
        End If
    End If
    ' Fill the item id list with the pointer of each folder item, rtns 0 on success
    If SHGetSpecialFolderLocation(frmSetup1.hwnd, lPrograms, IDL) = NOERROR Then
        sPath = String$(gintMAX_PATH_LEN, 0)
        SHGetPathFromIDListA IDL, sPath
        SHGetMalloc oMalloc
        oMalloc.Free IDL
        sPath = StringFromBuffer(sPath)
    End If
    AddDirSep sPath
    sPath = sPath & strFolderName
    fSuccess = MakePath(sPath)
    If Not fSuccess Then
        If gfNoUserInput Or (MsgError(ResolveResString(resCANTCREATEPROGRAMGROUP, "|1", strFolderName), vbRetryCancel Or vbExclamation, gstrTitle)) = vbCancel Then
            ExitSetup frmSetup1, gintRET_EXIT
            GoTo Retry
        End If
        '
        ' Determine if we should return so the user can
        ' correct the situation.
        '
        If Not fRetOnErr Then
            '
            ' Return so we can exit setup.
            '
            GoTo Retry
        End If
    End If

    
    fCreateShellGroup = fSuccess
End Function

'-----------------------------------------------------------
' SUB: CreateShellLink
'
' Creates (or replaces) a link in either Start>Programs or
' any of its immediate subfolders in the Windows 95 shell.
'
' IN: [strLinkPath] - full path to the target of the link
'                     Ex: 'c:\Program Files\My Application\MyApp.exe"
'     [strLinkArguments] - command-line arguments for the link
'                     Ex: '-f -c "c:\Program Files\My Application\MyApp.dat" -q'
'     [strLinkName] - text caption for the link
'     [fLog] - Whether or not to write to the logfile (default
'                is true if missing)
'
' OUT:
'   The link will be created in the folder strGroupName

'-----------------------------------------------------------
'
Sub CreateShellLink(ByVal strLinkPath As String, ByVal strGroupName As String, ByVal strLinkArguments As String, ByVal strLinkName As String, ByVal fPrivate As Boolean, sParent As String, Optional ByVal fLog As Boolean = True)
    
    If fLog Then
        NewAction gstrKEY_SHELLLINK, """" & strUnQuoteString(strGroupName) & """" & ", " & """" & strUnQuoteString(strLinkName) & """"
    End If
    
    'ReplaceDoubleQuotes strLinkName
    strLinkName = strUnQuoteString(strLinkName)
    strLinkPath = strUnQuoteString(strLinkPath)
    

Retry:

    Dim fSuccess As Boolean
    fSuccess = OSfCreateShellLink(strGroupName & vbNullChar, strLinkName, strLinkPath, strLinkArguments & vbNullChar, fPrivate, sParent) 'the path should never be enclosed in double quotes
    If fSuccess Then
        If fLog Then
            CommitAction
        End If
    Else
        Dim intMsgRet As Integer
        intMsgRet = MsgError(ResolveResString(resCANTCREATEPROGRAMICON, "|1", strLinkName), vbAbortRetryIgnore Or vbExclamation, gstrTitle)
        If gfNoUserInput Then
            intMsgRet = vbAbort
        End If
        Select Case intMsgRet
            Case vbAbort
                ExitSetup frmSetup1, gintRET_ABORT
                GoTo Retry
            Case vbRetry
                GoTo Retry
            Case vbIgnore
                If fLog Then
                    AbortAction
                End If
            'End Case
        End Select
    End If
End Sub

'-----------------------------------------------------------
' FUNCTION: DecideIncrementRefCount
'
' Increments the reference count of a file under 32-bits
' if the file is a shared file.
'
' IN: [strFullPath] - full pathname of the file to reference
'                     count.  Example:
'                     'C:\MYAPP\MYAPP.DAT'
'     [fShared] - whether the file is shared or private
'     [fSystem] - The file is a system file
'     [fFileAlreadyExisted] - whether or not the file already
'                             existed on the hard drive
'                             before our setup program
'-----------------------------------------------------------
'
Sub DecideIncrementRefCount(ByVal strFullPath As String, ByVal fShared As Boolean, ByVal fSystem As Boolean, ByVal fFileAlreadyExisted As Boolean)
    'Reference counting takes place under both Windows 95 and Windows NT
    If fShared Or fSystem Then
        IncrementRefCount strFullPath, fFileAlreadyExisted
    End If
End Sub
            
'-----------------------------------------------------------
' FUNCTION: DetectFile
'
' Detects whether the specified file exists.  If it can't
' be found, the user is given the opportunity to abort,
' retry, or ignore finding the file.  This call is used,
' for example, to ensure that a floppy with the specified
' file name is in the drive before continuing.
'
' IN: [strFileName] - name of file to detect, usually
'                     should include full path, Example:
'                     'A:\MYAPP.DAT'
'
' Returns: TRUE if the file was detected, vbignore if
'          the user chose ignore when the file couldn't
'          be found, or calls ExitSetup upon 'Abort'
'-----------------------------------------------------------
'
Function DetectFile(ByVal strFilename As String) As Integer
    Dim strMsg As String

    DetectFile = True
                      
    Do While FileExists(strFilename) = False


        strMsg = ResolveResString(resCANTOPEN) & vbLf & vbLf & strFilename
        Select Case MsgError(strMsg, vbAbortRetryIgnore Or vbExclamation Or IIf(gfNoUserInput, vbDefaultButton1, vbDefaultButton2), gstrSETMSG)
            Case vbAbort
                ExitSetup frmCopy, gintRET_ABORT
            Case vbIgnore
                DetectFile = vbIgnore
                Exit Do
            'End Case
        End Select
    Loop
End Function


'-----------------------------------------------------------
' SUB: EtchedLine
'
' Draws an 'etched' line upon the specified form starting
' at the X,Y location passed in and of the specified length.
' Coordinates are in the current ScaleMode of the passed
' in form.
'
' IN: [frmEtch] - form to draw the line upon
'     [intX1] - starting horizontal of line
'     [intY1] - starting vertical of line
'     [intLength] - length of the line
'-----------------------------------------------------------
'
Sub EtchedLine(frmEtch As Form, ByVal intX1 As Integer, ByVal intY1 As Integer, ByVal intLength As Integer)
    Const lWHITE& = vb3DHighlight
    Const lGRAY& = vb3DShadow

    frmEtch.Line (intX1, intY1)-(intX1 + intLength, intY1), lGRAY
    frmEtch.Line (frmEtch.CurrentX + 5, intY1 + 20)-(intX1 - 5, intY1 + 20), lWHITE
End Sub

'-----------------------------------------------------------
' SUB: ExeSelfRegister
'
' Synchronously runs the file passed in (which should be
' an executable file that supports the /REGSERVER switch,
' for instance, a VB5 generated ActiveX Component .EXE).
'
' IN: [strFileName] - .EXE file to register
'-----------------------------------------------------------
'
Sub ExeSelfRegister(ByVal strFilename As String)
    Const strREGSWITCH$ = " /REGSERVER"

    Dim fShell As Integer

    '
    'Synchronously shell out and run the .EXE with the self registration switch
    '
    fShell = SyncShell(AddQuotesToFN(strFilename) & strREGSWITCH, INFINITE, , True)
    frmSetup1.Refresh
End Sub

'-----------------------------------------------------------
' SUB: ExitSetup
'
' Handles shutdown of the setup app.  Depending upon the
' value of the intExitCode parm, may prompt the user and
' exit the sub if the user chooses to cancel the exit
' process.
'
' IN: [frm] - active form to unload upon exit
'     [intExitCode] - code specifying exit action
'-----------------------------------------------------------
'
Sub ExitSetup(frm As Form, intExitCode As Integer)
    Const sKEY As String = "Software\Microsoft\Windows\CurrentVersion\RunOnce\Setup"
    Const sValue As String = "Configuring Data Access"
    
    Const iSUCCESS = 0
    Const iFAIL = 1
    Dim strMsg As String
    Dim strSilent As String
    Dim fNeedReboot As Boolean
    
    On Error Resume Next
    '
    ' If we aren't running in silent or sms mode give
    ' the user a chance to try again, if applicable.
    '
    If Not gfNoUserInput Then
        Select Case intExitCode
            Case gintRET_EXIT
                '
                'If user chose an Exit or Cancel button
                '
                If MsgWarning(ResolveResString(resASKEXIT), vbQuestion Or vbYesNo Or vbDefaultButton2, gstrTitle) = vbNo Then
                    Exit Sub
                End If
            Case gintRET_ABORT
                '
                'If user chose to abort before a pending action
                '
                strMsg = ResolveResString(resINCOMPLETE) & vbLf & vbLf & ResolveResString(resQUITNOW) & vbLf & vbLf
                strMsg = strMsg & ResolveResString(resQUITSETUP)
                If MsgWarning(strMsg, vbQuestion Or vbYesNo Or IIf(gfNoUserInput, vbDefaultButton1, vbDefaultButton2), gstrSETMSG) = vbNo Then
                    Exit Sub
                End If
            'End Case
        End Select
    End If

    'Abort any pending actions
    While fWithinAction()
        AbortAction
    Wend
    
    Close

    '
    'Clean up any temporary files from VerInstallFile or split file concatenation
    '
    Kill mstrVerTmpName
    If mintConcatFile > 0 Then
        Close mintConcatFile
        Kill mstrConcatDrive & mstrCONCATFILE
    End If

    If frm.hwnd <> frmSetup1.hwnd Then
        Unload frm
    End If
    
    If frmSetup1.Visible Then frmSetup1.SetFocus

    '
    'Give appropriate ending message depending upon exit code
    '
    Select Case intExitCode
        Case gintRET_EXIT, gintRET_ABORT
            gfSMSStatus = False
            strMsg = ResolveResString(resINTERRUPTED, "|1", gstrAppName) & vbLf & vbLf & ResolveResString(resCANRUN, "|1", gstrAppName)
            MsgWarning strMsg, vbOKOnly Or vbCritical, gstrTitle
        Case gintRET_FATAL
            gfSMSStatus = False
            MsgError ResolveResString(resERROR, "|1", gstrAppName), vbOKOnly Or vbCritical, gstrTitle
        Case gintRET_FINISHEDSUCCESS
            gfSMSStatus = True
            '
            ' Don't log this message to SMS since it is only a confirmation.
            '
            gfDontLogSMS = True
            MsgFunc ResolveResString(resSUCCESS, "|1", gstrAppName), vbOKOnly, gstrTitle
            
            If IsWindowsNT4WithoutSP2() Then
                'Recommend that the user upgrade to NT 4.0 SP2
                gfDontLogSMS = True
                MsgWarning ResolveResString(resNT4WithoutSP2), vbOKOnly Or vbInformation, gstrTitle
            End If
        Case Else
            strMsg = ResolveResString(resINTERRUPTED, "|1", gstrAppName) & vbLf & vbLf & ResolveResString(resCANRUN, "|1", gstrAppName)
            MsgWarning strMsg, vbOKOnly Or vbCritical, gstrTitle
        'End Case
    End Select

    'Stop logging
    DisableLogging
    
    ' Clean up an aborted installation
    If (intExitCode = gintRET_FINISHEDSUCCESS) Then
        'Check to see if we need to reboot for mdac_typ
        Dim sRet As String
        If GetKeyValue(HK_LOCAL_MACHINE, sKEY, sValue, sRet) Then
            'We need to reboot
            'Warn the user before rebooting.  If they choose to reboot, do so, otherwise
            'Warn them again.
            If MsgBox(ResolveResString(resREBOOT), vbYesNo Or vbInformation, gstrTitle) = vbYes Then
                fNeedReboot = True
            Else
                fNeedReboot = False
                intExitCode = gintRET_FATAL
                MsgBox ResolveResString(resREBOOTNO), vbOKOnly Or vbExclamation, gstrTitle
            End If
        End If
    Else
        'Setup has been aborted for one reason or another
        If (gstrAppRemovalEXE <> "") Then
            Dim nErrorLevel As Integer
            Select Case intExitCode
                Case gintRET_FATAL
                    nErrorLevel = APPREMERR_FATAL
                Case gintRET_EXIT
                    nErrorLevel = APPREMERR_USERCANCEL
                Case gintRET_ABORT
                    nErrorLevel = APPREMERR_NONFATAL
                Case Else
                    nErrorLevel = APPREMERR_FATAL
                'End Case
            End Select
        
            '
            ' We don't want to log this message to sms because it is
            ' only a confirmation message.
            '
            gfDontLogSMS = True
            MsgFunc ResolveResString(resLOG_ABOUTTOREMOVEAPP), vbInformation Or vbOKOnly, gstrTitle
            
            Err = 0
            '
            ' Ready to run the installer.  Determine if this is a
            ' silent uninstall or not.
            '
            If gfSilent Then
                strSilent = gstrSilentLog
            Else
                strSilent = vbNullString
            End If
            
            Shell GetAppRemovalCmdLine(gstrAppRemovalEXE, gstrAppRemovalLog, strSilent, gfSMS, nErrorLevel, True), vbNormalFocus
            If Err Then
                MsgError Error$ & vbLf & vbLf & ResolveResString(resLOG_CANTRUNAPPREMOVER), vbExclamation Or vbOKOnly, gstrTitle
            End If

            'Since the app removal program will attempt to delete this program and all of our runtime
            'files, we should exit as soon as possible (otherwise the app remover will not be
            'able to remove these files)
        End If
        
        'Note: We do not delete the logfile if an error occurs.
        'The application removal EXE will do that if needed.
        
    End If
    
    Unload frmSetup1

    If gfSMS = True Then
        WriteMIF gstrMIFFile, gfSMSStatus, gstrSMSDescription
    End If

    'Try the reboot (if necessary)...
    If fNeedReboot Then RebootSystem
    'End the program
    If (intExitCode = gintRET_FINISHEDSUCCESS) Then
        ExitProcess iSUCCESS
    Else
        ExitProcess iFAIL
    End If
End Sub

'-----------------------------------------------------------
' FUNCTION: ProcessCommandLine
'
' Processes the command-line arguments
'
' OUT: Fills in the passed-in byref parameters as appropriate
'-----------------------------------------------------------
'
Sub ProcessCommandLine(ByVal strCommand As String, ByRef fSilent As Boolean, ByRef strSilentLog As String, ByRef fSMS As Boolean, ByRef strMIFFile As String, ByRef strSrcPath As String, ByRef strAppRemovalLog As String, ByRef strAppRemovalEXE As String)
    Dim fErr As Boolean
    Dim intAnchor As Integer
    
    Const strTemp$ = ""
    
    strSrcPath = ""
    strAppRemovalLog = ""
    
    strCommand = Trim$(strCommand)
    
    '
    ' First, check to see if this is supposed to be a silent
    ' install (/s on the command line followed by
    ' a log file name) and set global variables appropriately.
    '
    ' If you are designing a silent install, the /s
    ' command line parameter should be added to the setup.exe
    ' command.  It will automatically be passed to setup1 as the
    ' first parameter.
    '
    ' The filename that follows the /s parameter must
    ' include the full path name.
    '
    intAnchor = InStr(LCase(strCommand), gstrSwitchPrefix2 & gstrSILENTSWITCH)
    If intAnchor > 0 Then
        fSilent = True
        strCommand = Trim(Mid(strCommand, intAnchor + 2))
        strSilentLog = strExtractFilenameArg(strCommand, fErr)
        If fErr Then GoTo BadCommandLine
    Else
        fSilent = False
    End If
    fSMS = False
    
    '
    ' We expect to find the source directory,
    ' name/path of the logfile, and name/path
    ' of the app removal executable, separated only by
    ' spaces
    '
    strSrcPath = strExtractFilenameArg(strCommand, fErr)
    If fErr Then GoTo BadCommandLine
    
    strAppRemovalLog = strExtractFilenameArg(strCommand, fErr)
    If fErr Then GoTo BadCommandLine
    

    strAppRemovalEXE = strExtractFilenameArg(strCommand, fErr)
    If fErr Then GoTo BadCommandLine
        
    ' Both the app removal logfile and executable must exist
    If Not FileExists(strAppRemovalLog) Then
        GoTo BadAppRemovalLog
    End If
    
    If Not FileExists(strAppRemovalEXE) Then
        GoTo BadAppRemovalEXE
    End If
    
    ' Last check:  There should be nothing else on the command line
    strCommand = Trim$(strCommand)
    If strCommand <> "" Then
        GoTo BadCommandLine
    End If
    
    Exit Sub
    
BadAppRemovalLog:
    MsgError ResolveResString(resCANTFINDAPPREMOVALLOG, "|1", strAppRemovalLog), vbExclamation Or vbOKOnly, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
    
BadAppRemovalEXE:
    MsgError ResolveResString(resCANTFINDAPPREMOVALEXE, "|1", strAppRemovalEXE), vbExclamation Or vbOKOnly, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
    
BadCommandLine:
    MsgError ResolveResString(resBADCOMMANDLINE), vbExclamation Or vbOKOnly, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
End Sub

'-----------------------------------------------------------
' FUNCTION: GetDrivesAllocUnit
'
' Gets the minimum file size allocation unit for the
' specified drive
'
' IN: [strDrive] - Drive to get allocation unit for
'
' Returns: minimum allocation unit of drive, or -1 if
'          this value couldn't be determined
'-----------------------------------------------------------
'
Function GetDrivesAllocUnit(ByVal strDrive As String) As Long
    Dim strCurDrive As String
    Dim lAllocUnit As Long

    On Error Resume Next

    '
    'Save current drive
    '
    strCurDrive = Left$(CurDir$, 2)

    '
    'append a colon to the end of the drivespec if none supplied
    '
    If InStr(strDrive, gstrCOLON) = 0 Or Len(strDrive) > 2 Then
        strDrive = Left$(strDrive, 1) & gstrCOLON
    End If

    '
    'Change to the drive to determine the allocation unit for.  The AllocUnit()
    'API returns this value for the current drive only
    '
    ChDrive strDrive

    '
    'If there was an error accessing the specified drive, flag error return.
    'It is also possible for the AllocUnit() API to return -1 on other failure
    '
    If Err <> 0 Or (strDrive <> Left$(CurDir$, 2)) Then
        lAllocUnit = -1
    Else
        Dim lRet As Long
        Dim lBytes As Long, lSect As Long, lClust As Long, lTot As Long
        
        lRet = GetDiskFreeSpace(vbNullString, lSect, lBytes, lClust, lTot)
        lAllocUnit = lSect * lBytes
        If Err <> 0 Then
            lAllocUnit = -1
        End If
    End If

    If lAllocUnit = -1 Then
        MsgError Error$ & vbLf & vbLf & ResolveResString(resALLOCUNIT) & strDrive, vbExclamation, gstrTitle
        If gfSMS Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If

    GetDrivesAllocUnit = lAllocUnit

    '
    'Restore to original drive
    '
    ChDrive strCurDrive

    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: GetFileName
'
' Return the filename portion of a path
'
'-----------------------------------------------------------
'
Function GetFileName(ByVal strPath As String) As String
    Dim strFilename As String
    Dim iSep As Integer
    
    strFilename = strPath
    Do
        iSep = InStr(strFilename, gstrSEP_DIR)
        If iSep = 0 Then iSep = InStr(strFilename, gstrCOLON)
        If iSep = 0 Then
            GetFileName = strFilename
            Exit Function
        Else
            strFilename = Right(strFilename, Len(strFilename) - iSep)
        End If
    Loop
End Function

'-----------------------------------------------------------
' FUNCTION: GetFileSize
'
' Determine the size (in bytes) of the specified file
'
' IN: [strFileName] - name of file to get size of
'
' Returns: size of file in bytes, or -1 if an error occurs
'-----------------------------------------------------------
'
Function GetFileSize(strFilename As String) As Long
    On Error Resume Next

    GetFileSize = FileLen(strFilename)

    If Err > 0 Then
        GetFileSize = -1
        Err = 0
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: GetAppRemovalCmdLine
'
' Returns the correct command-line arguments (including
' path to the executable for use in calling the
' application removal executable)
'
' IN: [strAppRemovalEXE] - Full path/filename of the app removal EXE
'     [strAppRemovalLog] - Full path/filename of the app removal logfile
'     [strSilentLog] - Full path/filename of the file to log messages to when in silent mode.
'                       If this is an empty string then silent mode is turned off for uninstall.
'     [fSMS] - Boolean.  If True, we have been doing an SMS install and must tell the Uninstaller
'              to also do an SMS uninstall.  SMS is the Microsoft Systems Management Server.
'     [nErrorLevel] - Error level:
'                        APPREMERR_NONE - no error
'                        APPREMERR_FATAL - fatal error
'                        APPREMERR_NONFATAL - non-fatal error, user chose to abort
'                        APPREMERR_USERCANCEL - user chose to cancel (no error)
'     [fWaitForParent] - True if the application removal utility should wait
'                        for the parent (this process) to finish before starting
'                        to remove files.  Otherwise it may not be able to remove
'                        this process' executable file, depending upon timing.
'                        Defaults to False if not specified.
'-----------------------------------------------------------
'
Function GetAppRemovalCmdLine(ByVal strAppRemovalEXE As String, ByVal strAppRemovalLog, ByVal strSilentLog As String, ByVal fSMS As Boolean, ByVal nErrorLevel As Integer, Optional fWaitForParent As Boolean = False)
    Dim strEXE As String
    Dim strLog As String
    Dim strSilent As String
    Dim strErrLevel As String
    Dim strForce As String
    Dim strWait As String
    Dim strSMS As String

    strEXE = AddQuotesToFN(strAppRemovalEXE)
    strLog = "-n " & """" & GetLongPathName(strAppRemovalLog) & """"
    If gfSilent And strSilentLog <> vbNullString Then
        strSilent = "/s " & """" & strSilentLog & """"
    Else
        strSilent = vbNullString
    End If
    
    strSMS = IIf(fSMS, " /q ", vbNullString)
    
    strErrLevel = IIf(nErrorLevel <> APPREMERR_NONE, "-e " & Format(nErrorLevel), "")
    If nErrorLevel <> APPREMERR_NONE Then
        strForce = " -f"
    End If
    If fWaitForParent Then
        Dim curProcessId As Currency
        Dim Wrap As Currency
        Dim lProcessId As Long
        Dim cProcessId As Currency
        
        Wrap = 2 * (CCur(&H7FFFFFFF) + 1)

        'Always print as an unsigned long
        lProcessId = GetCurrentProcessId()
        cProcessId = lProcessId
        If cProcessId < 0 Then cProcessId = cProcessId + Wrap

        strWait = " -w " & str(cProcessId)
    End If
    
    GetAppRemovalCmdLine = strEXE & " " & strLog & " " & strSilent & " " & strSMS & strErrLevel & strForce & strWait
End Function

'-----------------------------------------------------------
' FUNCTION: IncrementRefCount
'
' Increments the reference count on a file in the registry
' so that it may properly be removed if the user chooses
' to remove this application.
'
' IN: [strFullPath] - FULL path/filename of the file
'     [fFileAlreadyExisted] - indicates whether the given
'                             file already existed on the
'                             hard drive
'-----------------------------------------------------------
'
Sub IncrementRefCount(ByVal strFullPath As String, ByVal fFileAlreadyExisted As Boolean)
    Dim strSharedDLLsKey As String
    strSharedDLLsKey = RegPathWinCurrentVersion() & "\SharedDLLs"
    
    'We must always use the LFN for the filename, so that we can uniquely
    'and accurately identify the file in the registry.
    strFullPath = GetLongPathName(strFullPath)
    
    'Get the current reference count for this file
    Dim fSuccess As Boolean
    Dim hKey As Long
    fSuccess = RegCreateKey(HKEY_LOCAL_MACHINE, strSharedDLLsKey, "", hKey)
    If fSuccess Then
        Dim lCurRefCount As Long
        If Not RegQueryRefCount(hKey, strFullPath, lCurRefCount) Then
            'No current reference count for this file
            If fFileAlreadyExisted Then
                'If there was no reference count, but the file was found
                'on the hard drive, it means one of two things:
                '  1) This file is shipped with the operating system
                '  2) This file was installed by an older setup program
                '     that does not do reference counting
                'In either case, the correct conservative thing to do
                'is assume that the file is needed by some application,
                'which means it should have a reference count of at
                'least 1.  This way, our application removal program
                'will not delete this file.
                lCurRefCount = 1

            Else
                lCurRefCount = 0
            End If
        End If
        
        'Increment the count in the registry
        fSuccess = RegSetNumericValue(hKey, strFullPath, lCurRefCount + 1, False)
        If Not fSuccess Then
            GoTo DoErr
        End If
        RegCloseKey hKey
    Else
        GoTo DoErr
    End If
    
    Exit Sub
    
DoErr:
    'An error message should have already been shown to the user
    Exit Sub
End Sub

'-----------------------------------------------------------
' FUNCTION: InitDiskInfo
'
' Called before calculating disk space to initialize
' values used/determined when calculating disk space
' required.
'-----------------------------------------------------------
'
Sub InitDiskInfo()
    '
    'Initialize "table" of drives used and disk space array
    '
    gstrDrivesUsed = vbNullString
    Erase gsDiskSpace

    mlTotalToCopy = 0

    '
    'Get drive/directory for temporary files
    '
    mstrConcatDrive = UCase$(Environ$(gstrTMP_DIR))
    If mstrConcatDrive = vbNullString Then
        mstrConcatDrive = UCase$(Environ$(gstrTEMP_DIR))
    End If
    AddDirSep mstrConcatDrive

    If mstrConcatDrive <> vbNullString Then
        If CheckDrive(mstrConcatDrive, ResolveResString(resTEMPDRIVE)) = False Then
            mstrConcatDrive = vbNullString
        Else
            '
            'If we found a temp drive and the drive is "ready", add it to the
            'table of drives used
            '
            gstrDrivesUsed = Left$(mstrConcatDrive, 1)
            ReDim Preserve gsDiskSpace(1)
            gsDiskSpace(1).lAvail = GetDiskSpaceFree(mstrConcatDrive)
            gsDiskSpace(1).lMinAlloc = GetDrivesAllocUnit(mstrConcatDrive)
        End If
    End If
End Sub

'-----------------------------------------------------------
' FUNCTION: IsDisplayNameUnique
'
' Determines whether a given display name for registering
'   the application removal executable is unique or not.  This
'   display name is the title which is presented to the
'   user in Windows 95's control panel Add/Remove Programs
'   applet.
'
' IN: [hkeyAppRemoval] - open key to the path in the registry
'                       containing application removal entries
'     [strDisplayName] - the display name to test for uniqueness
'
' Returns: True if the given display name is already in use,
'          False if otherwise
'-----------------------------------------------------------
'
Function IsDisplayNameUnique(ByVal hkeyAppRemoval As Long, ByVal strDisplayName As String) As Boolean
    Dim lIdx As Long
    Dim strSubkey As String
    Dim strDisplayNameExisting As String
    Const strKEY_DISPLAYNAME$ = "DisplayName"
    
    IsDisplayNameUnique = True
    
    lIdx = 0
    Do
        Select Case RegEnumKey(hkeyAppRemoval, lIdx, strSubkey)
            Case ERROR_NO_MORE_ITEMS
                'No more keys - must be unique
                Exit Do
            Case ERROR_SUCCESS
                'We have a key to some application removal program.  Compare its
                '  display name with ours
                Dim hkeyExisting As Long
                
                If RegOpenKey(hkeyAppRemoval, strSubkey, hkeyExisting) Then
                    If RegQueryStringValue(hkeyExisting, strKEY_DISPLAYNAME, strDisplayNameExisting) Then
                        If strDisplayNameExisting = strDisplayName Then
                            'There is a match to an existing display name
                            IsDisplayNameUnique = False
                            RegCloseKey hkeyExisting
                            Exit Do
                        End If
                    End If
                    RegCloseKey hkeyExisting
                End If
            Case Else
                'Error, we must assume it's unique.  An error will probably
                '  occur later when trying to add to the registry
                Exit Do
            'End Case
        End Select
        lIdx = lIdx + 1
    Loop
End Function

'-----------------------------------------------------------
' FUNCTION: IsNewerVer
'
' Compares two file version structures and determines
' whether the source file version is newer (greater) than
' the destination file version.  This is used to determine
' whether a file needs to be installed or not
'
' IN: [sSrcVer] - source file version information
'     [sDestVer] - dest file version information
'
' Returns: True if source file is newer than dest file,
'          False if otherwise
'-----------------------------------------------------------
'
Function IsNewerVer(sSrcVer As VERINFO, sDestVer As VERINFO) As Integer
    IsNewerVer = False

    With sSrcVer
     If .FileVerPart1 > sDestVer.FileVerPart1 Then GoTo INVNewer
     If .FileVerPart1 < sDestVer.FileVerPart1 Then GoTo INVOlder
     
     If .FileVerPart2 > sDestVer.FileVerPart2 Then GoTo INVNewer
     If .FileVerPart2 < sDestVer.FileVerPart2 Then GoTo INVOlder
    
     If .FileVerPart3 > sDestVer.FileVerPart3 Then GoTo INVNewer
     If .FileVerPart3 < sDestVer.FileVerPart3 Then GoTo INVOlder
     
     If .FileVerPart4 > sDestVer.FileVerPart4 Then GoTo INVNewer
    End With
    With sSrcVer
        If (.FileVerPart1 = sDestVer.FileVerPart1) And _
            (.FileVerPart2 = sDestVer.FileVerPart2) And _
            (.FileVerPart3 = sDestVer.FileVerPart3) And _
            (.FileVerPart4 = sDestVer.FileVerPart4) Then GoTo INVNewer
    End With
    GoTo INVOlder

INVNewer:
    IsNewerVer = True
INVOlder:
End Function

'-----------------------------------------------------------
' FUNCTION: IsValidDestDir
'
' Determines whether or not the destination directory
' specifed in the "DefaultDir" key of the [Setup] section
' in SETUP.LST or a destination dir entered by the user
' is not a subdirectory of the source directory.
'
' Notes: [gstrSrcPath] - points to the source directory
'        [strDestDir] - points to the dest directory
'
' Returns: True if dest dir is a valid location, False
'          otherwise
'-----------------------------------------------------------
'
Function IsValidDestDir(strDestDir As String) As Integer
    Dim strMsg As String

    '
    ' Both of these paths, strDestDir and gstrSrcPath, are *always*
    ' in the format 'X:\' or 'X:\DIRNAME\'.
    '
    If InStr(strDestDir, gstrSrcPath) > 0 Then
        IsValidDestDir = False
        strMsg = ResolveResString(resDIRSPECIFIED) & vbLf & strDestDir & vbLf & ResolveResString(resSAMEASSRC)
        MsgFunc strMsg, vbOKOnly Or vbExclamation, gstrTitle
    Else
        IsValidDestDir = True
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: MakePath
'
' Creates the specified directory path
'
' IN: [strDirName] - name of the dir path to make
'     [fAllowIgnore] - whether or not to allow the user to
'                      ignore any encountered errors.  If
'                      false, the function only returns
'                      if successful.  If missing, this
'                      defaults to True.
'
' Returns: True if successful, False if error and the user
'          chose to ignore.  (The function does not return
'          if the user selects ABORT/CANCEL on an error.)
'-----------------------------------------------------------
'
Public Function MakePath(ByVal strDir As String, Optional ByVal fAllowIgnore As Boolean = True) As Boolean
    Do
        If MakePathAux(strDir) Then
            MakePath = True
            Exit Function
        Else
            Dim strMsg As String
            Dim iRet As Integer
            
            strMsg = ResolveResString(resMAKEDIR) & vbLf & strDir
            iRet = MsgError(strMsg, IIf(fAllowIgnore, vbAbortRetryIgnore, vbRetryCancel) Or vbExclamation Or vbDefaultButton2, gstrSETMSG)
            '
            ' if we are running silent then we
            ' can't continue.  Previous MsgError
            ' took care of write silent log entry.
            '
            If gfNoUserInput = True Then
                ExitSetup frmCopy, gintRET_FATAL
            End If
            
            Select Case iRet
                Case vbAbort, vbCancel
                    ExitSetup frmCopy, gintRET_ABORT
                Case vbIgnore
                    MakePath = False
                    Exit Function
                Case vbRetry
                'End Case
            End Select
        End If
    Loop
End Function

'----------------------------------------------------------
' SUB: MoveAppRemovalFiles
'
' Moves the app removal logfile to the application directory,
' and registers the app removal executable with the operating
' system.
'----------------------------------------------------------
Sub MoveAppRemovalFiles(ByVal strGroupName As String)
    Dim strNewAppRemovalLogName As String
    Dim lCount As Long
    Dim sCab As String
    Dim sTemp As String
    
    lCount = 0
    'Get rid of the cab file in the windows dir (if it exists).
    Do
        If gintCabs = 1 Then
            sCab = gstrWinDir
            AddDirSep sCab
            sCab = sCab & BaseName(GetShortPathName(gsCABNAME))
            If FileExists(sCab) Then Kill sCab
            Exit Do
        End If
        lCount = lCount + 1
        sCab = gstrWinDir
        AddDirSep sCab
        sTemp = Left(gsCABNAME, Len(gsCABNAME) - 5) & CStr(lCount) & gstrSEP_EXT & gsINI_CABNAME
        sCab = sCab & BaseName(sTemp)
        If FileExists(sCab) Then
            Kill sCab
        Else
            Exit Do
        End If
    Loop
    'Get rid of the temp dir
    'Bug fix for #6-34583
    KillTempFolder
    'Find a unique name for the app removal logfile in the
    'application directory
    
    '...First try the default extension
    strNewAppRemovalLogName = gstrDestDir & mstrFILE_APPREMOVALLOGBASE & mstrFILE_APPREMOVALLOGEXT
    If FileExists(strNewAppRemovalLogName) Then
        '...Next try incrementing integral extensions
        Dim iExt As Integer
        Do
            If iExt > 999 Then
                GoTo CopyErr
            End If
            

            strNewAppRemovalLogName = gstrDestDir & mstrFILE_APPREMOVALLOGBASE & gstrSEP_EXT & Format(iExt, "000")
            If Not FileExists(strNewAppRemovalLogName) Then
                Exit Do 'Unique name was found
            Else
                iExt = iExt + 1
            End If
        Loop
    End If
    
    
    
    On Error GoTo CopyErr
    FileCopy gstrAppRemovalLog, strNewAppRemovalLogName
    
    'Now we need to start logging in the new logfile, so that the
    'creation of the application removal icon under NT gets logged.
    EnableLogging strNewAppRemovalLogName
    
    On Error GoTo 0
    If Not RegisterAppRemovalEXE(gstrAppRemovalEXE, strNewAppRemovalLogName, strGroupName) Then
        If TreatAsWin95() Then
            MsgError ResolveResString(resCANTREGISTERAPPREMOVER), vbExclamation Or vbOKOnly, gstrTitle
        Else
            MsgError ResolveResString(resCANTCREATEAPPREMOVALICON), vbExclamation Or vbOKOnly, gstrTitle
        End If
        ExitSetup frmSetup1, gintRET_FATAL
    End If
    
    'Now we can delete the original logfile, since we no longer have a reference
    'to it, and start using the new logfile
    On Error Resume Next
    Kill gstrAppRemovalLog
    
    'This temporary app removal logfile should no longer be used
    gstrAppRemovalLog = strNewAppRemovalLogName
    gfAppRemovalFilesMoved = True
    
    Exit Sub
    
CleanUpOnErr:
    On Error Resume Next
    Kill strNewAppRemovalLogName
    On Error GoTo 0
    MsgError ResolveResString(resCANTCOPYLOG, "|1", gstrAppRemovalLog), vbExclamation Or vbOKOnly, gstrTitle
    ExitSetup Screen.ActiveForm, gintRET_FATAL
    
CopyErr:
    Resume CleanUpOnErr
End Sub
'-----------------------------------------------------------
' SUB: KillTempFolder
' BUG FIX #6-34583
'
' Deletes the temporary files stored in the temp folder
'
Private Sub KillTempFolder()

    Const sWILD As String = "*.*"
    Dim sFile As String
    
    sFile = Dir(gsTEMPDIR & sWILD)
    While sFile <> vbNullString
        SetAttr gsTEMPDIR & sFile, vbNormal
        Kill gsTEMPDIR & sFile
        sFile = Dir
    Wend
    RmDir gsTEMPDIR
End Sub

'-----------------------------------------------------------
' SUB: ParseDateTime
'
' Same as CDate with a string argument, except that it
' ignores the current localization settings.  This is
' important because SETUP.LST always uses the same
' format for dates.
'
' IN: [strDate] - string representing the date in
'                 the format mm/dd/yy or mm/dd/yyyy
' OUT: The date which strDate represents
'-----------------------------------------------------------
'
Function ParseDateTime(ByVal strDateTime As String) As Date
Dim Var As Variant
    Var = strDateTime
    If 0 = VariantChangeTypeEx(VarPtr(Var), VarPtr(Var), &H409, 0, vbDate) Then
        ParseDateTime = Var
    Else
        'Raise same error as CDate
        Err.Raise 13
    End If
End Function

'-----------------------------------------------------------
' SUB: PerformDDE
'
' Performs a Program Manager DDE operation as specified
' by the intDDE flag and the passed in parameters.
' Possible operations are:
'
'   mintDDE_ITEMADD:  Add an icon to the active group
'   mintDDE_GRPADD:   Create a program manager group
'
' IN: [frm] - form containing a label named 'lblDDE'
'     [strGroup] - name of group to create or insert icon
'     [strTitle] - title of icon or group
'     [strCmd] - command line for icon/item to add
'     [intDDE] - ProgMan DDE action to perform
'-----------------------------------------------------------
'
Sub PerformDDE(frm As Form, ByVal strGroup As String, ByVal strCmd As String, ByVal strTitle As String, ByVal intDDE As Integer, ByVal fLog As Boolean)
    Const strCOMMA$ = ","
    Const strRESTORE$ = ", 1)]"
    Const strACTIVATE$ = ", 5)]"
    Const strENDCMD$ = ")]"
    Const strSHOWGRP$ = "[ShowGroup("
    Const strADDGRP$ = "[CreateGroup("
    Const strREPLITEM$ = "[ReplaceItem("
    Const strADDITEM$ = "[AddItem("

    Dim intIdx As Integer        'loop variable

    SetMousePtr vbHourglass

    '
    'Initialize for DDE Conversation with Windows Program Manager in
    'manual mode (.LinkMode = 2) where destination control is not auto-
    'matically updated.  Set DDE timeout for 10 seconds.  The loop around
    'DoEvents() is to allow time for the DDE Execute to be processsed.
    '

    Dim intRetry As Integer
    For intRetry = 1 To 20
        On Error Resume Next
        frm.lblDDE.LinkTopic = "PROGMAN|PROGMAN"
        If Err = 0 Then
            Exit For
        End If
        DoEvents
    Next intRetry
        
    frm.lblDDE.LinkMode = 2
    For intIdx = 1 To 10
      DoEvents
    Next
    frm.lblDDE.LinkTimeout = 100

    On Error Resume Next

    If Err = 0 Then
        Select Case intDDE
            Case mintDDE_ITEMADD
                '
                ' The item will be created in the group titled strGroup
                '
                ' Write the action to the logfile
                '
                If fLog Then
                    NewAction gstrKEY_PROGMANITEM, """" & strUnQuoteString(strGroup) & """" & ", " & """" & strUnQuoteString(strTitle) & """"
                End If
                '
                ' Force the group strGroup to be the active group.  Additem only
                ' puts icons in the active group.
                '
                #If 0 Then
                    frm.lblDDE.LinkExecute strSHOWGRP & strGroup & strACTIVATE
                #Else
                    ' strShowGRP doesn't seem to work if ProgMan is minimized.
                    '  : strADDGRP does the trick fine, though, and it doesn't matter if it already exists.
                    frm.lblDDE.LinkExecute strADDGRP & strGroup & strENDCMD
                #End If
                frm.lblDDE.LinkExecute strREPLITEM & strTitle & strENDCMD
                Err = 0
                frm.lblDDE.LinkExecute strADDITEM & strCmd & strCOMMA & strTitle & String$(3, strCOMMA) & strENDCMD
            Case mintDDE_GRPADD
                '
                ' Write the action to the logfile
                '
                If fLog Then
                    NewAction gstrKEY_PROGMANGROUP, """" & strUnQuoteString(strGroup) & """"
                End If
                frm.lblDDE.LinkExecute strADDGRP & strGroup & strENDCMD
                frm.lblDDE.LinkExecute strSHOWGRP & strGroup & strRESTORE
            'End Case
        End Select
    End If

    
    '
    'Disconnect DDE Link
    '

    frm.lblDDE.LinkMode = 0
    frm.lblDDE.LinkTopic = ""


    SetMousePtr gintMOUSE_DEFAULT

    If fLog Then
        CommitAction
    End If
    
    
    Err = 0
End Sub

'-----------------------------------------------------------
' SUB: PromptForNextDisk
'
' If the source media is removable or a network connection,
' prompts the user to insert the specified disk number
' containing the filename which is used to determine that
' the correct disk is inserted.
'
' IN: [intDiskNum] - disk number to insert
'     [strDetectFile] - file to search for to ensure that
'                       the correct disk was inserted
'
' Notes: [gstrSrcPath] - used to identify the source drive
'-----------------------------------------------------------
'
Sub PromptForNextDisk(ByVal intDiskNum As Integer, ByVal strDetectFile As String)
    Static intDrvType As Integer

    Dim intRC As Integer
    Dim strMsg As String
    Dim strDrive As String
    Dim strMultDirBaseName As String
    Dim strDetectPath As String

    On Error Resume Next

    strMultDirBaseName = ResolveResString(resCOMMON_MULTDIRBASENAME)
    '
    'Get source drive and, if we haven't yet determined it, get the
    'source drive type
    '
    
    strDrive = Left$(gstrSrcPath, 2)
    If intDrvType = 0 Then
        If IsUNCName(strDrive) Then
            intDrvType = intDRIVE_REMOTE
            strDrive = gstrSrcPath
        Else
            intDrvType = GetDriveType(Asc(strDrive) - 65)
        End If
    End If

    While SrcFileMissing(gstrSrcPath, strDetectFile, intDiskNum) = True
        Select Case intDrvType
            Case 0, intDRIVE_REMOVABLE, intDRIVE_CDROM
                strMsg = ResolveResString(resINSERT) & vbLf & ResolveResString(resDISK) & Format$(intDiskNum)
                strMsg = strMsg & ResolveResString(resINTO) & strDrive
            Case intDRIVE_REMOTE
                strMsg = ResolveResString(resCHKCONNECT) & strDrive
            Case intDRIVE_FIXED
                If DirExists(gstrSrcPath & strMultDirBaseName & Format(intDiskNum)) = True Then
                    strDetectPath = gstrSrcPath & strMultDirBaseName & Format(intDiskNum)
                Else
                    strDetectPath = gstrSrcPath
                End If
                strMsg = ResolveResString(resCOMMON_CANTFINDSRCFILE, "|1", strDetectPath & gstrSEP_DIR & strDetectFile)
            'End Case
        End Select

        Beep
        intRC = MsgFunc(strMsg, vbOKCancel Or vbExclamation, gstrSETMSG)
        '
        ' We should always fail if in silent or sms mode.
        '
        If intRC = vbCancel Or gfNoUserInput Then
            ExitSetup frmCopy, gintRET_EXIT
        End If
    Wend

    gintCurrentDisk = intDiskNum
End Sub
Function SrcFileMissing(ByVal strSrcDir As String, ByVal strSrcFile As String, ByVal intDiskNum As Integer) As Boolean
'-----------------------------------------------------------
' FUNCTION: SrcFileMissing
'
' Tries to locate the file strSrcFile by first looking
' in the strSrcDir directory, then in the DISK(x+1)
' directory if it exists.
'
' IN: [strSrcDir] - Directory/Path where file should be.
'     [strSrcFile] - File we are looking for.
'     [intDiskNum] - Disk number we are expecting file
'                    to be on.
'
' Returns: True if file not found; otherwise, false
'-----------------------------------------------------------
    Dim fFound As Boolean
    Dim strMultDirBaseName As String
    
    fFound = False
    
    AddDirSep strSrcDir
    '
    ' First check to see if it's in the main src directory.
    ' This would happen if someone copied the contents of
    ' all the floppy disks to a single directory on the
    ' hard drive.  We should allow this to work.
    '
    ' This test would also let us know if the user inserted
    ' the wrong floppy disk or if a network connection is
    ' unavailable.
    '
    If FileExists(strSrcDir & strSrcFile) = True Then
        fFound = True
        GoTo doneSFM
    End If
    '
    ' Next try the DISK(x) subdirectory of the main src
    ' directory.  This would happen if the floppy disks
    ' were copied into directories named DISK1, DISK2,
    ' DISK3,..., DISKN, etc.
    '
    strMultDirBaseName = ResolveResString(resCOMMON_MULTDIRBASENAME)
    If FileExists(strSrcDir & ".." & gstrSEP_DIR & strMultDirBaseName & Format(intDiskNum) & gstrSEP_DIR & strSrcFile) = True Then
        fFound = True
        GoTo doneSFM
    End If
    
doneSFM:
    SrcFileMissing = Not fFound
End Function
'-----------------------------------------------------------
' FUNCTION: ReadIniFile
'
' Reads a value from the specified section/key of the
' specified .INI file
'
' IN: [strIniFile] - name of .INI file to read
'     [strSection] - section where key is found
'     [strKey] - name of key to get the value of
'
' Returns: non-zero terminated value of .INI file key
'-----------------------------------------------------------
'
Function ReadIniFile(ByVal strIniFile As String, ByVal strsection As String, ByVal strKey As String) As String
    Dim strBuffer As String
    Dim intPos As Integer

    '
    'If successful read of .INI file, strip any trailing zero returned by the Windows API GetPrivateProfileString
    '
    strBuffer = Space$(gintMAX_SIZE)
    
    If GetPrivateProfileString(strsection, strKey, vbNullString, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        ReadIniFile = RTrim$(StripTerminator(strBuffer))
    Else
        ReadIniFile = vbNullString
    End If
End Function

'-----------------------------------------------------------
' SUB: ReadSetupFileLine
'
' Reads the requested 'FileX=' key from the specified
' section of the setup information file (SETUP.LST).
'
' IN: [strSection] - name of section to read from SETUP.LST,
'                    Ex: "Files"
'     [intFileNum] - file number index to read
'
' OUT: [sFile] - FILEINFO Type variable that, after parsing,
'                holds the information for the file
'                described.
'
' Returns: True if the requested info was successfully read,
'          False otherwise
'
' Notes: Lines in the setup information file have the
'        following format:
'
'        #,[SPLIT],SrcName,DestName,DestDir,Register,
'        Date,Size,Version
'
'        [#] - disk number where this file is located
'        [SPLIT] - optional, determines whether this is
'                  an extent of a split file.  The last
'                  extent does not specify this key
'        [SrcName] - filename on the installation media
'        [DestName] - file name to use when copied
'
'        (For split files, the following info is required only
'        for the *first* extent)
'
'        [DestDir] - dirname or macro specifying destdir
'        [Register] - reginfo file name or macro specifying
'                     file registration action
'        [Date] - date of the source file
'        [Size] - size of the source file
'        [Version] - optional, version number string
'        [Reserved] - Must be empty, else error!
'        [ProgIcon] - Caption for icon, if there is one.
'        [ProgCmdLine] - Command line for icon, if there is one.
'-----------------------------------------------------------
'
Function ReadSetupFileLine(ByVal strsection As String, ByVal intFileNum As Integer, sFile As FILEINFO) As Integer
    Static strSplitName As String
    Const CompareBinary = 0

    Dim strLine As String
    Dim strMsg As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim fDone As Integer
    Dim fErr As Boolean
    Dim strVersion As String
    Dim strFilename As String

    ReadSetupFileLine = False

    sFile.fSystem = False
    sFile.fShared = False
    
    '
    ' Read the requested line, if unable to read it (strLine = vbnullstring) then exit
    '
    strLine = ReadIniFile(gstrSetupInfoFile, strsection, gstrINI_FILE & Format$(intFileNum))
    If strLine = vbNullString Then
        Exit Function
    End If

    '
    'source file name, ensure it's not a UNC name
    '
    intAnchor = 1
    sFile.strSrcName = strExtractFilenameItem(strLine, intAnchor, fErr)
    If fErr Then GoTo RSFLError
    If IsUNCName(sFile.strSrcName) = True Then GoTo RSFLError
    intAnchor = intAnchor + 1 'Skip past the comma
    
    '
    'dest file name, ensure it's not a UNC name
    '
    If Left(sFile.strSrcName, 1) = gstrSEP_AMPERSAND Then
        sFile.strDestName = Right(sFile.strSrcName, Len(sFile.strSrcName) - 1)
    Else
        sFile.strDestName = sFile.strSrcName
    End If
    strFilename = GetFileName(sFile.strDestName)

    '
    'parse and resolve destination directory
    '
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    If intOffset > 0 Then
        Dim strInitialDestDir As String
        strInitialDestDir = Mid$(strLine, intAnchor, intOffset - intAnchor)
        If InStr(strInitialDestDir, gstrWINSYSDESTSYSFILE) Then
            sFile.fSystem = True
        End If
        If InStr(strInitialDestDir, gstrDAODEST) Then
            '
            ' Special case for DAO destinations.  If there
            ' are any DAO files, we need to add special
            ' DAO reg info later.  gfRegDAO tells us to do that.
            '
            gfRegDAO = True
        End If
        sFile.strDestDir = ResolveDestDir(strInitialDestDir)
        If sFile.strDestDir <> "?" Then
            sFile.strDestDir = ResolveDir(sFile.strDestDir, False, False)
            If sFile.strDestDir = vbNullString Or IsUNCName(sFile.strDestDir) Then
                GoTo RSFLError
            End If
        End If
    Else
        GoTo RSFLError
    End If

    '
    'file registration information
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    If intOffset > 0 Then
        sFile.strRegister = Mid$(strLine, intAnchor, intOffset - intAnchor)
    Else
        GoTo RSFLError
    End If

    '
    'Extract file share type
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    sFile.fShared = False
    If intOffset > 0 Then
        Dim strShareType As String
        strShareType = Mid$(strLine, intAnchor, intOffset - intAnchor)
        Select Case strShareType
            Case mstrPRIVATEFILE
                sFile.fShared = False
            Case mstrSHAREDFILE
                If sFile.fSystem Then
                    'A file cannot be both system and shared
                    GoTo RSFLError
                End If
                
                sFile.fShared = True
            Case Else
                GoTo RSFLError
            'End Case
        End Select
    End If
    
    '
    'Extract file date and convert to a date variant
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    If intOffset > 0 Then
        On Error GoTo RSFLError
        sFile.varDate = ParseDateTime(Mid$(strLine, intAnchor, intOffset - intAnchor))
        On Error GoTo 0
    End If

    '
    'Get file size
    '
    intAnchor = intOffset + 1
    intOffset = intGetNextFldOffset(intAnchor, strLine, gstrCOMMA, CompareBinary)
    If intOffset > 0 Then
        sFile.lFileSize = Val(Mid$(strLine, intAnchor, intOffset - intAnchor))
    Else
        GoTo RSFLError
    End If

    '
    ' Get the version number, otherwise flag that there is no version info
    '
    intAnchor = intOffset + 1
    If intOffset > 0 Then
        strVersion = Trim(Right$(strLine, Len(strLine) - intOffset))
        If strVersion = "" Then
            sFile.sVerInfo.FileVerPart1 = gintNOVERINFO
        Else
            PackVerInfo strVersion, sFile.sVerInfo
        End If
    Else
        GoTo RSFLError
    End If
    
RSFLDone:
    ReadSetupFileLine = True
    Exit Function

RSFLError:
    strMsg = gstrSetupInfoFile & vbLf & vbLf & ResolveResString(resINVLINE) & vbLf & vbLf
    strMsg = strMsg & ResolveResString(resSECTNAME) & strsection & vbLf & strLine
    MsgError strMsg, vbCritical, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
End Function

'-----------------------------------------------------------
' SUB: ReadSetupRemoteLine
'
' Reads the requested 'RemoteX=' key from the specified
' section of the setup information file (SETUP.LST).
'
' IN: [strSection] - name of section to read from SETUP.LST,
'                    Ex: "Files"
'     [intFileNum] - remote number index to read
'
' OUT: [rInfo] - REGINFO Type variable that, after parsing,
'                holds the information for the line
'                described.
'
' Returns: True if the requested info was successfully read,
'          False otherwise
'
' Notes: Remote server lines in the setup information file
'        have the following format:
'
'        address,protocol,authentication-level
'
'        [address] - network address of the server, if known
'        [protocol] - network protocol name, if known
'        [authentication level] - authentication level (or 0 for default)
'-----------------------------------------------------------
'
Function ReadSetupRemoteLine(ByVal strsection As String, ByVal intFileNum As Integer, rInfo As REGINFO) As Integer
    Dim strLine As String
    Dim strMsg As String
    Dim intAnchor As Integer
    Dim intOffset As Integer
    Dim fErr As Boolean

    ReadSetupRemoteLine = False

    '
    'Read the requested line, if unable to read it (strLine = vbnullstring) then exit
    '
    strLine = ReadIniFile(gstrSetupInfoFile, strsection, gstrINI_REMOTE & Format$(intFileNum))
    If strLine = vbNullString Then
        Exit Function
    End If

    '
    'Get the network address
    '
    intAnchor = 1
    fErr = False
    If Mid$(strLine, intAnchor, 1) = gstrCOMMA Then
        rInfo.strNetworkAddress = ""
    Else
        rInfo.strNetworkAddress = strExtractFilenameItem(strLine, intAnchor, fErr)
    End If
    If fErr Then GoTo RSRLError
    intAnchor = intAnchor + 1 'Skip past the comma

    '
    'Get the network protocol
    '
    If Mid$(strLine, intAnchor, 1) = gstrCOMMA Then
        rInfo.strNetworkProtocol = ""
    Else
        rInfo.strNetworkProtocol = strExtractFilenameItem(strLine, intAnchor, fErr)
    End If
    If fErr Then GoTo RSRLError
    intAnchor = intAnchor + 1 'Skip past the comma

    '
    'Get the authentication level (must be a single digit
    '  in the range 0..6)
    '
    Const intMaxAuthentication = 6
    Dim strAuthentication As String
    
    strAuthentication = Mid$(strLine, intAnchor, 1)
    If Len(strAuthentication) <> 1 Then GoTo RSRLError
    If (Asc(strAuthentication) < Asc("0")) Or (Asc(strAuthentication) > Asc("9")) Then GoTo RSRLError
    rInfo.intAuthentication = Val(strAuthentication)
    If rInfo.intAuthentication > intMaxAuthentication Then GoTo RSRLError
    '
    ' Is this dcom or remote automation?
    '
    intAnchor = InStr(intAnchor + 1, strLine, gstrCOMMA)
    If intAnchor > 0 Then
        rInfo.fDCOM = (Trim(Mid$(strLine, intAnchor + 1)) = gstrDCOM)
    End If
    
    ReadSetupRemoteLine = True
    Exit Function

RSRLError:
    strMsg = gstrSetupInfoFile & vbLf & vbLf & ResolveResString(resINVLINE) & vbLf & vbLf
    strMsg = strMsg & ResolveResString(resSECTNAME) & strsection & vbLf & strLine
    MsgError strMsg, vbCritical, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
End Function

'-----------------------------------------------------------
' FUNCTION: RegCloseKey
'
' Closes an open registry key.
'
' Returns: True on success, else False.
'-----------------------------------------------------------
'
Function RegCloseKey(ByVal hKey As Long) As Boolean
    Dim lResult As Long
    
    On Error GoTo 0
    lResult = OSRegCloseKey(hKey)
    RegCloseKey = (lResult = ERROR_SUCCESS)
End Function

'-----------------------------------------------------------
' FUNCTION: RegCreateKey
'
' Opens (creates if already exists) a key in the system registry.
'
' IN: [hkey] - The HKEY parent.
'     [lpszSubKeyPermanent] - The first part of the subkey of
'         'hkey' that will be created or opened.  The application
'         removal utility (32-bit only) will never delete any part
'         of this subkey.  May NOT be an empty string ("").
'     [lpszSubKeyRemovable] - The subkey of hkey\lpszSubKeyPermanent
'         that will be created or opened.  If the application is
'         removed (32-bit only), then this entire subtree will be
'         deleted, if it is empty at the time of application removal.
'         If this parameter is an empty string (""), then the entry
'         will not be logged.
'
' OUT: [phkResult] - The HKEY of the newly-created or -opened key.
'
' Returns: True if the key was created/opened OK, False otherwise
'   Upon success, phkResult is set to the handle of the key.
'
'-----------------------------------------------------------
Function RegCreateKey(ByVal hKey As Long, ByVal lpszSubKeyPermanent As String, ByVal lpszSubKeyRemovable As String, phkResult As Long) As Boolean
    Dim lResult As Long
    Dim strHkey As String
    Dim fLog As Boolean
    Dim strSubKeyFull As String

    On Error GoTo 0

    If lpszSubKeyPermanent = "" Then
        RegCreateKey = False 'Error: lpszSubKeyPermanent must not = ""
        Exit Function
    End If
    
    If Left$(lpszSubKeyRemovable, 1) = "\" Then
        lpszSubKeyRemovable = Mid$(lpszSubKeyRemovable, 2)
    End If

    If lpszSubKeyRemovable = "" Then
        fLog = False
    Else
        fLog = True
    End If
    
    If lpszSubKeyRemovable <> "" Then
        strSubKeyFull = lpszSubKeyPermanent & "\" & lpszSubKeyRemovable
    Else
        strSubKeyFull = lpszSubKeyPermanent
    End If
    strHkey = strGetHKEYString(hKey)

    If fLog Then
        NewAction _
          gstrKEY_REGKEY, _
          """" & strHkey & "\" & lpszSubKeyPermanent & """" _
            & ", " & """" & lpszSubKeyRemovable & """"
    End If

    lResult = OSRegCreateKey(hKey, strSubKeyFull, phkResult)
    If lResult = ERROR_SUCCESS Then
        RegCreateKey = True
        If fLog Then
            CommitAction
        End If
        AddHkeyToCache phkResult, strHkey & "\" & strSubKeyFull
    Else
        RegCreateKey = False
        MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
        If fLog Then
            AbortAction
        End If
        If gfNoUserInput Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: RegDeleteKey
'
' Deletes an existing key in the system registry.
'
' Returns: True on success, False otherwise
'-----------------------------------------------------------
'
Function RegDeleteKey(ByVal hKey As Long, ByVal lpszSubKey As String) As Boolean
    Dim lResult As Long
    
    On Error GoTo 0
    lResult = OSRegDeleteKey(hKey, lpszSubKey)
    RegDeleteKey = (lResult = ERROR_SUCCESS)
End Function

'-----------------------------------------------------------
' SUB: RegEdit
'
' Calls REGEDIT to add the information in the specifed file
' to the system registry.  If your .REG file requires path
' information based upon the destination directory given by
' the user, then you will need to write and call a .REG fixup
' routine before performing the registration below.
'
' WARNING: Use of this functionality under Win32 is not recommended,
' WARNING: because the application removal utility does not support
' WARNING: undoing changes that occur as a result of calling
' WARNING: REGEDIT on an arbitrary .REG file.
' WARNING: Instead, it is recommended that you use the RegCreateKey(),
' WARNING: RegOpenKey(), RegSetStringValue(), etc. functions in
' WARNING: this module instead.  These make entries to the
' WARNING: application removal logfile, thus enabling application
' WARNING: removal to undo such changes.
'
' IN: [strRegFile] - name of file containing reg. info
'-----------------------------------------------------------
'
Sub RegEdit(ByVal strRegFile As String)
    Const strREGEDIT$ = "REGEDIT /S "

    Dim fShellOK As Integer

    On Error Resume Next

    If FileExists(strRegFile) = True Then
        'Because regedit is a 16-bit application, it does not accept
        'double quotes around the filename.  Thus, if strRegFile
        'contains spaces, the only way to get this to work is to pass
        'regedit the short pathname version of the filename.
        strRegFile = GetShortPathName(strRegFile)
        
        fShellOK = SyncShell(strREGEDIT & strRegFile, INFINITE, , True)
        frmSetup1.Refresh
    Else
        MsgError ResolveResString(resCANTFINDREGFILE, "|1", strRegFile), vbExclamation Or vbOKOnly, gstrTitle
        ExitSetup frmSetup1, gintRET_FATAL
    End If

    Err = 0
End Sub

' FUNCTION: RegEnumKey
'
' Enumerates through the subkeys of an open registry
' key (returns the "i"th subkey of hkey, if it exists)
'
' Returns:
'   ERROR_SUCCESS on success.  strSubkeyName is set to the name of the subkey.
'   ERROR_NO_MORE_ITEMS if there are no more subkeys (32-bit only)
'   anything else - error
'
Function RegEnumKey(ByVal hKey As Long, ByVal i As Long, strKeyName As String) As Long
    Dim strResult As String
    
    strResult = String(300, " ")
    RegEnumKey = OSRegEnumKey(hKey, i, strResult, Len(strResult))
    strKeyName = StripTerminator(strResult)
End Function
'-----------------------------------------------------------
' SUB: RegisterDAO
'
' Special keys need to be added to the registry if
' DAO is installed.  This routine adds those keys.
'
' Note, these keys will not be uninstalled.
'
Sub RegisterDAO()
    Const strDAOKey = "CLSID\{F7A9C6E0-EFF2-101A-8185-00DD01108C6B}"
    Const strDAOKeyVal = "OLE 2.0 Link"
    Const strDAOInprocHandlerKey = "CLSID\{F7A9C6E0-EFF2-101A-8185-00DD01108C6B}\InprocHandler"
    Const strDAOInprocHandlerKeyVal = "ole2.dll"
    Const strDAOProgIDKey = "CLSID\{F7A9C6E0-EFF2-101A-8185-00DD01108C6B}\ProgID"
    Const strDAOProgIDKeyVal = "Access.OLE2Link"
    
    Dim hKey As Long
    
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strDAOKey, "", hKey) Then
        '
        ' RegCreateKey displays an error if something goes wrong.
        '
        GoTo REGDAOError
    End If
    '
    ' Set the key's value
    '
    If Not RegSetStringValue(hKey, "", strDAOKeyVal, False) Then
        '
        ' RegSetStringValue displays an error if something goes wrong.
        '
        GoTo REGDAOError
    End If
    '
    ' Close the key
    '
    RegCloseKey hKey
    '
    ' Repeat the same process for the other two keys.
    '
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strDAOInprocHandlerKey, "", hKey) Then GoTo REGDAOError
    If Not RegSetStringValue(hKey, "", strDAOInprocHandlerKeyVal, False) Then GoTo REGDAOError
    RegCloseKey hKey
    
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strDAOProgIDKey, "", hKey) Then GoTo REGDAOError
    If Not RegSetStringValue(hKey, "", strDAOProgIDKeyVal, False) Then GoTo REGDAOError
    RegCloseKey hKey

    Exit Sub
        
REGDAOError:
    '
    ' Error messages should have already been displayed.
    '
    ExitSetup frmSetup1, gintRET_FATAL
        
End Sub
'-----------------------------------------------------------
' SUB: RegisterFiles
'
' Loop through the list (array) of files to register that
' was created in the CopySection function and register
' each file therein as required
'
' Notes: msRegInfo() array created by CopySection function
'-----------------------------------------------------------
'
Sub RegisterFiles()
    Const strEXT_EXE$ = "EXE"

    Dim intIdx As Integer
    Dim intLastIdx As Integer
    Dim strFilename As String
    Dim strMsg As String
    Dim sDrive As String, sPath As String
    On Error Resume Next

    '
    'Get number of items to register, if none then we can get out of here
    '
    intLastIdx = UBound(msRegInfo)
    If Err > 0 Then
        GoTo RFCleanup
    End If

    For intIdx = 0 To intLastIdx
        strFilename = msRegInfo(intIdx).strFilename

        If Extension(msRegInfo(intIdx).strRegister) = gsEXT_REG Then
            If BaseName(msRegInfo(intIdx).strFilename) = BaseName(msRegInfo(intIdx).strRegister) Then
                Kill msRegInfo(intIdx).strRegister
            End If
            GoTo GoodToGo
        End If
        Select Case msRegInfo(intIdx).strRegister
            Case mstrDLLSELFREGISTER
                Dim intDllSelfRegRet As Integer
                Dim intErrRes As Integer
                Const FAIL_OLE = 2
                Const FAIL_LOAD = 3
                Const FAIL_ENTRY = 4
                Const FAIL_REG = 5
                
                NewAction gstrKEY_DLLSELFREGISTER, """" & strFilename & """"
                
RetryDllSelfReg:
                sDrive = CurDir
                sPath = CurDir
                ChDrive GetPathName(strFilename)
                ChDir GetPathName(strFilename)
                Err = 0
                intErrRes = 0
                intDllSelfRegRet = DLLSelfRegister(strFilename)
                If (Err <> 49) And (Err <> 0) Then
                    intErrRes = resCOMMON_CANTREGUNEXPECTED
                Else
                    Select Case intDllSelfRegRet
                        Case 0
                            'Good - everything's okay
                        Case FAIL_OLE
                            intErrRes = resCOMMON_CANTREGOLE
                        Case FAIL_LOAD
                            intErrRes = resCOMMON_CANTREGLOAD
                        Case FAIL_ENTRY
                            intErrRes = resCOMMON_CANTREGENTRY
                        Case FAIL_REG
                            intErrRes = resCOMMON_CANTREGREG
                        Case Else
                            intErrRes = resCOMMON_CANTREGUNEXPECTED
                        'End Case
                    End Select
                End If
                ChDrive sDrive
                ChDir sPath
                If intErrRes Then
                    'There was some kind of error
                    
                    'Log the more technical version of the error message -
                    'this would be too confusing to show to the end user
                    LogError ResolveResString(intErrRes, "|1", strFilename)
                    
                    'Now show a general error message to the user
AskWhatToDo:
                    strMsg = ResolveResString(resCOMMON_CANTREG, "|1", strFilename)
                    
                    Select Case MsgError(strMsg, vbExclamation Or vbAbortRetryIgnore, gstrTitle)
                        Case vbAbort:
                            ExitSetup frmSetup1, gintRET_ABORT
                            GoTo AskWhatToDo
                        Case vbRetry:
                            GoTo RetryDllSelfReg
                        Case vbIgnore:
                            AbortAction
                        'End Case
                    End Select
                Else
                    CommitAction
                End If
            Case mstrEXESELFREGISTER
                '
                'Only self register EXE files
                '
                If Extension(strFilename) = strEXT_EXE Then
                    NewAction gstrKEY_EXESELFREGISTER, """" & strFilename & """"
                    Err = 0
                    ExeSelfRegister strFilename
                    If Err Then
                        AbortAction
                    Else
                        CommitAction
                    End If
                End If
            Case mstrREMOTEREGISTER
                NewAction gstrKEY_REMOTEREGISTER, """" & strFilename & """"
                Err = 0
                RemoteRegister strFilename, msRegInfo(intIdx)
                If Err Then
                    AbortAction
                Else
                    CommitAction
                End If
            Case mstrTLBREGISTER
                NewAction gstrKEY_TLBREGISTER, """" & strFilename & """"
                '
                ' Call vb6stkit.dll's RegisterTLB export which calls
                ' LoadTypeLib and RegisterTypeLib.
                '
RetryTLBReg:
                If Not RegisterTLB(strFilename) Then
                    '
                    ' Registration of the TLB file failed.
                    '
                    LogError ResolveResString(resCOMMON_CANTREGTLB, "|1", strFilename)
TLBAskWhatToDo:
                    strMsg = ResolveResString(resCOMMON_CANTREGTLB, "|1", strFilename)
                    
                    Select Case MsgError(strMsg, vbExclamation Or vbAbortRetryIgnore, gstrTitle)
                        Case vbAbort:
                            ExitSetup frmSetup1, gintRET_ABORT
                            GoTo TLBAskWhatToDo
                        Case vbRetry:
                            GoTo RetryTLBReg
                        Case vbIgnore:
                            AbortAction
                        'End Case
                    End Select
                Else
                    CommitAction
                End If
            Case mstrVBLREGISTER
                '
                ' RegisterVBLFile takes care of logging, etc.
                '

                RegisterVBLFile strFilename
            Case Else
                RegEdit msRegInfo(intIdx).strRegister
            'End Case
        End Select
GoodToGo:
    Next


    Erase msRegInfo

RFCleanup:
    Err = 0
End Sub
'-----------------------------------------------------------
' SUB: RegisterLicenses
'
' Find all the setup.lst license entries and register
' them.
'-----------------------------------------------------------
'
Sub RegisterLicenses()
    Const strINI_LICENSES = "Licenses"
    Const strREG_LICENSES = "Licenses"
    Dim iLic As Integer
    Dim strLine As String
    Dim strLicKey As String
    Dim strLicVal As String
    Dim iCommaPos As Integer
    Dim strMsg As String
    Dim hkeyLicenses As Long
    Const strCopyright$ = "Licensing: Copying the keys may be a violation of established copyrights."

    'Make sure the Licenses key exists
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strREG_LICENSES, "", hkeyLicenses) Then
        'RegCreateKey will have already displayed an error message
        '  if something's wrong
        ExitSetup frmSetup1, gintRET_FATAL
    End If
    If Not RegSetStringValue(hkeyLicenses, "", strCopyright, False) Then
        RegCloseKey hkeyLicenses
        ExitSetup frmSetup1, gintRET_FATAL
    End If
    RegCloseKey hkeyLicenses
    
    iLic = 1
    Do
        strLine = ReadIniFile(gstrSetupInfoFile, strINI_LICENSES, gstrINI_LICENSE & iLic)
        If strLine = vbNullString Then
            '
            ' We've got all the licenses.
            '
            Exit Sub
        End If
        strLine = strUnQuoteString(strLine)
        '
        ' We have a license, parse it and register it.
        '
        iCommaPos = InStr(strLine, gstrCOMMA)
        If iCommaPos = 0 Then
            '
            ' Looks like the setup.lst file is corrupt.  There should
            ' always be a comma in the license information that separates
            ' the license key from the license value.
            '
            GoTo RLError
        End If
        strLicKey = Left(strLine, iCommaPos - 1)
        strLicVal = Mid(strLine, iCommaPos + 1)
        
        RegisterLicense strLicKey, strLicVal
        
        iLic = iLic + 1
    Loop While strLine <> vbNullString
    Exit Sub
        
RLError:
    strMsg = gstrSetupInfoFile & vbLf & vbLf & ResolveResString(resINVLINE) & vbLf & vbLf
    strMsg = strMsg & ResolveResString(resSECTNAME) & strINI_LICENSES & vbLf & strLine
    MsgError strMsg, vbCritical, gstrTitle
    ExitSetup frmSetup1, gintRET_FATAL
End Sub
'-----------------------------------------------------------
' SUB: RegisterLicense
'
' Register license information given the key and default
' value.  License information always goes into
' HKEY_CLASSES_ROOT\Licenses.
'-----------------------------------------------------------
'
Sub RegisterLicense(strLicKey As String, strLicVal As String)
    Const strREG_LICENSES = "Licenses"
    Dim hKey As Long
    '
    ' Create the key
    '
    If Not RegCreateKey(HKEY_CLASSES_ROOT, strREG_LICENSES, strLicKey, hKey) Then
        '
        ' RegCreateKey displays an error if something goes wrong.
        '
        GoTo REGError
    End If
    '
    ' Set the key's value
    '
    If Not RegSetStringValue(hKey, "", strLicVal, True) Then
        '
        ' RegSetStringValue displays an error if something goes wrong.
        '
        GoTo REGError
    End If
    '
    ' Close the key
    '
    RegCloseKey hKey

    Exit Sub
        
REGError:
    '
    ' Error messages should have already been displayed.
    '
    ExitSetup frmSetup1, gintRET_FATAL
End Sub
'-----------------------------------------------------------
' SUB: RegisterVBLFile
'
' Register license information in a VB License (vbl) file.
' Basically, parse out the license info and then call
' RegisterLicense.
'
' If strVBLFile is not a valid VBL file, nothing is
' registered.
'-----------------------------------------------------------
'
Sub RegisterVBLFile(strVBLFile As String)
    Dim strLicKey As String
    Dim strLicVal As String
    
    GetLicInfoFromVBL strVBLFile, strLicKey, strLicVal
    If strLicKey <> vbNullString Then
        RegisterLicense strLicKey, strLicVal
    End If
End Sub

'----------------------------------------------------------
' SUB: RegisterAppRemovalEXE
'
' Registers the application removal program (Windows 95 only)
' or else places an icon for it in the application directory.
'
' Returns True on success, False otherwise.
'----------------------------------------------------------
Function RegisterAppRemovalEXE(ByVal strAppRemovalEXE As String, ByVal strAppRemovalLog As String, ByVal strGroupName As String) As Boolean
    On Error GoTo Err
    
    Const strREGSTR_VAL_AppRemoval_APPNAMELINE = "ApplicationName"
    Const strREGSTR_VAL_AppRemoval_DISPLAYNAME = "DisplayName"
    Const strREGSTR_VAL_AppRemoval_COMMANDLINE = "UninstallString"
    Const strREGSTR_VAL_AppRemoval_APPTOUNINSTALL = "AppToUninstall"
    
    
    Dim strREGSTR_PATH_UNINSTALL As String
    strREGSTR_PATH_UNINSTALL = RegPathWinCurrentVersion() & "\Uninstall"
    
    'The command-line for the application removal executable is simply the path
    'for the installation logfile
    Dim strAppRemovalCmdLine As String
    strAppRemovalCmdLine = GetAppRemovalCmdLine(strAppRemovalEXE, strAppRemovalLog, vbNullString, False, APPREMERR_NONE)
    '
    ' Make sure that the Removal command line (including path, filename, commandline args, etc.
    ' is not longer than the max allowed, which is _MAX_PATH.
    '
    If Not fCheckFNLength(strAppRemovalCmdLine) Then
        Dim strMsg As String
        strMsg = ResolveResString(resCANTCREATEICONPATHTOOLONG) & vbLf & vbLf & ResolveResString(resCHOOSENEWDEST) & vbLf & vbLf & strAppRemovalCmdLine
        Call MsgError(strMsg, vbOKOnly, gstrSETMSG)
        ExitSetup frmCopy, gintRET_FATAL
        Exit Function
    End If
    '
    ' Create registry entries to tell Windows where the app removal executable is,
    ' how it should be displayed to the user, and what the command-line arguments are
    '
    Dim iAppend As Integer
    Dim fOk As Boolean
    Dim hkeyAppRemoval As Long
    Dim hkeyOurs As Long
    Dim i As Integer
    
    'Go ahead and create a key to the main Uninstall branch
    If Not RegCreateKey(HKEY_LOCAL_MACHINE, strREGSTR_PATH_UNINSTALL, "", hkeyAppRemoval) Then
        GoTo Err
    End If
    
    'We need a unique key.  This key is never shown to the end user.  We will use a key of
    'the form 'ST5UNST #xxx'
    Dim strAppRemovalKey As String
    Dim strAppRemovalKeyBase As String
    Dim hkeyTest As Long
    strAppRemovalKeyBase = mstrFILE_APPREMOVALLOGBASE$ & " #"
    iAppend = 1
    
    Do
        strAppRemovalKey = strAppRemovalKeyBase & Format(iAppend)
        If RegOpenKey(hkeyAppRemoval, strAppRemovalKey, hkeyTest) Then
            'This key already exists.  But we need a unique key.
            RegCloseKey hkeyTest
        Else
            'We've found a key that doesn't already exist.  Use it.
            Exit Do
        End If
        
        iAppend = iAppend + 1
    Loop
    
    '
    ' We also need a unique displayname.  This name is
    ' the only means the user has to identify the application
    ' to remove
    '
    Dim strDisplayName As String
    strDisplayName = gstrAppName 'First try... Application name
    If Not IsDisplayNameUnique(hkeyAppRemoval, strDisplayName) Then
        'Second try... Add path
        strDisplayName = strDisplayName & " (" & gstrDestDir & ")"
        If Not IsDisplayNameUnique(hkeyAppRemoval, strDisplayName) Then
            'Subsequent tries... Append a unique integer
            Dim strDisplayNameBase As String
            
            strDisplayNameBase = strDisplayName
            iAppend = 3
            Do
                strDisplayName = strDisplayNameBase & " #" & Format(iAppend)
                If IsDisplayNameUnique(hkeyAppRemoval, strDisplayName) Then
                    Exit Do
                Else
                    iAppend = iAppend + 1
                End If
            Loop
        End If
    End If
    
    'Go ahead and fill in entries for the app removal executable
    If Not RegCreateKey(hkeyAppRemoval, strAppRemovalKey, "", hkeyOurs) Then
        GoTo Err
    End If
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_APPNAMELINE, gstrAppExe, False) Then
        GoTo Err
    End If
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_DISPLAYNAME, strDisplayName, False) Then
        GoTo Err
    End If
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_COMMANDLINE, strAppRemovalCmdLine, False) Then
        GoTo Err
    End If
    If gstrAppToUninstall = vbNullString Then gstrAppToUninstall = gstrAppExe
    If Not RegSetStringValue(hkeyOurs, strREGSTR_VAL_AppRemoval_APPTOUNINSTALL, gstrAppToUninstall, False) Then
        GoTo Err
    End If
    If Not TreatAsWin95() Then
        '
        ' Under NT3.51, we simply place an icon to the app removal EXE in the program manager
        '
        If fMainGroupWasCreated Then
            CreateProgManItem frmSetup1, strGroupName, strAppRemovalCmdLine, ResolveResString(resAPPREMOVALICONNAME, "|1", gstrAppName)
        Else
            'If you get this message, it means that you incorrectly customized Form_Load().
            'Under 32-bits and NT 3.51, a Program Manager group must always be created.
            MsgError ResolveResString(resNOFOLDERFORICON, "|1", strAppRemovalEXE), vbOKOnly Or vbExclamation, gstrTitle
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
    
    RegCloseKey hkeyAppRemoval
    RegCloseKey hkeyOurs
    
    RegisterAppRemovalEXE = True
    Exit Function
    
Err:
    If hkeyOurs Then
        RegCloseKey hkeyOurs
        RegDeleteKey hkeyAppRemoval, strAppRemovalKey
    End If
    If hkeyAppRemoval Then
        RegCloseKey hkeyAppRemoval
    End If
    
    RegisterAppRemovalEXE = False
    Exit Function
End Function

'-----------------------------------------------------------
' FUNCTION: RegOpenKey
'
' Opens an existing key in the system registry.
'
' Returns: True if the key was opened OK, False otherwise
'   Upon success, phkResult is set to the handle of the key.
'-----------------------------------------------------------
'
Function RegOpenKey(ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Boolean
    Dim lResult As Long
    Dim strHkey As String

    On Error GoTo 0

    strHkey = strGetHKEYString(hKey)

    lResult = OSRegOpenKey(hKey, lpszSubKey, phkResult)
    If lResult = ERROR_SUCCESS Then
        RegOpenKey = True
        AddHkeyToCache phkResult, strHkey & "\" & lpszSubKey
    Else
        RegOpenKey = False
    End If
End Function
'----------------------------------------------------------
' FUNCTION: RegPathWinPrograms
'
' Returns the name of the registry key
' "\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
'----------------------------------------------------------
Function RegPathWinPrograms() As String
    RegPathWinPrograms = RegPathWinCurrentVersion() & "\Explorer\Shell Folders"
End Function
 
'----------------------------------------------------------
' FUNCTION: RegPathWinCurrentVersion
'
' Returns the name of the registry key
' "\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion"
'----------------------------------------------------------
Function RegPathWinCurrentVersion() As String
    RegPathWinCurrentVersion = "SOFTWARE\Microsoft\Windows\CurrentVersion"
End Function

'----------------------------------------------------------
' FUNCTION: RegQueryIntValue
'
' Retrieves the integer data for a named
' (strValueName = name) or unnamed (strValueName = "")
' value within a registry key.  If the named value
' exists, but its data is not a REG_DWORD, this function
' fails.
'
' NOTE: There is no 16-bit version of this function.
'
' Returns: True on success, else False.
'   On success, lData is set to the numeric data value
'
'----------------------------------------------------------
Function RegQueryNumericValue(ByVal hKey As Long, ByVal strValueName As String, lData As Long) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    
    RegQueryNumericValue = False
    
    On Error GoTo 0
    
    ' Get length/data type
    lDataBufSize = 4
        
    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            lData = lBuf
            RegQueryNumericValue = True
        End If
    End If
End Function

' FUNCTION: RegQueryStringValue
'
' Retrieves the string data for a named
' (strValueName = name) or unnamed (strValueName = "")
' value within a registry key.  If the named value
' exists, but its data is not a string, this function
' fails.
'
' NOTE: For 16-bits, strValueName MUST be "" (but the
' NOTE: parameter is left in for source code compatability)
'
' Returns: True on success, else False.
'   On success, strData is set to the string data value
'
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String, strData As String) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    
    RegQueryStringValue = False
    On Error GoTo 0
    ' Get length/data type
    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = OSRegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = True
                strData = StripTerminator(strBuf)
            End If
        End If
    End If
End Function

'----------------------------------------------------------
' FUNCTION: RegQueryRefCount
'
' Retrieves the data inteded as a reference count for a
' particular value within a registry key.  Although
' REG_DWORD is the preferred way of storing reference
' counts, it is possible that some installation programs
' may incorrect use a string or binary value instead.
' This routine accepts the data whether it is a string,
' a binary value or a DWORD (Long).
'
' NOTE: There is no 16-bit version of this function.
'
' Returns: True on success, else False.
'   On success, lrefcount is set to the numeric data value
'
'----------------------------------------------------------
Function RegQueryRefCount(ByVal hKey As Long, ByVal strValueName As String, lRefCount As Long) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long

    RegQueryRefCount = False

    On Error GoTo 0

    ' Get length/data type
    lDataBufSize = 4

    lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        Select Case lValueType
            Case REG_DWORD
                lRefCount = lBuf
                RegQueryRefCount = True
            Case REG_BINARY
                If lDataBufSize = 4 Then
                    lRefCount = lBuf
                    RegQueryRefCount = True
                End If
            Case REG_SZ
                Dim strRefCount As String
                
                If RegQueryStringValue(hKey, strValueName, strRefCount) Then
                    lRefCount = Val(strRefCount)
                    RegQueryRefCount = True
                End If
            'End Case
        End Select
    End If
End Function

' FUNCTION: RegSetNumericValue
'
' Associates a named (strValueName = name) or unnamed (strValueName = "")
'   value with a registry key.
'
' If fLog is missing or is True, then this action is logged in the logfile,
' and the value will be deleted by the application removal utility if the
' user choose to remove the installed application.
'
' NOTE: There is no 16-bit version of this function.
'
' Returns: True on success, else False.
'
Function RegSetNumericValue(ByVal hKey As Long, ByVal strValueName As String, ByVal lData As Long, Optional ByVal fLog As Boolean = True) As Boolean
    Dim lResult As Long
    Dim strHkey As String

    On Error GoTo 0
    
    strHkey = strGetHKEYString(hKey)
    
    If fLog Then
        NewAction _
          gstrKEY_REGVALUE, _
          """" & strHkey & """" _
            & ", " & """" & strValueName & """"
    End If

    lResult = OSRegSetValueNumEx(hKey, strValueName, 0, REG_DWORD, lData, 4)
    If lResult = ERROR_SUCCESS Then
        RegSetNumericValue = True
        If fLog Then
            CommitAction
        End If
    Else
        RegSetNumericValue = False
        MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
        If fLog Then
            AbortAction
        End If
        If gfNoUserInput Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
End Function

' FUNCTION: RegSetStringValue
'
' Associates a named (strValueName = name) or unnamed (strValueName = "")
'   value with a registry key.
'
' If fLog is missing or is True, then this action is logged in the
' logfile, and the value will be deleted by the application removal
' utility if the user choose to remove the installed application.
'
' Returns: True on success, else False.
'
Function RegSetStringValue(ByVal hKey As Long, ByVal strValueName As String, ByVal strData As String, Optional ByVal fLog As Boolean = True) As Boolean
    Dim lResult As Long
    Dim strHkey As String
    
    On Error GoTo 0
    
    If hKey = 0 Then
        Exit Function
    End If
    
    strHkey = strGetHKEYString(hKey)

    If fLog Then
        NewAction _
          gstrKEY_REGVALUE, _
          """" & strHkey & """" _
            & ", " & """" & strValueName & """"
    End If

    lResult = OSRegSetValueEx(hKey, strValueName, 0&, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)) + 1)
    'lResult = OSRegSetValueEx(hKey, strValueName, 0&, REG_SZ, ByVal strData, Len(strData) + 1)
    
    If lResult = ERROR_SUCCESS Then
        RegSetStringValue = True
        If fLog Then
            CommitAction
        End If
    Else
        RegSetStringValue = False
        MsgError ResolveResString(resERR_REG), vbOKOnly Or vbExclamation, gstrTitle
        If fLog Then
            AbortAction
        End If
        If gfNoUserInput Then
            ExitSetup frmSetup1, gintRET_FATAL
        End If
    End If
End Function

'-----------------------------------------------------------
' SUB: RemoteRegister
'
' Synchronously run the client registration utility on the
' given remote server registration file in order to set it
' up properly in the registry.
'
' IN: [strFileName] - .EXE file to register

'-----------------------------------------------------------
'
Sub RemoteRegister(ByVal strFilename As String, rInfo As REGINFO)
    Const strClientRegistrationUtility$ = "CLIREG32.EXE"
    Const strAddressSwitch = " /s "
    Const strProtocolSwitch = " /p "
    Const strSilentSwitch = " /q "
    Const strNoLogoSwitch = " /nologo "
    Const strAuthenticationSwitch = " /a "
    Const strTypelibSwitch = " /t "
    Const strDCOMSwitch = " /d "
    Const strEXT_REMOTE$ = "VBR"
    Const strEXT_REMOTETLB$ = "TLB"

    Dim strAddress As String
    Dim strProtocol As String
    Dim intAuthentication As Integer
    Dim strCmdLine As String
    Dim fShell As Integer
    Dim strMatchingTLB As String
    Dim fDCOM As Boolean

    'Find the name of the matching typelib file.  This should have already
    'been installed to the same directory as the .VBR file.
    strMatchingTLB = strFilename
    If Right$(strMatchingTLB, Len(strEXT_REMOTE)) = strEXT_REMOTE Then
        strMatchingTLB = Left$(strMatchingTLB, Len(strMatchingTLB) - Len(strEXT_REMOTE))
    End If
    strMatchingTLB = strMatchingTLB & strEXT_REMOTETLB

    strAddress = rInfo.strNetworkAddress
    strProtocol = rInfo.strNetworkProtocol
    intAuthentication = rInfo.intAuthentication
    fDCOM = rInfo.fDCOM
    frmRemoteServerDetails.GetServerDetails strFilename, strAddress, strProtocol, fDCOM
    frmMessage.Refresh
    strCmdLine = _
      strClientRegistrationUtility _
      & strAddressSwitch & """" & strAddress & """" _
      & IIf(fDCOM, " ", strProtocolSwitch & strProtocol) _
      & IIf(fDCOM, " ", strAuthenticationSwitch & Format$(intAuthentication) & " ") _
      & strNoLogoSwitch _
      & strTypelibSwitch & """" & strMatchingTLB & """" & " " _
      & IIf(fDCOM, strDCOMSwitch, "") _
      & IIf(gfNoUserInput, strSilentSwitch, "") _
      & """" & strFilename & """"
      
    '
    'Synchronously shell out and run the utility with the correct switches
    '
    fShell = SyncShell(strCmdLine, INFINITE, , False)

    If Not fShell Then
        MsgError ResolveResString(resCANTRUNPROGRAM, "|1", strClientRegistrationUtility), vbOKOnly Or vbExclamation, gstrTitle, gintRET_FATAL
        ExitSetup frmSetup1, gintRET_FATAL
    End If
End Sub

'-----------------------------------------------------------
' SUB: RemoveShellLink
'
' Removes a link in either Start>Programs or any of its

' immediate subfolders in the Windows 95 shell.
'
' IN: [strFolderName] - text name of the immediate folder
'                       in which the link to be removed
'                       currently exists, or else the
'                       empty string ("") to indicate that
'                       the link can be found directly in
'                       the Start>Programs menu.
'     [strLinkName] - text caption for the link
'
' This action is never logged in the app removal logfile.
'
' PRECONDITION: strFolderName has already been created and is
'               an immediate subfolder of Start>Programs, if it
'               is not equal to ""
'-----------------------------------------------------------
'
Sub RemoveShellLink(ByVal strFolderName As String, ByVal strLinkName As String)
    Dim fSuccess As Boolean
    
    ReplaceDoubleQuotes strFolderName
    ReplaceDoubleQuotes strLinkName
    
    fSuccess = OSfRemoveShellLink(strFolderName, strLinkName)
End Sub

'-----------------------------------------------------------
' FUNCTION: ResolveDestDir
'
' Given a destination directory string, equate any macro
' portions of the string to their runtime determined
' actual locations and return a string reflecting the
' actual path.
'
' IN: [strDestDir] - string containing directory macro info
'                    and/or actual dir path info
'
'     [fAssumeDir] - boolean that if true, causes this routine
'                    to assume that strDestDir contains a dir
'                    path.  If a directory isn't given it will
'                    make it the application path.  If false,
'                    this routine will return strDestDir as
'                    is after performing expansion.  Set this
'                    to False when you are not sure it is a
'                    directory but you want to expand macros
'                    if it contains any.  E.g., If this is a
'                    command line parameter, you can't be
'                    certain if it refers to a path.  In this
'                    case, set fAssumeDir = False.  Default
'                    is True.
'
' Return: A string containing the resolved dir name
'-----------------------------------------------------------
'
Function ResolveDestDir(ByVal strDestDir As String, Optional fAssumeDir As Boolean = True) As String
    Const strMACROSTART$ = "$("
    Const strMACROEND$ = ")"

    Dim intPos As Integer
    Dim strResolved As String
    Dim hKey As Long
    Dim strPathsKey As String
    Dim fQuoted As Boolean
    
    strPathsKey = RegPathWinCurrentVersion()
    strDestDir = Trim(strDestDir)
    '
    ' If strDestDir is quoted when passed to this routine, it
    ' should be quoted when it's returned.  The quotes need
    ' to be temporarily removed, though, for processing.
    '
    If Left(strDestDir, 1) = gstrQUOTE Then
        fQuoted = True
        strDestDir = strUnQuoteString(strDestDir)
    End If
    '
    ' We take the first part of destdir, and if its $( then we need to get the portion
    ' of destdir up to and including the last paren.  We then test against this for
    ' macro expansion.  If no ) is found after finding $(, then must assume that it's
    ' just a normal file name and do no processing.  Only enter the case statement
    ' if strDestDir starts with $(.
    '
    If Left$(strDestDir, 2) = strMACROSTART Then
        intPos = InStr(strDestDir, strMACROEND)

        Select Case Left$(strDestDir, intPos)
            Case gstrAPPDEST
                If gstrDestDir <> vbNullString Then

                    strResolved = gstrDestDir
                Else
                    strResolved = "?"
                End If
            Case gstrWINDEST
                strResolved = gstrWinDir
            Case gstrFONTDEST
                strResolved = gstrFontDir
            Case gstrWINSYSDEST, gstrWINSYSDESTSYSFILE
                strResolved = gstrWinSysDir
            Case gstrPROGRAMFILES
                If TreatAsWin95() Then
                    Const strProgramFilesKey = "ProgramFilesDir"
    
                    If RegOpenKey(HKEY_LOCAL_MACHINE, strPathsKey, hKey) Then
                        RegQueryStringValue hKey, strProgramFilesKey, strResolved
                        RegCloseKey hKey
                    End If
                End If
    
                If strResolved = "" Then
                    'If not otherwise set, let strResolved be the root of the first fixed disk
                    strResolved = strRootDrive()
                End If
            Case gstrCOMMONFILES
                'First determine the correct path of Program Files\Common Files, if under Win95
                strResolved = strGetCommonFilesPath()
                If strResolved = "" Then
                    'If not otherwise set, let strResolved be the Windows directory
                    strResolved = gstrWinDir
                End If
            Case gstrCOMMONFILESSYS
                'First determine the correct path of Program Files\Common Files, if under Win95
                Dim strCommonFiles As String
                
                strCommonFiles = strGetCommonFilesPath()
                If strCommonFiles <> "" Then
                    'Okay, now just add \System, and we're done
                    strResolved = strCommonFiles & "System\"
                Else
                    'If Common Files isn't in the registry, then map the
                    'entire macro to the Windows\{system,system32} directory
                    strResolved = gstrWinSysDir
                End If
            Case gstrDAODEST
                strResolved = strGetDAOPath()
            Case Else
                intPos = 0
            'End Case
        End Select
    End If
    
    If intPos <> 0 Then
        AddDirSep strResolved
    End If

    If fAssumeDir = True Then
        If intPos = 0 Then
            '
            'if no drive spec, and doesn't begin with any root path indicator ("\"),
            'then we assume that this destination is relative to the app dest dir
            '
            If Mid$(strDestDir, 2, 1) <> gstrCOLON Then
                If Left$(strDestDir, 1) <> gstrSEP_DIR Then
                    strResolved = gstrDestDir
                End If
            End If
        Else
            If Mid$(strDestDir, intPos + 1, 1) = gstrSEP_DIR Then
                intPos = intPos + 1
            End If
        End If
    End If

    If fQuoted = True Then
        ResolveDestDir = strQuoteString(strResolved & Mid$(strDestDir, intPos + 1), True, False)
    Else
        ResolveDestDir = strResolved & Mid$(strDestDir, intPos + 1)
    End If
End Function
'-----------------------------------------------------------
' FUNCTION: ResolveDestDirs
'
' Given a space delimited string, this routine finds all
' Destination directory macros and expands them by making
' repeated calls to ResolveDestDir.  See ResolveDestDir.
'
' Note that the macro must immediately follow a space (or
' a space followed by a quote) delimiter or else it will
' be ignored.
'
' Note that this routine does not assume that each item
' in the delimited string is actually a directory path.
' Therefore, the last parameter in the call to ResolveDestDir,
' below, is false.
'
' IN: [str] - string containing directory macro(s) info
'             and/or actual dir path info
'
' Return: str with destdir macros expanded.
'-----------------------------------------------------------
'
Function ResolveDestDirs(str As String)
    Dim intAnchor As Integer
    Dim intOffset As Integer
    Dim strField As String
    Dim strExpField As String
    Dim strExpanded As String
    
    If Len(Trim(strUnQuoteString(str))) = 0 Then
        ResolveDestDirs = str
        Exit Function
    End If
        
    intAnchor = 1
    strExpanded = ""
    
    Do
        intOffset = intGetNextFldOffset(intAnchor, str, " ")
        If intOffset = 0 Then intOffset = Len(str) + 1
        strField = Mid(str, intAnchor, intOffset - intAnchor)
        strExpField = ResolveDestDir(strField, False)
        strExpanded = strExpanded & strExpField & " "
        intAnchor = intOffset + 1
    Loop While intAnchor < Len(str)
    
    ResolveDestDirs = Trim(strExpanded)
End Function
'-----------------------------------------------------------
' FUNCTION: ResolveDir
'
' Given a pathname, resolve it to its smallest form.  If
' the pathname is invalid, then optionally warn the user.
'
' IN: [strPathName] - pathname to resolve
'     [fMustExist] - enforce that the path actually exists
'     [fWarn] - If True, warn user upon invalid path
'
' Return: A string containing the resolved dir name
'-----------------------------------------------------------
'
Function ResolveDir(ByVal strPathName As String, fMustExist As Integer, fWarn As Integer) As String
    Dim strMsg As String
    Dim fInValid As Integer
    Dim strUnResolvedPath As String
    Dim strResolvedPath As String
    Dim strIgnore As String
    Dim cbResolved As Long

    On Error Resume Next

    fInValid = False
    '
    'If the pathname is a UNC name (16-bit only), or if it's in actuality a file name, then it's invalid
    '
    If FileExists(strPathName) = True Then
        fInValid = True
        GoTo RDContinue
    End If

    strUnResolvedPath = strPathName

    If InStr(3, strUnResolvedPath, gstrSEP_DIR) > 0 Then

        strResolvedPath = Space(gintMAX_PATH_LEN * 2)
        cbResolved = GetFullPathName(strUnResolvedPath, gintMAX_PATH_LEN, strResolvedPath, strIgnore)
        If cbResolved = 0 Then
            '
            ' The path couldn't be resolved.  If we can actually
            ' switch to the directory we want, continue anyway.
            '
            ChDir strUnResolvedPath
            AddDirSep strUnResolvedPath
            If Err > 0 Then
                Err = 0
                ChDir strUnResolvedPath
                If Err > 0 Then
                    fInValid = True
                Else
                    strResolvedPath = strUnResolvedPath
                End If
            Else
                strResolvedPath = strUnResolvedPath
            End If
        Else
            '
            ' GetFullPathName returned us a NULL terminated string in
            ' strResolvedPath.  Remove the NULL.
            '
            strResolvedPath = StripTerminator(strResolvedPath)
            If CheckDrive(strResolvedPath, gstrTitle) = False Then
                fInValid = True
            Else
                AddDirSep strResolvedPath
                If fMustExist = True Then
                    Err = 0
                    
                    Dim strDummy As String
                    strDummy = Dir$(strResolvedPath & "*.*")
                    
                    If Err > 0 Then
                        strMsg = ResolveResString(resNOTEXIST) & vbLf & vbLf
                        fInValid = True
                    End If
                End If
            End If
        End If
    Else
        fInValid = True
    End If

RDContinue:
    If fInValid = True Then
        If fWarn = True Then
            strMsg = strMsg & ResolveResString(resDIRSPECIFIED) & vbLf & vbLf & strPathName & vbLf & vbLf
            strMsg = strMsg & ResolveResString(resDIRINVALID)
            MsgError strMsg, vbOKOnly Or vbExclamation, ResolveResString(resDIRINVNAME)
            If gfNoUserInput Then
                ExitSetup frmSetup1, gintRET_FATAL
            End If
        End If

        ResolveDir = vbNullString
    Else
        ResolveDir = strResolvedPath
    End If

    Err = 0
End Function

'-----------------------------------------------------------
' SUB: RestoreProgMan
'
' Restores Windows Program Manager
'-----------------------------------------------------------
'
Sub RestoreProgMan()
    Const strPMTITLE$ = "Program Manager"

    On Error Resume Next

    'Try the localized name first
    AppActivate ResolveResString(resPROGRAMMANAGER)
    
    If Err Then
        'If that doesn't work, try the English name
        AppActivate strPMTITLE
    End If

    Err = 0
End Sub

'-----------------------------------------------------------
' SUB: ShowPathDialog
'
' Display form to allow user to get either a source or
' destination path
'
' IN: [strPathRequest] - determines whether to ask for the
'                        source or destination pathname.
'                        gstrDIR_SRC for source path
'                        gstrDIR_DEST for destination path
'-----------------------------------------------------------
'
Sub ShowPathDialog(ByVal strPathRequest As String)
    frmSetup1.Tag = strPathRequest

    '
    'frmPath.Form_Load() reads frmSetup1.Tag to determine whether
    'this is a request for the source or destination path
    '
    frmPath.Show vbModal

    If strPathRequest = gstrDIR_SRC Then
        gstrSrcPath = frmSetup1.Tag
    Else
        If gfRetVal = gintRET_CONT Then
            gstrDestDir = frmSetup1.Tag
        End If
    End If
End Sub

'-----------------------------------------------------------
' FUNCTION: strExtractFilenameArg
'
' Extracts a quoted or unquoted filename from a string
'   containing command-line arguments
'
' IN: [str] - string containing a filename.  This filename
'             begins at the first character, and continues
'             to the end of the string or to the first space
'             or switch character, or, if the string begins
'             with a double quote, continues until the next
'             double quote
' OUT: Returns the filename, without quotes
'      str is set to be the remainder of the string after
'      the filename and quote (if any)
'
'-----------------------------------------------------------
'
Function strExtractFilenameArg(str As String, fErr As Boolean)
    Dim strFilename As String
    
    str = Trim$(str)
    
    Dim iEndFilenamePos As Integer
    If Left$(str, 1) = """" Then
        ' Filenames is surrounded by quotes
        iEndFilenamePos = InStr(2, str, """") ' Find matching quote
        If iEndFilenamePos > 0 Then
            strFilename = Mid$(str, 2, iEndFilenamePos - 2)
            str = Right$(str, Len(str) - iEndFilenamePos)
        Else
            fErr = True
            Exit Function
        End If
    Else
        ' Filename continues until next switch or space or quote
        Dim iSpacePos As Integer
        Dim iSwitch1 As Integer
        Dim iSwitch2 As Integer
        Dim iQuote As Integer
        
        iSpacePos = InStr(str, " ")
        iSwitch2 = InStr(str, gstrSwitchPrefix2)
        iQuote = InStr(str, """")
        
        If iSpacePos = 0 Then iSpacePos = Len(str) + 1
        If iSwitch1 = 0 Then iSwitch1 = Len(str) + 1
        If iSwitch2 = 0 Then iSwitch2 = Len(str) + 1
        If iQuote = 0 Then iQuote = Len(str) + 1
        
        iEndFilenamePos = iSpacePos
        If iSwitch2 < iEndFilenamePos Then iEndFilenamePos = iSwitch2
        If iQuote < iEndFilenamePos Then iEndFilenamePos = iQuote
        
        strFilename = Left$(str, iEndFilenamePos - 1)
        If iEndFilenamePos > Len(str) Then
            str = ""
        Else
            str = Right(str, Len(str) - iEndFilenamePos + 1)
        End If
    End If
    
    strFilename = Trim$(strFilename)
    If strFilename = "" Then
        fErr = True
        Exit Function
    End If
    
    fErr = False
    strExtractFilenameArg = strFilename
    str = Trim$(str)
End Function



'-----------------------------------------------------------
' SUB: UpdateStatus
'
' "Fill" (by percentage) inside the PictureBox and also
' display the percentage filled
'
' IN: [pic] - PictureBox used to bound "fill" region
'     [sngPercent] - Percentage of the shape to fill
'     [fBorderCase] - Indicates whether the percentage
'        specified is a "border case", i.e. exactly 0%
'        or exactly 100%.  Unless fBorderCase is True,
'        the values 0% and 100% will be assumed to be
'        "close" to these values, and 1% and 99% will
'        be used instead.
'
' Notes: Set AutoRedraw property of the PictureBox to True
'        so that the status bar and percentage can be auto-
'        matically repainted if necessary
'-----------------------------------------------------------
'
Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Single, Optional ByVal fBorderCase As Boolean = False)
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer

    'For this to work well, we need a white background and any color foreground (blue)
    Const colBackground = &HFFFFFF ' white
    Const colForeground = &H800000 ' dark blue

    pic.ForeColor = colForeground
    pic.BackColor = colBackground
    
    '
    'Format percentage and get attributes of text
    '
    Dim intPercent
    intPercent = Int(100 * sngPercent + 0.5)
    
    'Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
    'prevents, for instance, the status bar from reaching 100% until we are entirely done.
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    
    strPercent = Format$(intPercent) & "%"
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)

    '
    'Now set intX and intY to the starting location for printing the percentage
    '
    intX = pic.Width / 2 - intWidth / 2
    intY = pic.Height / 2 - intHeight / 2

    '
    'Need to draw a filled box with the pics background color to wipe out previous
    'percentage display (if any)
    '
    pic.DrawMode = 13 ' Copy Pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF

    '
    'Back to the center print position and print the text
    '
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent

    '
    'Now fill in the box with the ribbon color to the desired percentage
    'If percentage is 0, fill the whole box with the background color to clear it
    'Use the "Not XOR" pen so that we change the color of the text to white
    'wherever we touch it, and change the color of the background to blue
    'wherever we touch it.
    '
    pic.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
        pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF
    Else
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
    End If

    pic.Refresh
End Sub

'-----------------------------------------------------------
' FUNCTION: WriteAccess
'
' Determines whether there is write access to the specified
' directory.
'
' IN: [strDirName] - directory to check for write access
'
' Returns: True if write access, False otherwise
'-----------------------------------------------------------
'
Function WriteAccess(ByVal strDirName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    AddDirSep strDirName

    intFileNum = FreeFile
    Open strDirName & mstrCONCATFILE For Output As intFileNum

    WriteAccess = IIf(Err, False, True)
    
    Close intFileNum

    Kill strDirName & mstrCONCATFILE

    Err = 0
End Function
'-----------------------------------------------------------
' FUNCTION: WriteMIF
'
' If this is a SMS install, this routine writes the
' failed MIF status file if something goes wrong or
' a successful MIF if everything installs correctly.
'
' The MIF file requires a special format specified
' by SMS.  Currently, this routine implements the
' minimum requirements.  The hardcoded strings below
' that are written to the MIF should be written
' character by character as they are; except that
' status message should change depending on the
' circumstances of the install.  DO NOT LOCALIZE
' anything except the status message.
'
' IN: [strMIFFilename] - The name of the MIF file.
'                        Passed in to setup1 by
'                        setup.exe.  It is probably
'                        named <appname>.mif where
'                        <appname> is the name of the
'                        application you are installing.
'
'     [fStatus] - False to write a failed MIF (i.e. setup
'                 failed); True to write a successful MIF.
'
'     [strSMSDescription] - This is the description string
'                           to be written to the MIF file.
'                           It cannot be longer than 255
'                           characters and cannot contain
'                           carriage returns and/or line
'                           feeds.  This routine will
'                           enforce these requirements.
'
' Note, when running in SMS mode, there is no other way
' to display a message to the user than to write it to
' the MIF file.  Displaying a MsgBox will cause the
' computer to appear as if it has hung.  Therefore, this
' routine makes no attempt to display an error message.
'
'-----------------------------------------------------------
'
Sub WriteMIF(ByVal strMIFFilename As String, ByVal fStatus As Boolean, ByVal strSMSDescription As String)
    Const strSUCCESS = """SUCCESS"""                 ' Cannot be localized as per SMS
    Const strFAILED = """FAILED"""                   ' Cannot be localized as per SMS
    
    Dim fn As Integer
    Dim intOffset As Integer
    Dim fOpened As Boolean
        
    fOpened = False
        
    On Error GoTo WMIFFAILED  ' If we fail, we just return without doing anything
                              ' because there is no way to inform the user while
                              ' in SMS mode.

    '
    ' If the description string is greater than 255 characters,
    ' truncate it.  Required my SMS.
    '
    strSMSDescription = Left(strSMSDescription, MAX_SMS_DESCRIP)
    '
    ' Remove any carriage returns or line feeds and replace
    ' them with spaces.  The message must be a single line.
    '
    For intOffset = 1 To Len(strSMSDescription)
        If (Mid(strSMSDescription, intOffset, 1) = Chr(10)) Or (Mid(strSMSDescription, intOffset, 1) = Chr(13)) Then
            Mid(strSMSDescription, intOffset, 1) = " "
        End If
    Next intOffset
    '
    ' Open the MIF file for append, but first delete any existing
    ' ones with the same name.  Note, that setup.exe passed a
    ' unique name so if there is one with this name already in
    ' on the disk, it was put there by setup.exe.
    '
    If FileExists(strMIFFilename) Then
        Kill strMIFFilename
    End If
    
    fn = FreeFile
    Open strMIFFilename For Append As fn
    fOpened = True
    '
    ' We are ready to write the actual MIF file
    ' Note, none of the string below are supposed
    ' to be localized.
    '
    Print #fn, "Start Component"
        Print #fn, Tab; "Name = ""Workstation"""
        Print #fn, Tab; "Start Group"
            Print #fn, Tab; Tab; "Name = ""InstallStatus"""
            Print #fn, Tab; Tab; "ID = 1"
            Print #fn, Tab; Tab; "Class = ""MICROSOFT|JOBSTATUS|1.0"""
            Print #fn, Tab; Tab; "Start Attribute"
                Print #fn, Tab; Tab; Tab; "Name = ""Status"""
                Print #fn, Tab; Tab; Tab; "ID = 1"
                Print #fn, Tab; Tab; Tab; "Type = String(16)"
                Print #fn, Tab; Tab; Tab; "Value = "; IIf(fStatus, strSUCCESS, strFAILED)
            Print #fn, Tab; Tab; "End Attribute"
            Print #fn, Tab; Tab; "Start Attribute"
                Print #fn, Tab; Tab; Tab; "Name = ""Description"""
                Print #fn, Tab; Tab; Tab; "ID = 2"
                Print #fn, Tab; Tab; Tab; "Type = String(256)"
                Print #fn, Tab; Tab; Tab; "Value = "; strSMSDescription
            Print #fn, Tab; Tab; "End Attribute"
        Print #fn, Tab; "End Group"
    Print #fn, "End Component"

    Close fn
    '
    ' Success
    '
    Exit Sub

WMIFFAILED:
    '
    ' At this point we are unable to create the MIF file.
    ' Since we are running under SMS there is no one to
    ' tell, so we don't generate an error message at all.
    '
    If fOpened = True Then
        Close fn
    End If
    Exit Sub
End Sub

'Adds or replaces an HKEY to the list of HKEYs in cache.
'Note that it is not necessary to remove keys from
'this list.
Private Sub AddHkeyToCache(ByVal hKey As Long, ByVal strHkey As String)
    Dim intIdx As Integer
    
    intIdx = intGetHKEYIndex(hKey)
    If intIdx < 0 Then
        'The key does not already exist.  Add it to the end.
        On Error Resume Next
        ReDim Preserve hkeyCache(0 To UBound(hkeyCache) + 1)
        If Err Then
            'If there was an error, it means the cache was empty.
            On Error GoTo 0
            ReDim hkeyCache(0 To 0)
        End If
        On Error GoTo 0

        intIdx = UBound(hkeyCache)
    Else
        'The key already exists.  It will be replaced.
    End If

    hkeyCache(intIdx).hKey = hKey
    hkeyCache(intIdx).strHkey = strHkey
End Sub

'Given a predefined HKEY, return the text string representing that
'key, or else return "".
Private Function strGetPredefinedHKEYString(ByVal hKey As Long) As String
    Select Case hKey
        Case HKEY_CLASSES_ROOT
            strGetPredefinedHKEYString = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_USER
            strGetPredefinedHKEYString = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE
            strGetPredefinedHKEYString = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS
            strGetPredefinedHKEYString = "HKEY_USERS"
        'End Case
    End Select
End Function

'Given an HKEY, return the text string representing that
'key.
Private Function strGetHKEYString(ByVal hKey As Long) As String
    Dim strKey As String

    'Is the hkey predefined?
    strKey = strGetPredefinedHKEYString(hKey)
    If strKey <> "" Then
        strGetHKEYString = strKey
        Exit Function
    End If
    
    'It is not predefined.  Look in the cache.
    Dim intIdx As Integer
    intIdx = intGetHKEYIndex(hKey)
    If intIdx >= 0 Then
        strGetHKEYString = hkeyCache(intIdx).strHkey
    Else
        strGetHKEYString = ""
    End If
End Function

'Searches the cache for the index of the given HKEY.
'Returns the index if found, else returns -1.
Private Function intGetHKEYIndex(ByVal hKey As Long) As Integer
    Dim intUBound As Integer
    
    On Error Resume Next
    intUBound = UBound(hkeyCache)
    If Err Then
        'If there was an error accessing the ubound of the array,
        'then the cache is empty
        GoTo NotFound
    End If
    On Error GoTo 0

    Dim intIdx As Integer
    For intIdx = 0 To intUBound
        If hkeyCache(intIdx).hKey = hKey Then
            intGetHKEYIndex = intIdx
            Exit Function
        End If
    Next intIdx
    
NotFound:
    intGetHKEYIndex = -1
End Function

'Returns the location of the Program Files\Common Files path, if
'it is present in the registry.  Otherwise, returns "".
Public Function strGetCommonFilesPath() As String
    Dim hKey As Long
    Dim strPath As String
    
    If TreatAsWin95() Then
        Const strCommonFilesKey = "CommonFilesDir"

        If RegOpenKey(HKEY_LOCAL_MACHINE, RegPathWinCurrentVersion(), hKey) Then
            RegQueryStringValue hKey, strCommonFilesKey, strPath
            RegCloseKey hKey
        End If
    End If

    If strPath <> "" Then
        AddDirSep strPath
    End If
    
    strGetCommonFilesPath = strPath
End Function
'Returns the location of the "Windows\Start Menu\Programs" Files path, if
'it is present in the registry.  Otherwise, returns "".
Public Function strGetProgramsFilesPath() As String
    Dim hKey As Long
    Dim strPath As String
    
    strPath = ""
    If TreatAsWin95() Then
        Const strProgramsKey = "Programs"

        If RegOpenKey(HKEY_CURRENT_USER, RegPathWinPrograms(), hKey) Then
            RegQueryStringValue hKey, strProgramsKey, strPath
            RegCloseKey hKey
        End If
    End If

    If strPath <> "" Then
        AddDirSep strPath
    End If
    
    strGetProgramsFilesPath = strPath
End Function

'Returns the directory where DAO is or should be installed.  If the
'key does not exist in the registry, it is created.  For instance, under
'NT 3.51 this location is normally 'C:\WINDOWS\MSAPPS\DAO'
Private Function strGetDAOPath() As String
    Const strMSAPPS$ = "MSAPPS\"
    Const strDAO3032$ = "DAO350.DLL"
    
    'first look in the registry
    Const strKey = "SOFTWARE\Microsoft\Shared Tools\DAO350"
    Const strValueName = "Path"
    Dim hKey As Long
    Dim strPath As String

    If RegOpenKey(HKEY_LOCAL_MACHINE, strKey, hKey) Then
        RegQueryStringValue hKey, strValueName, strPath
        RegCloseKey hKey
    End If

    If strPath <> "" Then
        strPath = GetPathName(strPath)
        AddDirSep strPath
        strGetDAOPath = strPath
        Exit Function
    End If
    
    'It's not yet in the registry, so we need to decide
    'where the directory should be, and then need to place
    'that location in the registry.

    If TreatAsWin95() Then
        'For Win95, use "Common Files\Microsoft Shared\DAO"
        strPath = strGetCommonFilesPath() & ResolveResString(resMICROSOFTSHARED) & "DAO\"
    Else
        'Otherwise use Windows\MSAPPS\DAO
        strPath = gstrWinDir & strMSAPPS & "DAO\"
    End If
    
    'Place this information in the registry (note that we point to DAO3032.DLL
    'itself, not just to the directory)
    If RegCreateKey(HKEY_LOCAL_MACHINE, strKey, "", hKey) Then
        RegSetStringValue hKey, strValueName, strPath & strDAO3032, False
        RegCloseKey hKey
    End If

    strGetDAOPath = strPath
End Function

' Replace all double quotes with single quotes
Public Sub ReplaceDoubleQuotes(str As String)
    Dim i As Integer
    
    For i = 1 To Len(str)
        If Mid$(str, i, 1) = """" Then
            Mid$(str, i, 1) = "'"
        End If
    Next i
End Sub

'Get the path portion of a filename
Function GetPathName(ByVal strFilename As String) As String
    Dim sPath As String
    Dim sFile As String
    
    SeparatePathAndFileName strFilename, sPath, sFile
    
    GetPathName = sPath
End Function
'Determines if a character is a path separator (\ or /).
Public Function IsSeparator(Character As String) As Boolean
    Select Case Character
        Case gstrSEP_DIR
            IsSeparator = True
        Case gstrSEP_DIRALT
            IsSeparator = True
    End Select
End Function
'Given a fully qualified filename, returns the path portion and the file
'   portion.
Public Sub SeparatePathAndFileName(FullPath As String, ByRef Path As String, _
    ByRef FileName As String)

    Dim nSepPos As Long
    Dim sSEP As String

    nSepPos = Len(FullPath)
    sSEP = Mid$(FullPath, nSepPos, 1)
    Do Until IsSeparator(sSEP)
        nSepPos = nSepPos - 1
        If nSepPos = 0 Then Exit Do
        sSEP = Mid$(FullPath, nSepPos, 1)
    Loop

    Select Case nSepPos
        Case 0
            'Separator was not found.
            Path = CurDir$
            FileName = FullPath
        Case Else
            Path = Left$(FullPath, nSepPos - 1)
            FileName = Mid$(FullPath, nSepPos + 1)
    End Select
End Sub

'Returns the path to the root of the first fixed disk
Function strRootDrive() As String
    Dim intDriveNum As Integer
    
    For intDriveNum = 0 To Asc("Z") - Asc("A") - 1
        If GetDriveType(intDriveNum) = intDRIVE_FIXED Then
            strRootDrive = Chr$(Asc("A") + intDriveNum) & gstrCOLON & gstrSEP_DIR
            Exit Function
        End If
    Next intDriveNum
    
    strRootDrive = "C:\"
End Function

'Returns "" if the path is not complete, or is a UNC pathname
Function strGetDriveFromPath(ByVal strPath As String) As String
    If Len(strPath) < 2 Then
        Exit Function
    End If
    
    If Mid$(strPath, 2, 1) <> gstrCOLON Then
        Exit Function
    End If
    
    strGetDriveFromPath = Mid$(strPath, 1, 1) & gstrCOLON & gstrSEP_DIR
End Function

Public Function fValidFilename(strFilename As String) As Boolean
'
' This routine verifies that strFileName is a valid file name.
' It checks that its length is less than the max allowed
' and that it doesn't contain any invalid characters..
'
    If Not fCheckFNLength(strFilename) Then
        '
        ' Name is too long.
        '
        fValidFilename = False
        Exit Function
    End If
    '
    ' Search through the list of invalid filename characters and make
    ' sure none of them are in the string.
    '
    Dim iInvalidChar As Integer
    Dim iFilename As Integer
    Dim strInvalidChars As String
    
    strInvalidChars = ResolveResString(resCOMMON_INVALIDFILECHARS)
    
    For iInvalidChar = 1 To Len(strInvalidChars)
        If InStr(strFilename, Mid$(strInvalidChars, iInvalidChar, 1)) <> 0 Then
            fValidFilename = False
            Exit Function
        End If
    Next iInvalidChar
    
    fValidFilename = True
    
End Function
Public Function fValidNTGroupName(strGroupName) As Boolean
'
' This routine verifies that strGroupName is a valid group name.
' It checks that its length is less than the max allowed
' and that it doesn't contain any invalid characters.
'
    If Len(strGroupName) > gintMAX_GROUPNAME_LEN Then
        fValidNTGroupName = False
        Exit Function
    End If
    '
    ' Search through the list of invalid filename characters and make
    ' sure none of them are in the string.
    '
    Dim iInvalidChar As Integer
    Dim iFilename As Integer
    Dim strInvalidChars As String
    
    strInvalidChars = ResolveResString(resGROUPINVALIDCHARS)
    
    For iInvalidChar = 1 To Len(strInvalidChars)
        If InStr(strGroupName, Mid$(strInvalidChars, iInvalidChar, 1)) <> 0 Then
            fValidNTGroupName = False
            Exit Function
        End If
    Next iInvalidChar
    
    fValidNTGroupName = True
    
End Function
'-----------------------------------------------------------
' SUB: CountGroups
'
' Determines how many groups must be installed by counting
' them in the setup information file (SETUP.LST)
'-----------------------------------------------------------
'
Function CountGroups(ByVal strsection As String) As Integer
    Dim intIdx As Integer
    Dim sGroup As String
    
    intIdx = 0
    Do
        sGroup = ReadIniFile(gstrSetupInfoFile, strsection, gsGROUP & CStr(intIdx))
        If sGroup <> vbNullString Then 'Found a group
            intIdx = intIdx + 1
        Else
            Exit Do
        End If
    Loop
    CountGroups = intIdx
End Function
'-----------------------------------------------------------
' SUB: GetGroup
'
' Returns the Groupname specified by Index
'-----------------------------------------------------------
'
Function GetGroup(ByVal strsection As String, ByVal index As Integer)
    GetGroup = ReadIniFile(gstrSetupInfoFile, strsection, gsGROUP & CStr(index))
End Function
'-----------------------------------------------------------
' SUB: SetGroup
'
' Sets Groupname specified by Index
'-----------------------------------------------------------
'
Sub SetGroup(ByVal strsection As String, ByVal index As Integer, ByVal sGroupName As String)
    Const iBuf As Integer = 2048
    Const sEQUAL As String * 1 = "="
    Dim sGroup As String
    Dim sNames As String, ret As Long
    sGroup = ReadIniFile(gstrSetupInfoFile, strsection, gsGROUP & CStr(index))
    sNames = Space$(iBuf)
    ret = GetPrivateProfileSection(sGroup, sNames, iBuf, gstrSetupInfoFile)
    If ret = 0 Then 'We have nothing in this section, just quit.
        Exit Sub
    End If
    sNames = Left$(sNames, ret - 1)
    'We now have the Group name, modify the icons in that group
    Dim lCount As Long, sKEY As String, sValue As String, fKey As Boolean
    fKey = True
    For lCount = 1 To Len(sNames)
        If (Mid$(sNames, lCount, 1) = sEQUAL) Then
            fKey = False
        ElseIf (Asc(Mid$(sNames, lCount, 1)) = 0) Or (Len(sNames) = lCount) Then
            If Len(sNames) = lCount Then
                If fKey Then
                    sKEY = sKEY & Mid$(sNames, lCount, 1)
                Else
                    sValue = sValue & Mid$(sNames, lCount, 1)
                End If
            End If
            If Len(sKEY) <> 0 Then
                Call WritePrivateProfileString(sGroupName, sKEY, sValue, gstrSetupInfoFile)
            End If
            sKEY = vbNullString
            sValue = vbNullString
            fKey = True
        Else
            If fKey Then
                sKEY = sKEY & Mid$(sNames, lCount, 1)
            Else
                sValue = sValue & Mid$(sNames, lCount, 1)
            End If
        End If
    Next
    Call WritePrivateProfileString(strsection, gsGROUP & CStr(index), sGroupName, gstrSetupInfoFile)
End Sub
'-----------------------------------------------------------
' SUB: GetPrivate
'
' Returns the the value of whether the group is private specified by Index
'-----------------------------------------------------------
'
Function GetPrivate(ByVal strsection As String, ByVal index As Integer) As Boolean
    GetPrivate = CBool(ReadIniFile(gstrSetupInfoFile, strsection, gsPRIVATE & CStr(index)))
End Function
Function GetStart(ByVal strsection As String, ByVal index As Integer) As Boolean
    GetStart = ReadIniFile(gstrSetupInfoFile, strsection, gsPARENT & CStr(index)) = gsSTARTMENUKEY
End Function

'-----------------------------------------------------------
' SUB: CountIcons
'
' Determines how many icons must be installed by counting
' them in the setup information file (SETUP.LST)
'-----------------------------------------------------------
'
Function CountIcons(ByVal strsection As String) As Integer
    Dim intIdx As Integer
    Dim cIcons As Integer
    Dim sGroup As String
    Dim oCol As New Collection
    
    intIdx = 0
    cIcons = 0
    Do
        sGroup = ReadIniFile(gstrSetupInfoFile, strsection, gsGROUP & CStr(intIdx))
        If sGroup <> vbNullString Then 'Found a group
            oCol.Add sGroup
            intIdx = intIdx + 1
        Else
            Exit Do
        End If
    Loop
    Dim sGName As String
    Dim vGroup As Variant
    For Each vGroup In oCol
        intIdx = 1
        Do
            sGName = ReadIniFile(gstrSetupInfoFile, vGroup, gsICON & CStr(intIdx))
            If sGName <> vbNullString Then
                cIcons = cIcons + 1
                intIdx = intIdx + 1
            Else
                Exit Do
            End If
        Loop
    Next
    CountIcons = cIcons
    
End Function
'-----------------------------------------------------------
' SUB: CreateIcons
'
' Walks through the list of files in SETUP.LST and creates
' Icons in the Program Group for files needed it.
'-----------------------------------------------------------
'
Sub CreateIcons(ByVal strsection As String)
    Dim intIdx As Integer
    Dim sFile As FILEINFO
    Dim strProgramIconTitle As String
    Dim strProgramIconCmdLine As String
    Dim strProgramPath As String
    Dim strProgramArgs As String
    Dim intAnchor As Integer
    Dim intOffset As Integer
    Dim strGroup As String
    Dim sGroup As String
    Dim oCol As New Collection
    Const CompareBinary = 0
    '
    'For each file in the specified section, read info from the setup info file
    '
    intIdx = 0
    Do
        sGroup = ReadIniFile(gstrSetupInfoFile, strsection, gsGROUP & CStr(intIdx))
        If sGroup <> vbNullString Then 'Found a group
            oCol.Add sGroup
            intIdx = intIdx + 1
        Else
            Exit Do
        End If
    Loop
    Dim sGName As String
    Dim vGroup As Variant
    For Each vGroup In oCol
        intIdx = 0
        Do
            intIdx = intIdx + 1
            sGName = ReadIniFile(gstrSetupInfoFile, vGroup, gsICON & CStr(intIdx))
            If sGName <> vbNullString Then
                '
                ' Get the Icon's caption and command line
                '
                strProgramIconTitle = ReadIniFile(gstrSetupInfoFile, vGroup, gsTITLE & CStr(intIdx))
                strProgramIconCmdLine = ReadIniFile(gstrSetupInfoFile, vGroup, gsICON & CStr(intIdx))
                strGroup = vGroup
                '
                ' if the ProgramIcon is specified, then we create an icon,
                ' otherwise we don't.
                '
                If Trim(strUnQuoteString(strProgramIconTitle)) <> vbNullString Then
                    '
                    ' If the command line is not specified in SETUP.LST and the icon
                    ' is, then use the files destination path as the command line.  In
                    ' this case there are no parameters.
                    '
                    If Trim(strUnQuoteString(strProgramIconCmdLine)) = "" Then
                        strProgramPath = sFile.strDestDir & gstrSEP_DIR & sFile.strDestName
                        strProgramArgs = ""
                    Else
                        '
                        ' Parse the command line, to determine what is the exe, etc. and what
                        ' are the parameters.  The first space that is not contained within
                        ' quotes, marks the end of the exe, etc..  Everything afterwards are
                        ' parameters/arguments for the exe.  NOTE: It is important that if
                        ' the exe is contained within quotes that the parameters not be
                        ' contained within the same quotes.  The arguments can themselves
                        ' each be inside quotes as long as they are not in the same quotes
                        ' with the exe.
                        '
                        intAnchor = 1
                        intOffset = intGetNextFldOffset(intAnchor, strProgramIconCmdLine, " ", CompareBinary)
                        If intOffset = 0 Then intOffset = Len(strProgramIconCmdLine) + 1
                        strProgramPath = Trim(Left(strProgramIconCmdLine, intOffset - 1))
                        '
                        ' Got the exe, now the parameters.
                        '
                        strProgramArgs = Trim(Mid(strProgramIconCmdLine, intOffset + 1))
                    End If
                    '
                    ' Expand all the Destination Directory macros that are embedded in the
                    ' Program Path and the Arguments'
                    '
                    strProgramPath = ResolveDestDir(strProgramPath)
                    strProgramArgs = ResolveDestDirs(strProgramArgs)
                    '
                    ' Finally, we have everything we need, create the icon.
                    '
                    Dim fPrivate As Boolean, sParent As String
                    Dim intIdx2 As Integer
                    
                    intIdx2 = 0
                    Do
                        sGroup = ReadIniFile(gstrSetupInfoFile, gsICONGROUP, gsGROUP & CStr(intIdx2))
                        If sGroup = strGroup Then 'Found the group
                            If IsWindows95 Then
                                fPrivate = True
                            Else
                                fPrivate = GetPrivate(gsICONGROUP, intIdx2)
                            End If
                            If GetStart(gsICONGROUP, intIdx2) Then
                                sParent = gsSTARTMENUKEY
                            Else
                                sParent = gsPROGMENUKEY
                            End If
                            Exit Do
                        End If
                        intIdx2 = intIdx2 + 1
                    Loop
                    CreateOSLink frmSetup1, strGroup, strProgramPath, strProgramArgs, strProgramIconTitle, fPrivate, sParent
                ElseIf Trim(strUnQuoteString(strProgramIconCmdLine)) <> vbNullString Then
                    '
                    ' This file contained specified a command line in SETUP.LST but no icon.
                    ' This is an error.  Let the user know and skip this icon or abort.
        
                    '
                    If gfNoUserInput Or MsgWarning(ResolveResString(resICONMISSING, "|1", sFile.strDestName), vbYesNo Or vbExclamation, gstrSETMSG) = vbNo Then
                        ExitSetup frmSetup1, gintRET_FATAL
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
    Next
End Sub

Public Function RebootSystem() As Boolean
    Dim ret As Long
    Dim hToken As Long
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpOld As TOKEN_PRIVILEGES
    Dim fOkReboot As Boolean
    Const sSHUTDOWN As String = "SeShutdownPrivilege"
    'Check to see if we are running on Windows NT
    If IsWindowsNT() Then
        'We are running windows NT.  We need to do some security checks/modifications
        'to ensure we have the token that allows us to reboot.
        If OpenProcessToken(GetCurrentProcess(), _
                TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) Then
            ret = LookupPrivilegeValue(vbNullString, sSHUTDOWN, tkp.Privileges(0).pLuid)
            tkp.PrivilegeCount = 1
            tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
            fOkReboot = AdjustTokenPrivileges(hToken, 0, tkp, LenB(tkpOld), tkpOld, ret)
        End If
    Else
        'We are running Win95/98.  Nothing needs to be done.
        fOkReboot = True
    End If
    If fOkReboot Then RebootSystem = (ExitWindowsEx(EWX_REBOOT, 0) <> 0)
End Function

Private Function GetFileTime(ByVal aDate As Date) As FileTime
    Dim lTemp As SYSTEMTIME
    Dim lTime As FileTime
    
    VariantTimeToSystemTime aDate, lTemp
    SystemTimeToFileTime lTemp, lTime
    LocalFileTimeToFileTime lTime, GetFileTime
End Function

