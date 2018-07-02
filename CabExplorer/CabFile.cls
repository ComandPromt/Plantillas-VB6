VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CABFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrFileToExtract  As String
Private mstrOutputPath     As String
Private mstrOutputFile     As String
Private mlngCount          As Long
Private Const conXMLHeader As String = "<?xml version=""1.0""?>"

' Cab file to work with.
Public CabName As String
'
' ==================================
' Events raised by this class.
' ==================================
'
Public Event FileFound( _
        ByVal FileName As String, _
        ByVal DateTime As Date, _
        ByVal Size As Variant, _
        ByVal Path As String)

Public Event BeforeExtract( _
        ByVal FileName As String, _
        ByRef Cancel As Boolean)

Public Event AfterExtract( _
        ByVal FileName As String)
'
' ==========================================
' Name of the class, for error messages, and
' a base for user-defined error values (of
' which there aren't many, in this class).
' ==========================================
'
Private Const conClass As String = "CabFile"
Private Const conErrBase As Long = vbObjectError + 1956

Public Enum Errors
    errNoCabFile = conErrBase + 0
End Enum

Private Const conErrNoCabFile = _
    "You must set the CabName property before " & _
    "taking any action on the CabFile object."

Private Const conErrUnknown = "Unknown error."
'
' ==================================
' <XMLTags>
' ==================================
'
Private Const conXMLFile = "FILE"
Private Const conXMLName = "NAME"
Private Const conXMLDate = "DATE"
Private Const conXMLSize = "SIZE"
Private Const conXMLFullName = "FULLNAME"
Private Const conXMLPath = "PATH"
Private Const conXMLTop = "CABFILE"
Private Const conXMLCabFile = "CABFILENAME"
Private Const conXMLFileCount = "FILECOUNT"
'
' ==================================
' </XMLTags>
' ==================================
'
' Output XML string. See the GetXML method.
Private mstrXML As String

'
' Notification messages, handled in the callback
' procedure. This class doesn't handle them all.
'
Private Const SPFILENOTIFY_FILEINCABINET = &H11
Private Const SPFILENOTIFY_NEEDNEWCABINET = &H12
Private Const SPFILENOTIFY_FILEEXTRACTED = &H13

'
' Instructions sent out of the callback procedure.
' Tells Windows what to do next.
'
Private Enum FILEOP
    FILEOP_ABORT = 0
    FILEOP_DOIT = 1
    FILEOP_SKIP = 2
End Enum

'
' Local enum, indicating what action to
' take on each pass through the callback
' procedure.
'
Private Enum SetupIterateCabinetActions
    sicCount
    sicReport
    sicExtract
    sicGetXML
End Enum

'
' ==================================
' API Declarations
' ==================================
'
Private Const NO_ERROR = 0

Private Type FILEPATHS
    Target     As Long
    Source     As Long
    Win32Error As Integer
    Flags      As Long
End Type

Private Type SYSTEMTIME
    intYear         As Integer
    intMonth        As Integer
    intDayOfWeek    As Integer
    intDay          As Integer
    intHour         As Integer
    intMinute       As Integer
    intSecond       As Integer
    intMilliseconds As Integer
End Type

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" _
        (ByVal lpString As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function SetupIterateCabinet Lib "setupapi.dll" _
        Alias "SetupIterateCabinetA" (ByVal CabinetFile As String, _
        ByVal Reserved As Long, ByVal MsgHandler As Long, _
        ByVal Context As Long) As Long

Private Declare Function DosDateTimeToFileTime Lib "kernel32" _
        (ByVal wFatDate As Long, ByVal wFatTime As Long, _
        lpFileTime As FILETIME) As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" _
        (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" _
        (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long


Public Property Get FileCount() As Long
Dim lngReturn As Long
    
    On Error GoTo HandleErrors
        
    If Len(CabName) = 0 Then
        Err.Raise errNoCabFile, _
        conClass, fGetError(errNoCabFile)
    Else
        Call SetCabFile(Me)
        lngReturn = SetupIterateCabinet(CabName, 0, AddressOf CabinetCallback, sicCount)
        If lngReturn = 0 Then
            '
            ' If the return value is 0, the
            ' call to SetupIterateCabinet failed.
            ' Raise the error back to the caller,
            ' and convert the error to appropriate
            ' text, if possible. fErrToText doesn't
            ' catch all possible errors, but gets
            ' many of them.
            '
            Err.Raise Err.LastDllError, conClass, fErrToText(Err.LastDllError)
        End If
    End If

    FileCount = mlngCount

NormalExit:
    Exit Function

HandleErrors:
    Err.Raise Err.Number, _
     Err.Source, Err.Description
    Resume NormalExit
End Property

Public Function Extract(Optional FileToExtract As String = vbNullString, _
        Optional OutputPath As String = vbNullString, _
        Optional OutputFile As String = vbNullString) As Boolean
    
Dim lngReturn As Long
        
    '
    ' Extracts the file(s) from the cabinet. FileToExtract can specify
    ' the file to extract or if ommitted all files will be extracted.
    ' OutputPath can specify the folder to extract to. The default is the
    ' same folder as the cab file. When extracting a single file,
    ' OutputFile can specify the extract file name. The default is the
    ' original file name.
    '
    On Error GoTo HandleErrors
    
    If Len(CabName) = 0 Then
        Err.Raise errNoCabFile, _
            conClass, fGetError(errNoCabFile)
    Else
        Call SetCabFile(Me)
        '
        ' Set up the module-level variables
        ' tracking which file(s) you want to extract,
        ' and where you want to put them.
        '
        mstrFileToExtract = FileToExtract
        mstrOutputPath = OutputPath
        mstrOutputFile = OutputFile
        
        lngReturn = SetupIterateCabinet(CabName, 0, _
                AddressOf CabinetCallback, sicExtract)
                
        If lngReturn = 0 Then
            '
            ' If the return value is 0, the
            ' call to SetupIterateCabinet failed.
            ' Raise the error back to the caller,
            ' and convert the error to appropriate
            ' text, if possible. fErrToText doesn't
            ' catch all possible errors, but gets
            ' many of them.
            '
            Err.Raise Err.LastDllError, _
                conClass, fErrToText(Err.LastDllError)
        End If
    End If
    Extract = (lngReturn <> 0)

NormalExit:
    Exit Function

HandleErrors:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume NormalExit
    Resume
End Function
    
Public Function GetInfo(Optional FileToInvestigate As String = vbNullString) As Boolean
Dim lngReturn As Long
    '
    ' Iterate through all the files in the cab file,
    ' raising the FileFound event for each found file.
    ' If you specify a value for FileToInvestigate,
    ' the event will only occur zero or one times.
    '
    On Error GoTo HandleErrors
    
    If Len(CabName) = 0 Then
        Err.Raise errNoCabFile, _
            conClass, fGetError(errNoCabFile)
    Else
        Call SetCabFile(Me)
        mstrFileToExtract = FileToInvestigate
        
        lngReturn = SetupIterateCabinet(CabName, 0, _
         AddressOf CabinetCallback, sicReport)
        If lngReturn = 0 Then
            '
            ' If the return value is 0, the
            ' call to SetupIterateCabinet failed.
            ' Raise the error back to the caller,
            ' and convert the error to appropriate
            ' text, if possible. fErrToText doesn't
            ' catch all possible errors, but gets
            ' many of them.
            '
            Err.Raise Err.LastDllError, _
                conClass, fErrToText(Err.LastDllError)
        End If
    End If
    
    GetInfo = (lngReturn <> 0)

ExitHere:
    Exit Function

HandleErrors:
    Err.Raise Err.Number, _
        conClass & ".GetInfo", Err.Description
    Resume ExitHere
End Function
    
Public Function GetXML(Optional FileToInvestigate As String = vbNullString) As String
Dim lngReturn As Long
    '
    ' Similar to the GetInfo method, except it
    ' returns a simple XML stream describing the
    ' file(s) found.
    '
    On Error GoTo HandleErrors

    If Len(CabName) = 0 Then
        Err.Raise errNoCabFile, _
            conClass, fGetError(errNoCabFile)
    Else
        Call SetCabFile(Me)
        mstrFileToExtract = FileToInvestigate

        mstrXML = vbNullString
        lngReturn = SetupIterateCabinet(CabName, 0, AddressOf CabinetCallback, sicGetXML)
        '
        ' The return value will be 0 if the
        ' callback function failed.
        '
        If lngReturn = 0 Then
            Err.Raise Err.LastDllError, _
                conClass, fErrToText(Err.LastDllError)
        End If
    End If
    mstrXML = fBuildXMLElement(mstrXML, "FILES")
    GetXML = conXMLHeader & fBuildXMLElement(mstrXML, conXMLTop, conXMLCabFile, CabName)

NormalExit:
    Exit Function

HandleErrors:
    Err.Raise Err.Number, conClass & ".GetXML", Err.Description
    Resume NormalExit
End Function
    
Friend Function CabCallBack(ByVal InstallData As Long, _
        ByVal Notification As Long, Param1 As FileInCabinetInfo, _
        ByVal Param2 As Long) As Long
            
Dim fp As FILEPATHS

    '
    ' Callback procedure for SetupIterateCabinet. This procedure
    ' is called by a corresponding procedure in a standard module.
    '
    On Error GoTo HandleErrors

    '
    ' Handle the callback for the CAB file.
    '
    Select Case Notification
        Case SPFILENOTIFY_NEEDNEWCABINET
            ' Not handled here.
            CabCallBack = NO_ERROR
            
        Case SPFILENOTIFY_FILEEXTRACTED
            '
            ' Copy the bytes passed into a FILEPATHS structure.
            ' Although this procedure gets a parameter of
            ' type FileCabinetInfo, you want to cast it as a
            ' FILEPATHS structure. The LSET statement does that
            ' for you. You can also use the CopyMemory API function,
            ' but this is simpler.
            '
            LSet fp = Param1
            
            If fp.Win32Error = NO_ERROR Then
                RaiseEvent AfterExtract(fStringFromPointer(fp.Target))
            End If
            CabCallBack = fp.Win32Error
        
        Case SPFILENOTIFY_FILEINCABINET
            Select Case InstallData
                Case sicCount
                    '
                    ' Counting? Increment the private counter
                    ' variable each time, and tell Windows
                    ' to skip further processing for the file.
                    '
                    mlngCount = mlngCount + 1
                    CabCallBack = FILEOP_SKIP
                
                Case sicReport
                    CabCallBack = fHandleReport(Param1)
                    
                Case sicGetXML
                    CabCallBack = fHandleXML(Param1)
                    
                Case sicExtract
                    CabCallBack = fHandleExtract(Param1)
            End Select
    End Select

NormalExit:
    Exit Function

HandleErrors:
    Err.Raise Err.Number, conClass & ".CabCallBack", Err.Description
    Resume NormalExit
End Function

Private Function fHandleReport(fici As FileInCabinetInfo) As Long
Dim foAction   As FILEOP
Dim blnCancel  As Boolean
Dim blnDoIt    As Boolean
Dim strFile    As String
Dim ft         As FILETIME
Dim dtm        As Date
Dim strPath    As String
Dim strFull    As String
    '
    ' Assume you want to keep processing.
    '
    foAction = FILEOP_SKIP
    '
    ' Assume you don't want to process the
    ' current file.
    '
    blnDoIt = False
    '
    ' Get the file name passed to this callback function.
    '
    strFull = fStringFromPointer(fici.NameInCabinet)
    Call fSplitFile(strFull, strPath, strFile)
    '
    ' If you haven't specified a file to investigate,
    ' investigate them all.
    '
    If Len(mstrFileToExtract) = 0 Then
        blnDoIt = True
    Else
        '
        ' Otherwise, check to see if this is the file you
        ' want extracted.
        '
        If StrComp(strFull, mstrFileToExtract, vbTextCompare) = 0 Then
            '
            ' If you found a match for the one file
            ' you're interested in, abort processing afterwards.
            '
            foAction = FILEOP_ABORT
            blnDoIt = True
        End If
    End If
    
    If blnDoIt Then
        '
        ' Retrieve the file date information.
        '
        With fici
            Call DosDateTimeToFileTime(.DosDate, .DosTime, ft)
            dtm = fFileTimeToVBATime(ft, False)
        End With
        
        
        RaiseEvent FileFound(strFile, dtm, fici.FileSize, strPath)
    End If
    fHandleReport = CLng(foAction)
End Function

Private Function fHandleXML(fici As FileInCabinetInfo) As Long
Dim foAction   As FILEOP
Dim blnCancel  As Boolean
Dim blnDoIt    As Boolean
Dim strFile    As String
Dim ft         As FILETIME
Dim dtm        As Date
Dim strPath    As String
Dim strFull    As String

    On Error GoTo HandleErrors
    '
    ' Assume you want to keep processing.
    '
    foAction = FILEOP_SKIP
    '
    ' Assume you don't want to process the
    ' current file.
    '
    blnDoIt = False
    '
    ' Get the file name passed to this callback function.
    '
    strFull = fStringFromPointer(fici.NameInCabinet)
    Call fSplitFile(strFull, strPath, strFile)
    '
    ' If you haven't specified a file to investigate,
    ' investigate them all.
    '
    If Len(mstrFileToExtract) = 0 Then
        blnDoIt = True
    Else
        '
        ' Otherwise, check to see if this is the file you
        ' want extracted.
        '
        If StrComp(strFile, mstrFileToExtract, vbTextCompare) = 0 Then
            '
            ' If you found a match for the one file
            ' you're interested in, abort processing
            ' afterwards.
            '
            foAction = FILEOP_ABORT
            blnDoIt = True
        End If
    End If
    If blnDoIt Then
        '
        ' Retrieve the file date information.
        '
        With fici
            Call DosDateTimeToFileTime(.DosDate, .DosTime, ft)
            dtm = fFileTimeToVBATime(ft, False)
        End With
        mstrXML = mstrXML & _
         fBuildXMLElement( _
          fBuildXMLElement(strFull, conXMLFullName) & _
          fBuildXMLElement(strFile, conXMLName) & _
          fBuildXMLElement(CStr(dtm), conXMLDate) & _
          fBuildXMLElement(CStr(fici.FileSize), conXMLSize) & _
          fBuildXMLElement(strPath, conXMLPath), conXMLFile) & vbCrLf
    End If
    fHandleXML = CLng(foAction)

NormalExit:
    Exit Function

HandleErrors:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume NormalExit
End Function

Private Function fHandleExtract(fici As FileInCabinetInfo) As Long
Dim foAction  As FILEOP
Dim blnCancel As Boolean
Dim blnDoIt   As Boolean
Dim strFile   As String
Dim strPath   As String
Dim strFull   As String
    
    On Error GoTo HandleErrors
    '
    ' Assume you want to keep processing without
    ' extracting the file.
    '
    foAction = FILEOP_SKIP
    '
    ' Assume you don't want to process the
    ' current file at all.
    '
    blnDoIt = False
    '
    ' Get the file that's been found in the CAB file.
    '
    strFile = fStringFromPointer(fici.NameInCabinet)
    '
    ' If you haven't specified a file to extract,
    ' extract them all.
    '
    If Len(mstrFileToExtract) = 0 Then
        blnDoIt = True
    Else
        '
        ' Otherwise, check to see if this is the file you
        ' want extracted.
        '
        If StrComp(strFile, mstrFileToExtract, _
         vbTextCompare) = 0 Then
            '
            ' If you found a match for the one file
            ' you're interested in, abort processing
            ' afterwards.
            '
            blnDoIt = True
        End If
    End If
    
    If blnDoIt Then
        RaiseEvent BeforeExtract(strFile, blnCancel)
        If Not blnCancel Then
            strFull = fCalcOutputFile(strFile)
            Call fSplitFile(strFull, strPath, strFile)
            '
            ' Better make sure the path exists
            ' before attempting to create the file.
            '
            If fMakePath(strPath) Then
                Call pCopyStringToArray(fici.FullTargetName, strFull)
                foAction = FILEOP_DOIT
            End If
        End If
    End If
    fHandleExtract = CLng(foAction)
    
NormalExit:
    Exit Function

HandleErrors:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume NormalExit
End Function

Private Function fSysTimeToVBATime(stSysTime As SYSTEMTIME) As Date
    '
    ' Converts Windows SYSTEMTIME to VBA date/time.
    '
    ' Consrtuct a VBA date/time value using the
    ' DateSerial and TimeSerial functions.
    '
    With stSysTime
        fSysTimeToVBATime = DateSerial(.intYear, .intMonth, .intDay) + _
                        TimeSerial(.intHour, .intMinute, .intSecond)
    End With
End Function

Private Function fFileTimeToVBATime(ftFileTime As FILETIME, _
 Optional fLocal As Boolean = True) As Date

Dim stSystem        As SYSTEMTIME
Dim ftLocalFileTime As FILETIME
    '
    ' Converts Windows FILETIME date/time value to VBA.
    '
    ' If the user wants local time, convert the file
    ' time to local file time.
    '
    If fLocal Then
        Call FileTimeToLocalFileTime(ftFileTime, ftLocalFileTime)
        ftFileTime = ftLocalFileTime
    End If
    '
    ' Convert the file time to system time then
    ' call our own function to convert to VBA time
    '
    If CBool(FileTimeToSystemTime(ftFileTime, stSystem)) Then
        fFileTimeToVBATime = fSysTimeToVBATime(stSystem)
    End If
End Function

Private Function fStringFromPointer(ByVal ptr As Long) As String
Dim lngLen   As Long
Dim strBuffer As String
    '
    ' Given a string pointer, copy the value
    ' of the string into a new, safe location.
    '
    lngLen = lstrlen(ptr)
    strBuffer = Space(lngLen)
    Call CopyMemory(ByVal strBuffer, ByVal ptr, lngLen)
    fStringFromPointer = strBuffer
End Function

Private Sub pCopyStringToArray(bytArray() As Byte, strValue As String)
Dim i         As Long
Dim bytTemp() As Byte
    '
    ' bytArray is a fixed-sized array of bytes.
    ' You can't copy text directly into a fixed-
    ' sized array, but you can do it directly
    ' into a dynamic array. Therefore, copy
    ' the string into a dynamic array, and then
    ' copy it byte-for-byte into the fixed array.
    '
    ' Convert to ANSI, in a byte array.
    '
    bytTemp = StrConv(strValue, vbFromUnicode)
    '
    ' Copy the byte array to the output location.
    '
    For i = LBound(bytTemp) To UBound(bytTemp)
        bytArray(i + 1) = bytTemp(i)
    Next i
    '
    ' Tack on the extra null character.
    '
    bytArray(i + 1) = 0
End Sub

Private Function fCalcOutputFile(strFileFound As String) As String
Dim strPath As String
Dim strFile As String
Dim strOut  As String
    '
    ' strFile is the name of the file, found in the CAB file.
    ' Given the values of mstrFileToExtract, mstrOutputPath,
    ' mstrOutputFile, return the full path of the output file.
    ' If mstrFileToExtract is empty, then disregard mstrOutputFile,
    ' because you'll be extracting all the files.
    '
    ' Calculate the output path. Either use mstrOutputPath if it
    ' exists, or the CAB file's path if it doesn't.
    '
    If Len(mstrOutputPath) > 0 Then
        strPath = mstrOutputPath
    Else
        strPath = fGetPath(CabName)
    End If
    '
    ' Calculate the output file name. If mstrOutputFile exists,
    ' use it. If not, use the original name of the file.
    '
    If Len(mstrOutputFile) > 0 Then
        strFile = mstrOutputFile
    Else
        strFile = strFileFound
    End If
    
    fCalcOutputFile = fFixPath(strPath) & strFile
End Function

Private Function fGetPath(strFile As String) As String
Dim lngPos As Long
    '
    ' Given a file name with a path, pull off the path part.
    '
    
    lngPos = InStrRev(strFile, "\")
    If lngPos > 0 Then
        fGetPath = Left$(strFile, lngPos)
    Else
        fGetPath = ""
    End If
End Function

Private Function fErrToText(lngErr As Long) As String
Dim strOut As String
    '
    ' Given a Windows error number, convert to text.
    ' Only handles the most common errors.
    '
    Select Case lngErr
        Case 2
            strOut = "The system cannot find the file specified."
        
        Case 3
            strOut = "The system cannot find the path specified."
            
        Case 4
            strOut = "The system cannot open the file."
            
        Case 5
            strOut = "Access is denied."
        
        Case 8
            strOut = "Not enough storage is available to process this command."
            
        Case 13
            strOut = "Invalid data."
        
        Case 14
            strOut = "Not enough storage is available to complete this operation."
            
        Case 15
            strOut = "The system cannot find the drive specified."
        
        Case 19
            strOut = "The media is write protected."
            
        Case 20
            strOut = "The system cannot find the device specified."
            
        Case 21
            strOut = "The device is not ready."
            
        Case 23
            strOut = "Data error (cyclic redundancy check)."
            
        Case 25
            strOut = "The drive cannot locate a specific area or track on the disk."
            
        Case 26
            strOut = "The specified disk or diskette cannot be accessed."
            
        Case 27
            strOut = "The drive cannot find the sector requested."
        
        Case 29
            strOut = "The system cannot write to the specified device."
        
        Case 30
            strOut = "The system cannot read from the specified device."
            
        Case 31
            strOut = "A device attached to the system is not functioning."
        
        Case 32
            strOut = "The process cannot access the file because it is being used by another process."
            
        Case 33
            strOut = "The process cannot access the file because another process has locked a portion of the file."
            
        Case 39
            strOut = "The disk is full."
            
        Case 82
            strOut = "The directory or file cannot be created."
            
        Case 111
            strOut = "The file name is too long."
            
        Case 112
            strOut = "There is not enough space on the disk."
            
        Case 123
            strOut = "The filename, directory name, or volume label syntax is incorrect."
        
        Case Else
            strOut = "Unknown error."
    End Select
    fErrToText = strOut
End Function

Private Function fGetError(lngErr As Long) As String
    Dim strOut As String
    '
    ' Return an error message for an internal error.
    ' Add more to this SELECT CASE, if you need more.
    '
    Select Case lngErr
        Case errNoCabFile
            strOut = conErrNoCabFile
        Case Else
            strOut = conErrUnknown
    End Select
    fGetError = strOut
End Function

Private Function fBuildXMLElement(strValue As String, _
        strTag As String, Optional strAttributeName As String, _
        Optional strAttributeValue As String) As String

Dim strOut As String
    '
    ' Given a piece of text ("HELLO", for example) and a tag
    ' ("VALUE", for example), return a valid XML element:
    ' <VALUE>Hello</VALUE>
    '
    ' You can optionally specify a single attribute value.
    '
    strOut = "<" & strTag
    
    If Len(strAttributeName) > 0 Then
        strOut = strOut & " " & strAttributeName & " = '" & strAttributeValue & "'"
    End If
    strOut = strOut & ">" & strValue & "</" & strTag & ">"
    
    fBuildXMLElement = strOut
End Function

