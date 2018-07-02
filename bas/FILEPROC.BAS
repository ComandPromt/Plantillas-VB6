Attribute VB_Name = "Module2"
Option Explicit

Public Const conNumBoxes = 5
Public Const conSaveFile = 1, conLoadFile = 2
Public Const conReplaceFile = 1, conReadFile = 2, conAddToFile = 3
Public Const conRandomFile = 4, conBinaryFile = 5

Public Const errDeviceUnavailable = 68
Public Const errDiskNotReady = 71, errFileAlreadyExists = 58
Public Const errTooManyFiles = 67, errRenameAcrossDisks = 74
Public Const errPathFileAccessError = 75, errDeviceIO = 57
Public Const errDiskFull = 61, errBadFileName = 64
Public Const errBadFileNameOrNumber = 52, errFileNotFound = 53
Public Const errPathDoesNotExist = 76, errBadFileMode = 54
Public Const errFileAlreadyOpen = 55, errInputPastEndOfFile = 62

Function FileErrors(errVal As Integer) As Integer
Dim MsgType As Integer
Dim Response As Integer
Dim Action As Integer
Dim Msg As String
MsgType = vbExclamation
Select Case errVal
    Case errDeviceUnavailable                       ' Error #68
        Msg = "That device appears to be unavailable."
        MsgType = vbExclamation + 5
    Case errDiskNotReady                            ' Error #71
        Msg = "The disk is not ready."
    Case errDeviceIO
        Msg = "The disk is full."
    Case errBadFileName, errBadFileNameOrNumber     ' Errors #64 & 52
        Msg = "That filename is illegal."
    Case errPathDoesNotExist                        ' Error #76
        Msg = "That path doesn't exist."
    Case errBadFileMode                             ' Error #54
        Msg = "Can't open your file for that type of access."
    Case errFileAlreadyOpen                         ' Error #55
        Msg = "That file is already open."
    Case errInputPastEndOfFile                      ' Error #62
    Msg = "This file has a nonstandard end-of-file marker,"
    Msg = Msg + "or an attempt was made to read beyond "
    Msg = Msg + "the end-of-file marker."
    Case Else
        FileErrors = 3
        Exit Function
    End Select
    Response = MsgBox(Msg, MsgType, "File Error")
    Select Case Response
        Case 4          ' Retry button.
            FileErrors = 0
        Case 5          ' Ignore button.
            FileErrors = 1
        Case 1, 2, 3    ' OK and Cancel buttons.
            FileErrors = 2
        Case Else
            FileErrors = 3
    End Select
End Function

Function FileOpener(NewFileName As String, Mode As Integer, RecordLen As Integer, Confirm As Integer) As Integer
     Dim NewFileNum As Integer
     Dim Action As Integer
     Dim FileExists As Integer
     Dim Msg As String
     On Error GoTo OpenerError
     If NewFileName Like "*[;-?[* ]*" Or NewFileName Like "*]*" Then Error errBadFileName
     If Confirm Then
        If Dir(NewFileName) = "" Then
            FileExists = False
        Else
            FileExists = True
        End If
        If Mode = conReplaceFile And FileExists Then
            Msg = "Replace contents of " + NewFileName + "?"
            If MsgBox(Msg, 49, "Replace File?") = 2 Then
                FileOpener = 0
                Exit Function
            End If
        End If
        If Not FileExists Then
            Msg = "The file " + NewFileName + " does not exist. "
            Msg = Msg + "Do you want to create it?"
            If MsgBox(Msg, 1, "Create File?") = 2 Then
                FileOpener = 0
                Exit Function
            End If
        End If
     End If
     NewFileNum = FreeFile
     Select Case Mode
          Case conReplaceFile
            Open NewFileName For Output As NewFileNum
          Case conReadFile
            Open NewFileName For Input As NewFileNum
          Case conAddToFile
            Open NewFileName For Append As NewFileNum
          Case conRandomFile
            Open NewFileName For Random As NewFileNum Len = RecordLen
          Case conBinaryFile
            Open NewFileName For Binary As NewFileNum
          Case Else
            Exit Function
     End Select
     FileOpener = NewFileNum
Exit Function
OpenerError:
     Action = FileErrors(Err)
     Select Case Action
        Case 0
            Resume
        Case Else
            FileOpener = 0
            Exit Function
     End Select
End Function

Function GetFileName(Prompt As String) As String
    GetFileName = LTrim(RTrim(UCase(InputBox(Prompt, "Enter File Name"))))
End Function

