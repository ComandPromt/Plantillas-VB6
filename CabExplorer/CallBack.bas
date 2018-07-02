Attribute VB_Name = "Callback"
Option Explicit

Private CabFile As CabFile
Private Const MAXPATH = 260

Public Type FileInCabinetInfo
    NameInCabinet As Long
    FileSize      As Long
    Win32Error    As Long
    DosDate       As Integer
    DosTime       As Integer
    DosAttribs    As Integer
    FullTargetName(1 To MAXPATH) As Byte
End Type

Public Function fFixPath(strPath As String) As String
    '
    ' Append a trailing "\" to a path, if necessary.
    '
    If Right$(strPath, 1) = "\" Then
        fFixPath = strPath
    Else
        fFixPath = strPath & "\"
    End If
End Function

Public Function fSplitFile(ByVal strFull As String, ByRef strPath As String, _
        ByRef strFile As String)

Dim lngPos As Long

    '
    ' Given a full path, parse it and return
    ' the path and file name.
    '
    lngPos = InStrRev(strFull, "\")
    If lngPos > 0 Then
        strPath = Left$(strFull, lngPos)
        strFile = Mid$(strFull, lngPos + 1)
    Else
        strPath = vbNullString
        strFile = strFull
    End If
End Function

Public Function fMakePath(strPath As String) As Boolean
Dim strItems() As String
Dim strTemp    As String
Dim lngUB      As Long
Dim lngLB      As Long
Dim i          As Long
Dim lngStop    As Long

    '
    ' Create the folders in the path that was passed in.
    '
    If Len(strPath) = 0 Then
        fMakePath = False
        GoTo NormalExit
    End If
    
    On Error Resume Next
    
    ' Attempt to create the path.
    ' If this returns no error or error 75,
    ' you're done. Otherwise, do the work.
    '
    ' Get right of trailing "\", if it's there.
    '
    If Right$(strPath, 1) = "\" Then
        strPath = Left$(strPath, Len(strPath) - 1)
    End If
    
    MkDir strPath
    
    Select Case Err.Number
        Case 76
            ' Path doesn't exist.
        Case 75
            ' Path exists already, get out.
            fMakePath = True
            GoTo NormalExit
        Case 0
            ' Folder created successfully.
            fMakePath = True
            GoTo NormalExit
        Case Else
            ' This shouldn't happen.
            fMakePath = False
            GoTo NormalExit
    End Select
    
    '
    ' Create an array full of all the items
    ' in the path, delimited with "\".
    '
    strItems = Split(strPath, "\")
    
    '
    ' Store away the lower and upper bounds.
    '
    lngLB = LBound(strItems)
    lngUB = UBound(strItems)
    
    ' You've already determined that you cannot
    ' create the path, given all the items. That is,
    ' if the path is C:\a\b\c\d\e, you know that
    ' you cannot create the path with the "e" on there.
    ' Therefore, this loop works its way backwards, looking for
    ' the longest path that either exists, or that you
    ' can create, without error.
    '
    ' Once you've found or created a path, the rest of
    ' the code works the other direction--adds on the
    ' path items, creating folders, until you get them
    ' all created, or trigger a run-time error.
    '
    ' You're going to loop from the next-to-last item
    ' back to the start, attempting to create
    ' or locate the path.
    '
    lngStop = lngUB
    For lngStop = lngUB - 1 To lngLB Step -1
        Err.Clear
        strTemp = vbNullString
        ' Build up the path to be tested.
        For i = lngLB To lngStop
            strTemp = strTemp & "\" & strItems(i)
        Next i
        ' Remove the leading "\".
        If Len(strTemp) > 1 Then
            strTemp = Mid$(strTemp, 2)
        End If
        '
        ' Attempt to create the folder.
        ' This could succeed (error 0),
        ' fail because the folder exists (error 75),
        ' or fail because some parent folder
        ' doesn't exist (error 76). If you get
        ' error 0 or 75, you're done.
        '
        MkDir strTemp
        Select Case Err.Number
            Case 0, 75
                ' Path created or it exists.
                Exit For
            Case 76 ' Path wasn't found.
            Case Else
                fMakePath = False
                GoTo NormalExit
        End Select
    Next lngStop
    '
    ' Starting where you left off when working
    ' backwards, attempt to create the folders
    ' working downwards. At any point, if you get
    ' an error, you're done.
    '
    For i = lngStop + 1 To lngUB
        Err.Clear
        strTemp = strTemp & "\" & strItems(i)
        MkDir strTemp
        If Err.Number <> 0 Then
            '
            ' You can't create the path. Return False.
            '
            fMakePath = False
            GoTo NormalExit
        End If
    Next i
    fMakePath = True
    
NormalExit:
    Exit Function
End Function

Public Sub SetCabFile(cab As CabFile)
    Set CabFile = cab
End Sub

Public Function CabinetCallback(ByVal Context As Long, _
        ByVal Notification As Long, ByRef Param1 As FileInCabinetInfo, _
        ByVal Param2 As Long) As Long
    
    If Not CabFile Is Nothing Then
        CabinetCallback = CabFile.CabCallBack(Context, Notification, Param1, Param2)
    End If

End Function

