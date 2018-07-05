Attribute VB_Name = "basInitEntry"
Option Explicit

Private Type FormPosition
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
    Maxed   As Boolean
End Type

'sDefInitFileName is setup as (AppPath\AppEXEName.Ini)
'and is used as the Default Initialization Filename
Private sDefInitFileName As String

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub AddRecentFile(ByVal sNewFileName As String, mnuRecent As Variant, Optional ByVal iMaxEntries As Integer = 8, Optional ByVal iMaxFileNameLen As Integer = 60)

Dim lRet        As Long
Dim iArrayCnt   As Integer
Dim iFileCnt    As Integer
Dim sFilename   As String
Dim saFiles()    As String

    ReDim saFiles(iMaxEntries)
    
    'Add New File at First Position
    saFiles(0) = sNewFileName
    
    'Get all Files in Init File
    iFileCnt = 1
    sFilename = GetInitEntry("Recent Files", "File " & CStr(iFileCnt), "")
    While Len(sFilename) > 0 And iArrayCnt < iMaxEntries
        'Don't get New File Again
        If LCase$(sFilename) <> LCase$(sNewFileName) Then
            iArrayCnt = iArrayCnt + 1
            saFiles(iArrayCnt) = sFilename
        End If
        iFileCnt = iFileCnt + 1
        sFilename = GetInitEntry("Recent Files", "File " & CStr(iFileCnt), "")
    Wend
    
    'Release Excess Memory
    ReDim Preserve saFiles(iArrayCnt)
    
    'Clean up the Init File (Deletes the Entire "Recent Files" Section)
    lRet = SetInitEntry("Recent Files")
    
    'Put Files Back Into Init File in Their New Order
    For iFileCnt = 0 To iArrayCnt
        lRet = SetInitEntry("Recent Files", "File " & CStr(iFileCnt + 1), saFiles(iFileCnt))
    Next iFileCnt
    
    'Retrieve Ordered Files Back Into Menu
    Call GetRecentFiles(mnuRecent, iMaxEntries, iMaxFileNameLen)
    
    'Checkmark First Recent File
    mnuRecent(0).Checked = (mnuRecent(0).Caption <> "(Empty)")
    
End Sub

Public Sub ClearRecentFiles(mnuRecent As Variant)

Dim lRet As Long

    'Clear out the Recent Files (Deletes the Entire "Recent Files" Section)
    lRet = SetInitEntry("Recent Files")

    'Clear the Menus
    Call GetRecentFiles(mnuRecent)

End Sub


Public Sub GetRecentFiles(mnuRecent As Variant, Optional ByVal iMaxEntries As Integer = 8, Optional ByVal iMaxFileNameLen As Integer = 60)

'mnuRecent Must Be a Menu Array. At Design Time, create
'the first mnuRecent(0) with the Caption set to "(Empty)"
'and Disable it.

Dim iIdx        As Integer
Dim iFileCnt    As Integer
Dim iFullCnt    As Integer
Dim iMenuCnt    As Integer
Dim sFilename   As String

    On Error GoTo LocalError
    
    'Get the Menu Count
    iMenuCnt = mnuRecent.UBound
    
    'Unload all but first Menu
    For iIdx = 1 To iMenuCnt
        Unload mnuRecent(iIdx)
    Next iIdx
    mnuRecent(0).Checked = False
    mnuRecent(0).Tag = ""
    mnuRecent(0).Enabled = False
    mnuRecent(0).Caption = "(Empty)"
    
    'Get First Entry In InitFile
    sFilename = GetInitEntry("Recent Files", "File " & CStr(iFullCnt + 1), "")
    While Len(sFilename) > 0 And iFileCnt <= iMaxEntries
        If Exists(sFilename) Then
            'Load Menu Item if Not First Item
            If iFileCnt > 0 Then
                Load mnuRecent(iFileCnt)
            End If
            'Create Menu Caption
            'ex. "&1 C:\DirName\DirName\FileName"
            mnuRecent(iFileCnt).Caption = "&" & CStr(iFileCnt + 1) & " " & _
                ShortenFileName(sFilename, iMaxFileNameLen)
            'Menu.Tag Contains Actual Filename.
            'Menu.Caption May Contain A Shortened Version Of It.
            mnuRecent(iFileCnt).Tag = sFilename
            mnuRecent(iFileCnt).Enabled = True
            mnuRecent(iFileCnt).Visible = True
            iFileCnt = iFileCnt + 1
        End If
        iFullCnt = iFullCnt + 1
        'Get Next Entry
        sFilename = GetInitEntry("Recent Files", "File " & CStr(iFullCnt + 1), "")
        'Loops If Next Entry Is Valid
    Wend

NormalExit:
    Exit Sub
    
LocalError:
    MsgBox Err.Description, vbExclamation, App.EXEName
    Resume NormalExit
    
End Sub

Public Function Exists(ByVal sFilename As String) As Boolean

Dim iFlags As VbFileAttribute

    iFlags = vbArchive Or vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem Or vbVolume
    
    If Len(Trim$(sFilename)) > 0 Then
        On Error Resume Next
        sFilename = Dir$(sFilename, iFlags)
        Exists = Err.Number = 0 And Len(sFilename) > 0
    Else
        Exists = False
    End If
    
End Function

Public Sub RemoveRecentFile(ByVal sRemoveFileName As String, mnuRecent As Variant, Optional ByVal iMaxEntries As Integer = 8, Optional ByVal iMaxFileNameLen As Integer = 60)

Dim lRet        As Long
Dim iArrayCnt   As Integer
Dim iFileCnt    As Integer
Dim sFilename   As String
Dim saFiles()    As String

    ReDim saFiles(iMaxEntries)
    
    'Get all Files in Init File
    iFileCnt = 1
    sFilename = GetInitEntry("Recent Files", "File " & CStr(iFileCnt), "")
    While Len(sFilename) > 0 And iArrayCnt < iMaxEntries
        'Don't get the File to be removed
        If LCase$(sFilename) <> LCase$(sRemoveFileName) Then
            saFiles(iArrayCnt) = sFilename
            iArrayCnt = iArrayCnt + 1
        End If
        iFileCnt = iFileCnt + 1
        sFilename = GetInitEntry("Recent Files", "File " & CStr(iFileCnt), "")
    Wend
    
    'Release Excess Memory
    ReDim Preserve saFiles(iArrayCnt - 1)
    
    'Clean up the Init File (Deletes the Entire "Recent Files" Section)
    lRet = SetInitEntry("Recent Files")
    
    'Put Files Back Into Init File Without the Removed File
    For iFileCnt = 0 To iArrayCnt - 1
        lRet = SetInitEntry("Recent Files", "File " & CStr(iFileCnt + 1), saFiles(iFileCnt))
    Next iFileCnt
    
    'Retrieve Ordered Files Back Into Menu
    Call GetRecentFiles(mnuRecent, iMaxEntries, iMaxFileNameLen)
    
End Sub

Private Function ShortenFileName(ByVal sFilename As String, ByVal iMaxLen As Integer) As String

Dim iLen        As Integer
Dim iSlashPos   As Integer

    On Error GoTo LocalError
    
    'If Filename Is Longer Than MaxLen
    If Len(sFilename) > iMaxLen Then
        'Make Room For "..."
        iLen = iMaxLen - 3
        'Find First "\"
        iSlashPos = InStr(sFilename, "\")
        'Loop Until Filename is Shorter Than MaxLen
        While (iSlashPos > 0) And (Len(sFilename) > iLen)
            sFilename = Mid$(sFilename, iSlashPos)
            'Find Next "\"
            iSlashPos = InStr(2, sFilename, "\")
        Wend
        'If No "\" Was Found (FailSafe - This Should Not Happen)
        If Len(sFilename) > iLen Then
            '"Very Long FileName" = "...ong Filename"
            sFilename = "..." & Mid$(sFilename, Len(sFilename) - iLen + 1)
        Else
            '"C:\Dir1\Dir2\Dir3\File" = "...\Dir2\Dir3\File"
            sFilename = "..." & sFilename
        End If
    
    End If
    
    'Set Return Filename
    ShortenFileName = sFilename

NormalExit:
    Exit Function
    
LocalError:
    MsgBox Err.Description, vbExclamation, App.EXEName
    Resume NormalExit

End Function

Public Function GetInitEntry(ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "", Optional ByVal sInitFileName As String = "") As String

'This Function Reads In a String From The Init File.
'Returns Value From Init File or sDefault If No Value Exists.
'sDefault Defaults to an Empty String ("").
'Creates and Uses sDefInitFileName (AppPath\AppEXEName.Ini)
'if sInitFileName Parameter Is Not Passed In.

Dim sBuffer As String
Dim sInitFile As String

    'If Init Filename NOT Passed In
    If Len(sInitFileName) = 0 Then
        'If Static Init FileName NOT Already Created
        If Len(sDefInitFileName) = 0 Then
            'Create Static Init FileName
            sDefInitFileName = App.Path
            If Right$(sDefInitFileName, 1) <> "\" Then
                sDefInitFileName = sDefInitFileName & "\"
            End If
            sDefInitFileName = sDefInitFileName & App.EXEName & ".ini"
        End If
        sInitFile = sDefInitFileName
    Else    'If Init Filename Passed In
        sInitFile = sInitFileName
    End If
    
    sBuffer = String$(2048, " ")
    GetInitEntry = Left$(sBuffer, GetPrivateProfileString(sSection, ByVal sKeyName, sDefault, sBuffer, Len(sBuffer), sInitFile))

End Function

Public Sub SaveFormSize(frmForm As Form, Optional ByVal sInitFileName As String = "")

'This routine saves the size and position of frmForm
'using only 1 line in the init file per form. The
'init section is "Positions" and the init key is the
'form's name. The data is a comma delimited string.
'Data order = Left, Top, Width, Height, Maximized

Dim lRet        As Long
Dim sData       As String
Dim saSizes()   As String

    ReDim saSizes(4)
    
    If frmForm.WindowState = vbNormal Then
        'These values would be wrong if the
        'form was minimized or maximized.
        saSizes(0) = CStr(frmForm.Left)
        saSizes(1) = CStr(frmForm.Top)
        saSizes(2) = CStr(frmForm.Width)
        saSizes(3) = CStr(frmForm.Height)
        saSizes(4) = "False"
    Else
        'If minimized or maximized, retrieve the
        'form's previously saved positions.
        sData = GetInitEntry("Positions", frmForm.Name, "", sInitFileName)
        If Len(sData) = 0 Then
            'No previously saved positions.
            'Set all sizes to -1.
            saSizes(0) = "-1"
            saSizes(1) = "-1"
            saSizes(2) = "-1"
            saSizes(3) = "-1"
        Else
            'Use previously saved positions.
            saSizes() = Split(sData, ",")
            'Ensure they were all there.
            ReDim Preserve saSizes(4)
        End If
        '
        saSizes(4) = CStr(frmForm.WindowState = vbMaximized)
    End If
        
    lRet = SetInitEntry("Positions", frmForm.Name, Join(saSizes, ","), sInitFileName)
    
End Sub
Public Sub RestoreFormSize(frmForm As Form, Optional ByVal sInitFileName As String = "")

Dim sData       As String
Dim saSizes()   As String
Dim uPosition   As FormPosition

    With uPosition
        
        'Retrieve the form's saved positions
        sData = GetInitEntry("Positions", frmForm.Name, "", sInitFileName)
        
        If Len(sData) = 0 Then
            .Left = frmForm.Left
            .Top = frmForm.Top
            .Width = frmForm.Width
            .Height = frmForm.Height
            .Maxed = (frmForm.WindowState = vbMaximized)
        Else
            saSizes() = Split(sData, ",")
            If UBound(saSizes) < 4 Then
                ReDim Preserve saSizes(4)
            End If
            .Left = Val(Trim$(saSizes(0)))
            .Top = Val(Trim$(saSizes(1)))
            .Width = Val(Trim$(saSizes(2)))
            .Height = Val(Trim$(saSizes(3)))
            .Maxed = (LCase$(Trim$(saSizes(4))) = "true")
        End If
    
        'Test positions against screen resolution
        If .Width < 150 Then
            .Width = frmForm.Width
        ElseIf .Width > Screen.Width Then
            .Width = Screen.Width
        End If
        If .Left < 0 Then
            .Left = frmForm.Left
        End If
        If .Left > Screen.Width - .Width Then
            .Left = Screen.Width - .Width
        End If
        If .Height < 150 Then
            .Height = frmForm.Height
        ElseIf .Height > Screen.Height Then
            .Height = Screen.Height
        End If
        If .Top < 0 Then
            .Top = frmForm.Top
        End If
        If .Top > Screen.Height - .Height Then
            .Top = Screen.Height - .Height
        End If
        
        'Position the form. Moving the form here will establish
        'its normal restored positions. It may be maximized in
        'the code that follows, but when restored by the user,
        'it will return to these positions.
        frmForm.Move .Left, .Top, .Width, .Height
    
        'Maximize the form if that's how it was previously shown.
        If .Maxed Then
            frmForm.WindowState = vbMaximized
        End If

    End With
    
End Sub

Public Function SetInitEntry(ByVal sSection As String, Optional ByVal sKeyName As String, Optional ByVal sValue As String, Optional ByVal sInitFileName As String = "") As Long

'This Function Writes a String To The Init File.
'Returns WritePrivateProfileString Success or Error.
'Creates and Uses sDefInitFileName (AppPath\AppEXEName.Ini)
'if sInitFileName Parameter Is Not Passed In.

'***** CAUTION *****
'If sValue is Null then sKeyName is deleted from the Init File.
'If sKeyName is Null then sSection is deleted from the Init File.

Dim sInitFile As String

    'If Init Filename NOT Passed In
    If Len(sInitFileName) = 0 Then
        'If Static Init FileName NOT Already Created
        If Len(sDefInitFileName) = 0 Then
            'Create Static Init FileName
            sDefInitFileName = App.Path
            If Right$(sDefInitFileName, 1) <> "\" Then
                sDefInitFileName = sDefInitFileName & "\"
            End If
            sDefInitFileName = sDefInitFileName & App.EXEName & ".ini"
        End If
        sInitFile = sDefInitFileName
    Else    'If Init Filename Passed In
        sInitFile = sInitFileName
    End If
    
    If Len(sKeyName) > 0 And Len(sValue) > 0 Then
        SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, ByVal sValue, sInitFile)
    ElseIf Len(sKeyName) > 0 Then
        SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, vbNullString, sInitFile)
    Else
        SetInitEntry = WritePrivateProfileString(sSection, vbNullString, vbNullString, sInitFile)
    End If

End Function



