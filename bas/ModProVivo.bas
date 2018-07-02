Attribute VB_Name = "Module1"
'««««««««««««««««««««««««««-ÑÉ-»»»»»»»»»»»»»»»»»»»»»»»»»»
'
'    This Code Is The Visual Basic 5 For ProVivo98
'            Programmed By ******** SoftWareZ
'               No Rights Currently Reserved
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'    All I have To Say is that if anyone uses this code
'    and makes a better Vivo Player than ProVivo Please
'    Send it to me.. I need a good Vivo Player.
'    Especially if it Plays Using DirectX or OpenGL
'    I Hope You Can work out most of the code.. its
'    fairly Simple.
'               ProVivo@hotmail.com
'
'                    Nothing More.
'
'««««««««««««««««««««««««««-ÑÉ-»»»»»»»»»»»»»»»»»»»»»»»»»»


Type Rect
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type


Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Global Const BUFFERSIZE = 255
Public MsBuffer As String * 255
Global ListTemp() As String
Global CurrentPlay As Integer
Global BackPit As String

Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lpReserved As Any) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

Sub Main()
    On Error GoTo FirstStart
    Dim Winpath, Temp As String
    Dim TempNum As Integer
    Winpath = WinDirectory
    Open Winpath & "\ProVivo98.ini" For Input As #9
        While Not EOF(9)
            Input #9, Temp
            Select Case UCase(Left(Temp, 8))
                Case "TITLESCR"
                    FrmMenu.MnuScroll_Click
                Case "CONTROLB"
                    FrmMenu.mnuControlBar_Click
                Case "FORMONTO"
                    FrmMenu.mnuOnTop_Click
                Case "REPEATVI"
                    FrmMenu.mnuRepeat_Click
                Case "BACKPICT"
                    BackPit = Right(Temp, Len(Temp) - 10)
                Case "WINDOWSI"
                    TempNum = Right(Temp, 1)
                    FrmMenu.mnuSize_Click (TempNum)
                Case "SHOWBACK"
                    FrmMenu.mnuBacker_Click
            End Select
        Wend
        FrmVivo.Show
        Exit Sub
ResumeStart:
    On Error GoTo ErrorNext
    MsgBox ("This Should Be The First Time You Have Used This Program. Preparing Setup Initilization.")
    FileCopy App.Path & "\vvppweb.ocx", Winpath & "\System\vvppweb.ocx"
    TempString = "C:\WINDOWS\SYSTEM\Regsvr32.exe /s " & Winpath & "\System\vvppweb.ocx"
    Shell TempString, vbHide
    FileCopy App.Path & "\comdlg32.ocx", Winpath & "\System\comdlg32.ocx"
    TempString = "C:\WINDOWS\SYSTEM\Regsvr32.exe /s " & Winpath & "\System\comdlg32.ocx"
    Shell TempString, vbHide
    FileCopy App.Path & "\msvbvm50.dll", Winpath & "\System\msvbvm50.dll"
    TempString = "C:\WINDOWS\SYSTEM\Regsvr32.exe /s " & Winpath & "\System\msvbvm50.dll"
    Shell TempString, vbHide
    Open Winpath & "\ProVivo98.ini" For Output As #8
    Close #8
    SourcePath = App.Path & "\ProVivo.exe"
    Shortcutname = "ProVivo98"
    DestinationPath = Winpath + "\Desktop"
    DestinationPath = Mid(DestinationPath, InStr(DestinationPath, ":") + 1)
    DestinationPath = "...." + DestinationPath
    T = fCreateShellLink(DestinationPath, Shortcutname, SourcePath, "")
    FrmVivo.Show
FirstStart:
    Resume ResumeStart
ErrorNext:
    Resume Next
End Sub


Function WinDirectory() As String
    Dim lBytes As Long
    lBytes = GetWindowsDirectory(MsBuffer, BUFFERSIZE)
    WinDirectory = Left$(MsBuffer, lBytes)
End Function
Sub OpenVivoFile(VivoFileName As String)
    ReDim ListTemp(0)
    ListTemp(0) = VivoFileName
    FrmVivo.OpenVivo (VivoFileName)
End Sub
Sub GetINetList(ListFile As String)
    Dim TempString As String
    Dim n As Integer
    On Error GoTo ErEnd
    Open ListFile For Input As #1
    ReDim ListTemp(0)
    n = 0
    While Not EOF(1)
        Input #1, TempString
        If TempString <> "" Then
            ReDim ListTemp(n)
            ListTemp(n) = TempString
        End If
    Wend
ErEnd:
    Close #1
End Sub
Sub Get_PlayList(ListFile As String)
Dim n, EquNum As Integer
Dim Temp, Temp2, TempDir, TempVar As String
    On Error GoTo ErrorEnd
    Temp = ListFile
    ReDim ListTemp(0)
    n = 0
    If Temp = "" Then Exit Sub
    TempDir = Dir(Temp)
    TempDir = Left(Temp, (Len(Temp) - Len(TempDir)))
    Open Temp For Input As #9
    Do While Not EOF(9)
        Input #9, Temp2
        If UCase(Temp2) = "[PLAYLIST]" Then
            Do While Not EOF(9)
                Input #9, Temp2
                EquNum = InStr(Temp2, "=")
                If EquNum <> 0 And Left(UCase(Temp2), 6) <> "NUMBER" Then
                    Temp2 = Right(Temp2, Len(Temp2) - EquNum)
                    ReDim Preserve ListTemp(n)
                    If Dir(Temp2) <> "" Then
                        ListTemp(n) = Temp2
                    Else
                        ListTemp(n) = TempDir + Temp2
                    End If
                    n = n + 1
                End If
            Loop
        Else
            If Temp2 <> "" Then
                If Dir(Temp2) <> "" Or UCase(Left(Temp2, 4)) = "HTTP" Then
                    ListTemp(0) = Temp2
                Else
                    ListTemp(0) = TempDir + Temp2
                End If
                n = n + 1
            End If
            Do While Not EOF(9)
                Input #9, Temp2
                If Temp2 <> "" Then
                    ReDim Preserve ListTemp(n)
                    If Dir(Temp2) <> "" Or UCase(Left(Temp2, 4)) = "HTTP" Then
                        ListTemp(n) = Temp2
                    Else
                        ListTemp(n) = TempDir + Temp2
                    End If
                    n = n + 1
                End If
            Loop
        End If
    Loop
    For n = 0 To UBound(ListTemp)
        If ListTemp(n) = "" Then
            For m = n To UBound(ListTemp) - 1
                ListTemp(m) = ListTemp(m + 1)
            Next
            ReDim Preserve ListTemp(UBound(ListTemp) - 1)
        End If
    Next
ErrorEnd:
Close #9
End Sub

