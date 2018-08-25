Attribute VB_Name = "drives"
Option Explicit

'CD-ROM
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public ZTrack As Integer, ZAutoRepeatTrack As Integer, ZAutoRepeatCD As Boolean, ZProgramNumber As Integer, ZRandomTracks As Integer
Public Record_M As Integer, Record_S As Integer, ZFinished As Boolean

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public FoundWindows As String

Public fastForwardSpeed As Long    'seconds to seek for ff/rew
Public fPlaying As Boolean         'true if CD is currently playing
Public fCDLoaded As Boolean        'true if CD is the the player
Public numTracks As Integer        'number of tracks on audio CD
Public trackLength() As String     'array containing length of each track
Public track As Integer            'current track
Public min As Integer              'current minute on track
Public sec As Integer              'current second on track
Public cmd As String               'string to hold mci command strings

Public lCurTrack As Long, lTrackLengths() As Long, lStart As Long, lFinish As Long, aFile As String, bGroups As Boolean

Public Const DRIVE_ANY = 0
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Const GETDI_SERIAL = 1
Public Const GETDI_LABEL = 2
Public Const GETDI_TYPE = 3

'My Defs
Public OpInProgress As Boolean, MusicT As Integer, ZAutoRun As Boolean

'General Api Declarations
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'Keep Form on Top,
'Delete if you dont need it
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40
Public Function GetBitmapRegion(cPicture As StdPicture, cTransparent As Long)

Dim hRgn As Long, tRgn As Long
Dim X As Integer, Y As Integer, X0 As Integer
Dim hDC As Long, BM As BITMAP

hDC = CreateCompatibleDC(0)
If hDC Then
    SelectObject hDC, cPicture

    GetObject cPicture, Len(BM), BM
    hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
    For Y = 0 To BM.bmHeight
        For X = 0 To BM.bmWidth
            While X <= BM.bmWidth And GetPixel(hDC, X, Y) <> cTransparent
                X = X + 1
            Wend
            X0 = X
            While X <= BM.bmWidth And GetPixel(hDC, X, Y) = cTransparent
                X = X + 1
            Wend
            If X0 < X Then
                tRgn = CreateRectRgn(X0, Y, X, Y + 1)
                CombineRgn hRgn, hRgn, tRgn, 4
                DeleteObject tRgn
            End If
        Next X
    Next Y
    GetBitmapRegion = hRgn
    DeleteObject SelectObject(hDC, cPicture)
End If

DeleteDC hDC

End Function
Public Function GetDriveInfo(strDrive As String, iType As Integer) As String
On Local Error Resume Next
Err.Clear

Dim SerialNum As Long, strLabel As String, strType As String, lRetVal As Long

strLabel = Space(256)
strType = Space(256)
lRetVal = GetVolumeInformation(strDrive, strLabel, Len(strLabel), SerialNum, 0, 0, strType, Len(strType))
Select Case iType
    Case Is = 1
    GetDriveInfo = CStr(SerialNum)
    Case Is = 2
    GetDriveInfo = strLabel
    Case Is = 3
    GetDriveInfo = strType
End Select

End Function
Public Function GetLocalDrives(lngType As Long) As Variant
On Local Error Resume Next
Err.Clear

Dim cResult As Long, i As Integer, intCount As Integer
Dim strTmpArray() As String
ReDim strTmpArray(0 To 25)

If lngType = DRIVE_ANY Then      'Loop and check for any drive
    For i = 0 To 25
        cResult = GetDriveType(Chr(65 + i) & ":\")
        If cResult <> 1 Then
            strTmpArray(intCount) = Chr(65 + i)
            intCount = intCount + 1
        End If
    Next i
Else
    'Loop and check for a specific type of drive
    For i = 0 To 25
        cResult = GetDriveType(Chr(65 + i) & ":\")
        If cResult = lngType Then
            strTmpArray(intCount) = Chr(65 + i)
            intCount = intCount + 1
        End If
    Next i
End If
'Only redim if one or more drives were found
If intCount > 0 Then
   ReDim Preserve strTmpArray(0 To intCount - 1)
   GetLocalDrives = strTmpArray
End If

End Function

Public Sub IniRead()
On Local Error Resume Next
Err.Clear

FindPath

Dim A As String

Open FoundWindows & "RJSoftCDPlayer.INI" For Input As #5
If Err = 0 Then
    Line Input #5, A
    frmMain.CheckSBLiveSoundCard.Value = Val(A)
 Else
    Close #5
    Err.Clear
    frmMain.CheckSBLiveSoundCard.Value = False
    IniWrite
End If
Close #5

If frmMain.CheckSBLiveSoundCard.Value <> False Then
    frmMain.ProgressBar1.Max = 128
    frmMain.ProgressBar1.min = 0
    frmMain.ProgressBar2.Max = 128
    frmMain.ProgressBar2.min = 0
    frmMain.Timer3.Interval = 1
 Else
    frmMain.Timer3.Interval = 0
End If

Err.Clear

End Sub

Public Sub FindPath()
On Local Error Resume Next
Err.Clear

Dim temp9$, X

temp9$ = String$(145, 0)
X = GetWindowsDirectory(temp9$, 145)
temp9$ = Left$(temp9$, X)

If Right$(temp9$, 1) <> "\" Then
    FoundWindows = temp9$ & "\"
Else
    FoundWindows = temp9$
End If
        
End Sub
Public Sub IniWrite()
On Local Error Resume Next
Err.Clear

FindPath

Open FoundWindows & "RJSoftCDPlayer.INI" For Output As #5
If Err = 0 Then
    Print #5, CStr(frmMain.CheckSBLiveSoundCard.Value)
End If
Close #5
Err.Clear

End Sub
