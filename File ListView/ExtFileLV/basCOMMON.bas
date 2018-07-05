Attribute VB_Name = "basCOMMON"
Option Explicit

Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'------------------------------------------------------------
' Author:  Clint LaFever [clint.m.lafever@cpmx.saic.com]
' Purpose:  Freezes any window by it's hWnd.  Pass 0 to unlock window.
' Parameters:  hWND=Window to Lock.  0 to Unlock
' Example:  Only one window may be locked at a time.
'           Returns non zero on success
' Date: January,13 1999 @ 12:20:55
'------------------------------------------------------------
Public Enum IMG_SIZE
    IMG_SIXTEEN = 16
    IMG_THIRTYTWO = 32
    IMG_ALREADYSET = 0
    IMG_CUSTOM = 1
End Enum
Public Enum AppIcons
    icon_FOLDER_CLOSED = 101
End Enum
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@usa.net]
' Purpose:  Used to Add an image to a ImageList from the resource file.  Note.  AppIcons must be declared.
' Parameters:
' Example:
' Date: July,21 1998 @ 18:22:18
'------------------------------------------------------------
Public Sub AddImage(imgLIST As ImageList, resICONVAL As AppIcons, Optional imgSIZE As IMG_SIZE = IMG_ALREADYSET, Optional CustomHeight As Long = 16, Optional CustomWidth As Long = 16)
    On Error Resume Next
    With imgLIST
        If imgSIZE <> IMG_ALREADYSET Then
            If imgSIZE <> IMG_CUSTOM Then
                .ImageHeight = imgSIZE
                .ImageWidth = imgSIZE
            Else
                .ImageHeight = CustomHeight
                .ImageWidth = CustomWidth
            End If
        End If
        .ListImages.Add , , LoadResPicture(resICONVAL, vbResIcon)
    End With
End Sub
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@usa.net]
' Purpose:  Changes the size of icons within an ImageList at RunTime.
' Parameters:
' Example:
' Date: July,21 1998 @ 18:22:47
'------------------------------------------------------------
Public Sub ChangeImageSize(imgLIST As ImageList, imgSIZE As IMG_SIZE, Optional CustomHeight As Long = 16, Optional CustomWidth As Long = 16)
    On Error Resume Next
    With imgLIST
        If imgSIZE <> IMG_ALREADYSET Then
            If imgSIZE <> IMG_CUSTOM Then
                .ImageHeight = imgSIZE
                .ImageWidth = imgSIZE
            Else
                .ImageHeight = CustomHeight
                .ImageWidth = CustomHeight
            End If
        End If
    End With
End Sub







Public Function FreezeWindow(Optional mLNGhWnd As Long = 0) As Long
    On Error Resume Next
    Dim x As Long
    FreezeWindow = LockWindowUpdate(mLNGhWnd)
End Function

