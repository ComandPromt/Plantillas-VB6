Attribute VB_Name = "WinampModule"
'*******************************************************************
'********************** Winamp Control Module **********************
'***********************         by          ***********************
'*************************  James Crasta   *************************
'*******************************************************************
'*******************************************************************
' This module was created by me when i couldn't find any winamp
' module for VB that would do what I wanted so i made one using the
' winamp button code numbers found in the nsdn section of Winamp's site.
' The code is pretty self-explanatory, i have included comments to help you understand this

Option Explicit 'all variables must be declared!

'Constants used throughout the module
Public Const WM_COMMAND = 273
Public Const WM_USER = 1024
Public Const WM_WA_IPC = &H400

' Public Variable
Public WinampID As Long
Public WinampPath As String
Public LastWinampCaption As String
Public LastTitle

'**************************
'****** API DECLARES ******
'**************************
' Find Winamp Window
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
' SendMessage To Window (waits for reply)
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal WndID As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
' PostMessage To Window (returns true/false)
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
'Public Const WM_KEYDOWN = &H100
'Public Const WM_KEYUP = &H101
'Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'Public Const WM_LBUTTONDOWN = &H201
'Public Const WM_LBUTTONUP = &H202
'Public Const WM_SETTEXT = &HC
'Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long)
'Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long)
'Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long



'************postmessage commands
'************activate using WM_DO(constant_name)
'************i.e:  WM_DO(PREV_TRACK)
Public Const PREV_TRACK As Long = 40044
Public Const NEXT_TRACK As Long = 40048
Public Const PLAY_TRACK As Long = 40045
Public Const PAUSE_TRACK As Long = 40046
Public Const STOP_TRACK As Long = 40047
Public Const FADEOUT_STOP As Long = 40147
Public Const STOP_AFTER_CURRENT As Long = 40157
Public Const FORWARD_5_SEC As Long = 40148
Public Const BACK_5_SEC As Long = 40144
Public Const GO_BEGINNING_PLAYLIST As Long = 40154
Public Const GO_END_PLAYLIST As Long = 40158
Public Const OPEN_FILE_DIALOG As Long = 40029
Public Const OPEN_URL_DIALOG As Long = 40155
Public Const OPEN_FILE_INFO_DIALOG As Long = 40188
Public Const DISP_ELAPSED_TIME As Long = 40037
Public Const DISP_REMAINING_TIME As Long = 40038
Public Const PREFS_DIALOG As Long = 40012
Public Const VIS_OPTIONS As Long = 40190
Public Const VIS_PLUGIN_OPTIONS As Long = 40191
Public Const START_VIS = 40192
Public Const SHOW_ABOUT As Long = 40041
Public Const TOGGLE_TITLE_AUTOSCROLL As Long = 40189
Public Const TOGGLE_ALWAYS_ON_TOP As Long = 40019
Public Const TOGGLE_WINDOWSHADE As Long = 40064
Public Const TOGGLE_PLAYLIST_WINDOWSHADE As Long = 40266
Public Const TOGGLE_DOUBLESIZE As Long = 40165
Public Const TOGGLE_EQ As Long = 40036
Public Const TOGGLE_PLAYLIST As Long = 40040
Public Const TOGGLE_WINAMP_VISIBLE As Long = 40258
Public Const TOGGLE_MINIBROWSER As Long = 40298
Public Const TOGGLE_EASYMOVE As Long = 40186
Public Const VOLUME_RAISE As Long = 40058
Public Const VOLUME_LOWER As Long = 40059
Public Const TOGGLE_REPEAT As Long = 40022
Public Const TOGGLE_SHUFFLE As Long = 40023
Public Const JUMPTO_TIME_DIALOG As Long = 40193
Public Const JUMPTO_FILE_DIALOG As Long = 40194
Public Const OPEN_SKIN_SLECTOR As Long = 40219
Public Const CONFIG_CURRENT_VIS As Long = 40221
Public Const RELOAD_CUR_SKIN As Long = 40291
Public Const WINAMP_EXIT As Long = 40001

'sendmessage commands
'use variable = WM_GET(constant_name)
'example:  length = WM_GET(CLEAR_PLAYLIST)
Public Const WINAMP_VERSION As Long = 0         'Retrieves the version of Winamp running. Version will be 0x20yx for 2.yx. This is a good way to determine if you did in fact find the right window, etc.
Public Const CLEAR_PLAYLIST As Long = 1         'Clears Winamp 's internal playlist.
Public Const PLAYBACK_STATUS As Long = 2        'Returns the status of playback. If 'ret' is 1, Winamp is playing. If 'ret' is 3, Winamp is paused. Otherwise, playback is stopped.
Public Const SONG_LENGTH As Long = 3            'Returns the length of the song in seconds
Public Const SONG_POSITION As Long = 4          'Returns the current position in the song, in milliseconds
Public Const SEEK_CURRENT_TRACK As Long = 5     'Seeks within the current track. The offset is specified in 'data', in milliseconds.
Public Const WRITE_CURR_PLAYLIST As Long = 6    'Writes out the current playlist to Winampdir\winamp.m3u, and returns the current position in the playlist.
Public Const SET_PLAYLIST_POS As Long = 7       'Sets the playlist position to the position specified in tracks in 'data'.
Public Const PLAYLIST_LENGTH As Long = 8        'Returns length of the current playlist, in tracks.
Public Const PLAYLIST_POSITION As Long = 9      'Returns the position in the current playlist, in tracks (requires Winamp 2.05+).
Public Const SONG_SAMPLERATE  As Long = 10
Public Const SONG_BITRATE As Long = 11
Public Const SONG_NUMCHANNELS As Long = 12
Public Const SONG_TITLE As Long = 13

'###############################################
'########### the FindWinamp Function ###########
'###############################################
' This function will find the Window handle  of winamp and store it
' to the WinampID variable.  You must run this function first for any
' of the other functions in this module to work.  If no winamp is found,
' and a winamp path is specified below, then the FindWinamp function will
' run Winamp and wait for it to load before getting its window handle


Public Function FindWinamp()
On Error GoTo err
WinampPath = "D:\multimedia\winamp\winamp.exe"

WinampID = FindWindow("Winamp v1.x", 0)
If WinampID = 0 And WinampPath <> "" Then
    Dim deluseless As Variant
    deluseless = Shell(WinampPath)
    Dim cntr As Integer
    Do
        cntr = cntr + 1
        WinampID = FindWindow("Winamp v1.x", 0)
        DoEvents: DoEvents: DoEvents: DoEvents
        If cntr = 60000 Then
            MsgBox "Winamp failed to open in the time given.  Please contact the program creator for more information"
            End
        End If
    Loop Until WinampID <> 0
End If
    Exit Function
err:
    Call WinampModErrorHandler
End Function

' These are the standalone single-function functions.
' They are pretty self-explanatory

Public Function WM_SetVolume(Volume As Integer) As Long
    On Error GoTo err
    'Sets the volume (Volume must be between 0 - 255)
    WM_SetVolume = SendMessage(WinampID, WM_WA_IPC, Volume, 122)
    Exit Function
err:
    Call WinampModErrorHandler
End Function

Public Function WM_SetPlaylistPos(PosInPlaylist As Integer) As Long
    On Error GoTo err
    'Sets the playlist position to the position specified in tracks
    WM_SetPlaylistPos = SendMessage(WinampID, WM_WA_IPC, PosInPlaylist, 121)
    Exit Function
err:
    Call WinampModErrorHandler
End Function

Public Function WM_SetPanning(Panning As Integer)
    On Error GoTo err
    'Sets the panning to 'data', which can be between 0 (all left) and 255 (all right).
    WM_SetPanning = SendMessage(WinampID, WM_WA_IPC, Panning, 123)
    Exit Function
err:
    Call WinampModErrorHandler
End Function

Public Sub WinampModErrorHandler(Optional desc As String, Optional number As Variant)
    Debug.Print "There has been an error: "; desc; " and number "; number
End Sub



'**************************************************************************
'****************YOU ARE NOW ENTERING THE COMPLEX CODE AREA****************
'**************************************************************************
' Below is the code for the WM_DO and WM_GET commands.
' Be careful.  If you do not know what you are doing,
' you may end up messing something up



Public Function WM_DO(cmnd As Long)
    On Error GoTo err
    Dim tmp As String
    Dim ret As Integer
    Dim isplay As Integer
    'isplay = SendMessage(WinampID, WM_USER, 0, 104)
    ret = PostMessage(WinampID, WM_COMMAND, cmnd, 0)
    Exit Function
err:
    Call WinampModErrorHandler
End Function

Public Function WM_GET(cmnd As Long, Optional data As Long) As Variant
    On Error GoTo err:
    Dim tmp As String
    Dim ret As Integer
    Dim isplay As Integer
    isplay = SendMessage(WinampID, WM_USER, 0, 104)
    Select Case cmnd
        Case WINAMP_VERSION ' returns current winamp version
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 0) ' returns winamp version in LONG format
        Case CLEAR_PLAYLIST ' clears all items in the playlist
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 101) ' does not return anything
        Case PLAYBACK_STATUS 'gets current playback status
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 104) ' returns 1 for playing, 3 for paused, 0 for stopped
        Case SONG_LENGTH
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 1, 105) ' returns track length in seconds
        Case SONG_POSITION
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 105) ' returns position in the current track in milliseconds
        Case PLAYLIST_LENGTH
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 124) ' returns number of songs in playlist
        Case PLAYLIST_POSITION
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 125) ' returns the currently playing track number
        Case SONG_SAMPLERATE
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 0, 126) ' returns the currently playing song's sample rate
        Case SONG_BITRATE
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 1, 126) ' returns the currently playing song's bit rate
        Case SONG_NUMCHANNELS
            WM_GET = SendMessage(WinampID, WM_WA_IPC, 3, 126) ' returns the currently playing song's sample rate
        Case WRITE_CURR_PLAYLIST  'Writes out the current playlist to Winampdir\winamp.m3u, and returns the current position in the playlist.
                WM_GET = SendMessage(WinampID, WM_WA_IPC, 3, 120) ' writes the winamp.m3u file
        Case SONG_TITLE ' returns a string with the song title in it
            ' Since I couldnt find any API that
            ' would query winamp, i decided to
            ' read in the caption and trim it down
            ' to just the title.
            Dim strBuffer As String, lngtextlen As Long
            Let lngtextlen& = GetWindowTextLength(WinampID) 'gets the length of the caption
            Let strBuffer$ = String$(lngtextlen&, 0&) 'i dont know why this is necessary, i found it in someone else's API code
            Call GetWindowText(WinampID, strBuffer$, lngtextlen& + 1&) ' reads in the caption text
            If strBuffer$ = LastWinampCaption Then
                WM_GET = LastTitle
            Else
                LastWinampCaption = strBuffer
                If LCase(strBuffer$) Like "*[paused]" = False Then ' queries if the [Paused] string is there and removes it
                    strBuffer$ = Left(strBuffer$, Len(strBuffer) - 9)
                End If
                strBuffer$ = Mid(strBuffer$, 1, Len(strBuffer) - 8) ' removes the -Winamp
                Dim findDot As Integer
                findDot = InStr(1, strBuffer, ".") ' finds the dot in the number at the beginning
                LastTitle = Trim(Mid(strBuffer$, findDot + 1)) 'Returns the final title value
                WM_GET = LastTitle
            End If
    End Select
    Exit Function
err:
    Call WinampModErrorHandler(err.Description, err.number)
End Function

