Attribute VB_Name = "Multimedia"
'Module MPEG,AVI,sequencer,audio source code
'This code written by abdullah al-ahdal e-mal:a_ahdal@yahoo.com
'for planet source code
'I written this code (standard code) to make the best and the easist
'dealing with multimedia file (All types) By pure API.
'In this Module ready function to use it in your projects
'Just add this code to your project and you will have the
'easist way to Controlling with multimedia files just you
'must know how can call these function from this module
'and how you can deal with it "if it success or not"
'All Functions in this Module will return a value
'if the Function success or not.

'Special Thanks to:
'Janet because he solved the problem of File name and the
'Path

'For any request Contact to me at : a_ahdal@yahoo.com

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Dim glo_hWnd As Long
Dim glo_from As Long
Dim glo_to As Long

Public Function OpenMPEG(hwnd As Long, filename As String, typeAviOrMpeg As String) As String
'Callig OpenMPEG will open the multimedia file
'the first parameter for the handle of the window
'which you want play in. you can put handle for
'your desktop if you wanna playing move in your desktop
'the second parameter is the file name and the path it can contain any space "Special thanks to Janet"
'the third parameter is a type of MCI device and it could be from the folowing:
'Type MCI       description                 driver file
'sequencer      dealing with midi and rmi   mciseq.drv
'               files
'MPEGVideo      dealing with most multimedia  mciqtz.drv
'               like mpg,mp3,mp2..
'               au,aiff,..etc
'               I got this info from my
'               experiment when I opened
'               System.ini in section MCI
'avivideo       deling with avi movie          mciavi.drv
'note : Type "MpegVideo" support these extensions:
'qt , mov, dat,snd, mpg, mpa, mpv, enc, m1v, mp2,mp3, mpe, mpeg, mpm
'au , snd, aif, aiff, aifc,wav ,,etc.
'but type "MpegVideo" not support midi files because of this you
'will use type "sequencer" for midi and rmi files

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

'Okay make sure if you used this function don't forget to use function CloseMPEG
'When you will end your program or you will error message

Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128

Dim tmp As String * 255
Dim lenShort As Long
Dim ShortPathAndFie As String
lenShort = GetShortPathName(filename, tmp, 255)
ShortPathAndFie = Left$(tmp, lenShort)

    
glo_hWnd = hwnd
cmdToDo = "open " & ShortPathAndFie & " type " & typeAviOrMpeg & " Alias mpeg parent " & hwnd & " Style 1073741824"
dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)


If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128
    OpenMPEG = ret: Exit Function
End If
    
OpenMPEG = "Success"
End Function

Public Function PlayMPEG(from_where As String, to_where As String) As String
'calling PlayMPEG will playing the multimedia file
'the first parameter for from where playing file
'the second parameter for to where playing file
'if the first parameter is vbNullString and the second parameter is vbNullString the Function Will:
'playing from the beginning to end.
'if the first parameter is 10 and the second parameter is 100 the Function Will:
'playing from 10 to 100 and stop.
'if the first parameter is vbNullString and the second parameter is 100 the Function Will:
'playing from the beginning to 100 and stop.
'if the first parameter is 104 and the second parameter is vbNullString the Function Will:
'playing from 104 to end.
'Note :the numbers 10,100,104 is an example for from playing to where end playing

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

If from_where = vbNullString And to_where = vbNullString Then
glo_from = 1
glo_to = GetTotalframes
ElseIf Not from_where = vbNullString And Not to_where = vbNullString Then
glo_from = from_where
glo_to = to_where
ElseIf Not from_where = vbNullString And to_where = vbNullString Then
glo_from = from_where
glo_to = GetTotalframes
ElseIf from_where = vbNullString And Not to_where = vbNullString Then
glo_from = 1
glo_to = to_where
End If

Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128
cmdToDo = "play mpeg from " & glo_from & " to " & glo_to

dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128
    PlayMPEG = ret
    Exit Function
End If

PlayMPEG = "Success"
End Function

Public Function CloseMPEG() As String
'calling CloseMPEG will close the multimedia file
'you must call this function if you call OpenMPEG
'And want to close your program or you will get an
'error message

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Close mpeg", 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128
    CloseMPEG = ret
    Exit Function
End If

CloseMPEG = "Success"
End Function

Public Function PauseMPEG() As String
'calling PauseMPEG will pause the multimedia file

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Pause mpeg", 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128
    PauseMPEG = ret
    Exit Function
End If
    
PauseMPEG = "Success"
End Function

Public Function StopMPEG() As String
'calling StopMPEG will Stop the multimedia file

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Stop mpeg", 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128
    StopMPEG = ret
    Exit Function
End If

StopMPEG = "Success"
End Function

Public Function ResumeMPEG() As String
'calling ResumeMPEG will Resume the multimedia file

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Resume mpeg", 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128
    ResumeMPEG = ret
    Exit Function
End If

ResumeMPEG = "Success"
End Function

Public Function GetStatusMPEG() As String
'Calling Function GetStatusMPEG

'Note : if this Function success will return value string
'(the status of multimedia file) if it "playing" or "paused" or "stopped"
'or if not will return value string "ERROR"


'also you can exame the status like this: you can copy it
'Dim Result As String
'Result = GetStatusMPEG
'If Result = "ERROR" Then 'this mean failed then write your commands here
''.....
''....
''..
'ElseIf Result = "playing" Then 'this mean it now playing .ok write your commands here
''....
''...
''..
'ElseIf Result = "stopped" Then 'this mean it now stopped .ok write your commands here
''....
''...
''..
'ElseIf Result = "paused" Then 'this mean it now paused .ok write your commands here
''....
''...
''..

'End If


Dim dwReturn As Long
Dim status As String * 255
Dim ret As String * 255

dwReturn = mciSendString("status mpeg mode", status, 255, 0&)

If Not dwReturn = 0 Then  'not success
    GetStatusMPEG = "ERROR"
    Exit Function
End If

' haha what you think why I wirte these below lines
Dim I As Integer
Dim CharA As String
Dim RChar As String
RChar = Right$(status, 1)
For I = 1 To Len(status)
    CharA = Mid(status, I, 1)
    If CharA = RChar Then Exit For
    GetStatusMPEG = GetStatusMPEG + CharA
Next I
' the answer : to get just the string to we wanna make compare in future like this:
' if GetStatusMPEG = "playing" then msgbox "Now plying"
End Function

Public Function GetTotalframes() As Long
'calling GetTotalframes will Get the Total frames for
'the multimedia file

'Note : if this Function success will return value long
'is "number of total frames"
'or if not will return value long is -1

Dim dwReturn As Long
Dim Total As String * 255

dwReturn = mciSendString("set mpeg time format frames", Total, 255, 0&)
dwReturn = mciSendString("status mpeg length", Total, 255, 0&)

If Not dwReturn = 0 Then  'not success
    GetTotalframes = -1
    Exit Function
End If

GetTotalframes = Val(Total)
End Function

Public Function GetTotalTimeByMS() As Long
'calling GetTotalTimeByMS will Get the Total time by
'millisecond for the multimedia file

'Note : if this Function success will return value long
'is "the Total time by millisecond"
'or if not will return value long is -1

Dim dwReturn As Long
Dim TotalTime As String * 255


dwReturn = mciSendString("set mpeg time format ms", Total, 255, 0&)
dwReturn = mciSendString("status mpeg length", TotalTime, 255, 0&)

mciSendString "set mpeg time format frames", Total, 255, 0& ' return focus to frames not to time

If Not dwReturn = 0 Then  'not success
    GetTotalTimeByMS = -1
    Exit Function
End If

GetTotalTimeByMS = Val(TotalTime)
End Function

Public Function MoveMPEG(to_where As Long) As String
'calling MoveMPEG will seek (change the position)for
'the multimedia file
'it need just one parameter for to where want to change
'position (must be number of frame you want to go to)

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 255

dwReturn = mciSendString("seek mpeg to " & to_where, 0&, 0&, 0&)
mciSendString "Play mpeg", 0&, 0&, 0&

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128
    MoveMPEG = ret
    Exit Function
End If
MoveMPEG = "Success"
End Function

Public Function GetCurrentMPEGPos() As Long
'Calling Function GetCurrentMPEGPos
'the returned value from this function is number of current frame
'and if the function failed will return value -1


Dim dwReturn As Long
Dim pos As String * 255

dwReturn = mciSendString("status mpeg position", pos, 255, 0&)

If Not dwReturn = 0 Then  'not success
    GetCurrentMPEGPos = -1
    Exit Function
End If

GetCurrentMPEGPos = Val(pos)
End Function

Public Function PutMPEG(Left As Long, Top As Long, Width As Long, Height As Long) As String
'Calling PutMPEG will resize the move and :
'if you are set parameter width or Height zero
'the function will get the actual size of the window which
'want to play in and resize the movie to fit the window

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

Dim dwReturn As Long
Dim ret As String * 255
If Width = 0 Or Height = 0 Then
    Dim rec As RECT
    Call GetWindowRect(glo_hWnd, rec)
    Width = rec.Right - rec.Left
    Height = rec.Bottom - rec.Top
End If

dwReturn = mciSendString("put mpeg window at " & Left & " " & Top & " " & Width & " " & Height, 0&, 0&, 0&)

If Not dwReturn = 0 Then  'not success
    mciGetErrorString dwReturn, ret, 128
    PutMPEG = ret
    Exit Function
End If

PutMPEG = "Success"
End Function
Public Function GetPercent() As Long
'Calling Function GetPercent
'the returned value from this function is Percent "Progress"
'if it successed and if the function failed will return value -1

On Error Resume Next
Dim TotalFrames As Long
Dim CurrFrame As Long
TotalFrames = GetTotalframes
CurrFrame = GetCurrentMPEGPos

If TotalFrames = -1 Or CurrFrame = -1 Then
GetPercent = -1
Exit Function
End If

GetPercent = CurrFrame * 100 / TotalFrames
End Function
Public Function GetFramesPerSecond() As Long
'this Function Will return amount frames per second if it
'Success or if not will return value -1
Dim TotalFrames As Long
Dim TotalTime As Long
TotalTime = GetTotalTimeByMS
TotalFrames = GetTotalframes
If TotalFrames = -1 Or TotalTime = -1 Then
    GetFramesPerSecond = -1
    Exit Function
End If
GetFramesPerSecond = TotalFrames / (TotalTime / 1000)
End Function
Public Function AreMPEGAtEnd() As Boolean

'Note:I used API callback to make callback and opreator
'AddressOf but the program will be very slow because of this
'I removed that way and I used this way


'This Function will tell if multimedia file now at end
'to use this Function put it in a timer and set Interval
'for a timer = 100 and make the timer false and after Play
'Multimedia files Successfully set the timer true.
'The Commands Which you will put it in a timer the Following:


'Copy the Following in a timer
'If AreMPEGAtEnd = True Then
''this mean  file multimedia at the end now then
''write your commnad here or call you favourit Fucntion
''or even you can play the file again or paly the next file
''if you had a list of multimedia files.
'.....
'...
'..
'if you wanna know if the multimedia file
'at the end now don't use option Auto Repeat
'you must do auto repeat by yourself by the following command
'in this place after make the previous compare (I mean afyer compare in a timer)

'Result = PlayMPEG(txtFrom, TxtTo)
'or you have choice to close this File and open
'another file and play it( this if had a list of files)
'like this command after make the previous compare(I mean afyer compare in a timer)
'Dim Result As String
'Result = CloseMPEG
'Result = OpenMPEG(FrameVideo.hwnd, filename, typeDevice) 'call now function openMPEG
'Result = PlayMPEG(txtFrom, TxtTo)

'TimerName.Enabled = False ' and remeber don't forget
'write this line becuase you got what you want then
'Close the timer.Okay.
'Else
'this mean result calling function false and this mean the
'multimedia file not at the end now
'....
'...
'..

'End If

Dim currpos As Long
currpos = Val(GetCurrentMPEGPos)
If glo_to = currpos Or (glo_to - 1) < currpos Then
AreMPEGAtEnd = True
Else
AreMPEGAtEnd = False
End If
End Function
Public Sub SetAutoRepeat(autoTrueOrFalse As Boolean)
'This cool sub if you want to make the multimedia file
'auto repeat by it self or remove the auto repeat
'if the parameter is true will make the function
'Auto repeat or it else will remove the auto repeat

If autoTrueOrFalse = True Then
    Call SetTimer(glo_hWnd, 500, 100, AddressOf TimerFunction)
Else
    Call KillTimer(glo_hWnd, 500)
End If
End Sub

Sub TimerFunction()
'Important for auto repeat
Dim currpos As Long
currpos = Val(GetCurrentMPEGPos)
If glo_to = currpos Or (glo_to - 1) < currpos Then PlayMPEG Str(glo_from), Str(glo_to)
End Sub

Public Sub SetDefaultDevice(typeDevice As String, drvDefaultDevice As String)
'this sub is very important to set the default MCI device
'maybe xing mpeg installed in your computer and it not support
'all multimudia files
'because of this you can rest the default device of MCI to
'drivers microsft
'which came with windows or when install Microsft media player
'ok any way the default device Following:
'Device Type        Driver
'MPEGVideo          mciqtz.drv          this is the most important
'sequencer          mciseq.drv
'avivideo           mciavi.drv
'waveaudio          mciwave.drv
'videodisc          mcipionr.drv
'cdaudio            mcicda.drv

'e.g. :
'SetDefaultDevice "MPEGVideo", "mciqtz.drv"' this the most
'improtant device and it will receives calls mci
'Some programs change this device like xing mpeg
'and if this occur you can not play all mutimedia files
'and you will unexpected errors
'because of this write this line when your program loaded
'SetDefaultDevice "MPEGVideo", "mciqtz.drv"
'to set the strongest default device

Dim Res As String
Dim tmp As String * 255
Res = GetWindowsDirectory(tmp, 255)
Windir = Left$(tmp, Res)
Res = WritePrivateProfileString("MCI", typeDevice, drvDefaultDevice, Windir & "\" & "system.ini")
End Sub

Public Function GetDefaultDevice(typeDevice As String) As String
'this Function help you if you want to the default device
'the parameter must be the device type like:
'MPEGVideo
'sequencer
'avivideo
'waveaudio
'videodisc
'cdaudio
'and the returned value is a string for the default device
'Please the description of sub SetDefaultDevice

Dim tmp As String * 255
Dim Res As String
Res = GetWindowsDirectory(tmp, 255)
Windir = Left$(tmp, Res)
Res = GetPrivateProfileString("MCI", typeDevice, "None", tmp, 255, Windir & "\" & "system.ini")
GetDefaultDevice = Left$(tmp, Res)
End Function

'Okay I hope you Enjoyed
'You can use this module in your own projects if you wanna
'the easist deal with multimedia.
'Using API is more stronger than using controls and not take a space
'for any request, suggestions,Devlopment or bugs   e-mail at
'a_ahdl@yahoo.com
'Thank you

