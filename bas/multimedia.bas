Attribute VB_Name = "Multimedia"
'Declare the instruction

'DECLARATIONS FOR Multimedia.cls
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_SYNC = &H0         '  play synchronously (default)
Private Const SND_STOP = &H5         '  play synchronously (default)

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Function Short_Name(Long_Path As String) As String

    Dim Short_Path As String
    Dim PathLength As Long
    Short_Path = Space(250)
    PathLength = GetShortPathName(Long_Path, Short_Path, Len(Short_Path))


    If PathLength Then
        Short_Name = Left$(Short_Path, PathLength)
        
    End If
End Function

Public Sub PlayWave(FileName As String, Optional WaitUntilFinished As Boolean)
Dim ret As Long
If WaitUntilFinished = False Then ret = sndPlaySound(Short_Name(FileName), SND_ASYNC) Else ret = sndPlaySound(FileName, SND_SYNC)
End Sub

Public Sub LoopWave(FileName As String)
Dim ret As Long
ret = sndPlaySound(Short_Name(FileName), SND_LOOP)
End Sub

Public Sub LoopStop()
ret = sndPlaySound("", SND_STOP)
End Sub


Public Sub OpenCD()
    RetValue = MciSendString("set CDAudio door open", vbNullString, 0, 0)
End Sub
Public Sub CloseCD()

    RetValue = MciSendString("set CDAudio door closed", vbNullString, 0, 0)
End Sub



