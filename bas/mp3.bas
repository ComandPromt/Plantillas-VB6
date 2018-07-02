Attribute VB_Name = "mp3"
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


Public Sub MP3Open(FileName As String, Optional MultiPlayID As String = "GalaxyAudio1")
Dim ret As Long
ret = MciSendString("OPEN " & Short_Name(FileName) & " Alias " & MultiPlayID, 0, 0, 0)

End Sub

Public Sub MP3Play(Optional MultiPlayID As String = "GalaxyAudio1")
Dim ret As Long
ret = MciSendString("Play " & MultiPlayID, 0, 0, 0)

End Sub

Public Sub MP3Stop(Optional MultiPlayID As String = "GalaxyAudio1")
Dim ret As Long

ret = MciSendString("Stop " & MultiPlayID, 0, 0, 0)

End Sub
Public Sub MP3Close(Optional MultiPlayID As String = "GalaxyAudio1")
Dim ret As Long
ret = MciSendString("Close " & MultiPlayID, 0, 0, 0)

End Sub

Public Sub MP3Seek(ByVal nPosition As Single, Optional MultiPlayID As String = "GalaxyAudio1")
    Dim ret As Long
 
    ret = MciSendString("Seek " & MultiPlayID & " to " & nPosition, "", 0, 0)
    
End Sub
Public Property Get MP3Length(Optional MultiPlayID As String = "GalaxyAudio1") As Single

    Dim nReturn As Long, nLength As Integer
    
    Dim sLength As String * 255
    
    nReturn = MciSendString("Status " & MultiPlayID & " length", sLength, 255, 0)
    nLength = InStr(sLength, Chr$(0))
    MP3Length = Val(Left$(sLength, nLength - 1))
    
End Property


Public Property Get MP3Position(Optional MultiPlayID As String = "GalaxyAudio1") As Single

    Dim nReturn As Integer, nLength As Integer
    
    Dim sPosition As String * 255
    
       
    nReturn = MciSendString("Status " & MultiPlayID & " position", sPosition, 255, 0)
    nLength = InStr(sPosition, Chr$(0))
    MP3Position = Val(Left$(sPosition, nLength - 1))
    
End Property


Property Get MP3Status(Optional MultiPlayID As String = "GalaxyAudio1") As String

    Dim nReturn As Integer, nLength As Integer
    
    Dim sStatus As String * 255

    
    nReturn = MciSendString("Status " & MultiPlayID & " mode", sStatus, 255, 0)
    
    nLength = InStr(sStatus, Chr$(0))
    MP3Status = Left$(sStatus, nLength - 1)
    
End Property
Public Sub MP3Pause(Optional MultiPlayID As String = "GalaxyAudio1")
Dim ret As Long
ret = MciSendString("Pause " & MultiPlayID, 0, 0, 0)
End Sub

Public Sub MP3Resume(Optional MultiPlayID As String = "GalaxyAudio1")
Dim ret As Long
ret = MciSendString("Resume " & MultiPlayID, 0, 0, 0)
End Sub

Private Function Short_Name(Long_Path As String) As String

    Dim Short_Path As String
    Dim PathLength As Long
    Short_Path = Space(250)
    PathLength = GetShortPathName(Long_Path, Short_Path, Len(Short_Path))


    If PathLength Then
        Short_Name = Left$(Short_Path, PathLength)
        
    End If
End Function

