Attribute VB_Name = "CD_Serial_Number"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" _
(ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal Filename As String) As Long

Global CDMin As Integer
Global CDSec As Integer
Global TMin As Integer
Global TSec As Integer
Global RMin As Integer
Global RSec As Integer
Global TimeTrack As String
Global TimeElapsed As String
Global TimeRemaining As String

Global Artist1 As String
Global Title1 As String

Global Artist2 As String
Global Title2 As String


Sub CDAudioProperties()
Dim T As Double
On Error Resume Next
T = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,3", 5)

End Sub


Function GetRawRemainingTime(MMCOntrol1 As Object) As String
Dim Z As String, Min As String, Sec As String, _
Temp As String

Z = GetRunningTime(MMCOntrol1)
Z = GetTrackTime(MMCOntrol1)

Min = LTrim$(Str$(CDMin - TMin))
Sec = LTrim$(Str$(CDSec - TSec))

RMin = Val(Min)
RSec = Val(Sec)

If RSec < 0 Then
    RSec = 60 + Val(Sec)
    RMin = RMin - 1
End If
Min = LTrim$(Str$(RMin))
Temp = Trim$(Str$(RSec))
If Len(Temp) = 1 Then
    Sec = "0" + Temp
Else
    Sec = Temp
End If
GetRawRemainingTime = Min + Sec
End Function
Function GetRemainingTime(MMCOntrol1 As Object) As String
Dim Z As String, Min As String, Sec As String, _
Temp As String

Z = GetRunningTime(MMCOntrol1)
Z = GetTrackTime(MMCOntrol1)

Min = LTrim$(Str$(CDMin - TMin))
Sec = LTrim$(Str$(CDSec - TSec))

RMin = Val(Min)
RSec = Val(Sec)

If RSec < 0 Then
    RSec = 60 + Val(Sec)
    RMin = RMin - 1
End If
Min = LTrim$(Str$(RMin))
Temp = Trim$(Str$(RSec))
If Len(Temp) = 1 Then
    Sec = "0" + Temp
Else
    Sec = Temp
End If
GetRemainingTime = Min + ":" + Sec
End Function


Sub GetTime(MMCOntrol1 As Object)
Dim Z$
Z$ = GetRemainingTime(MMCOntrol1)

End Sub

Function GetTrackTime(MMCOntrol1 As Object) As String
Dim Length&, Entry2$, Min$, Sec$, D$, Entry$
MMCOntrol1.TimeFormat = 2
Length& = MMCOntrol1.TrackLength
Min$ = Str$(Length& And &HFF)
Sec$ = LTrim$(Str$((Length& And 65280) / 256))
Entry2$ = Min$ & ":" & Sec$
If Len(Sec$) = 1 Then Entry2$ = Min$ + ":0" + Sec$
Entry$ = Min$ + ":" + Sec$
If Len(Entry2$) = 4 Then
    D$ = "0" + Entry2$
Else
    D$ = Entry2$
End If
If Len(Entry2$) = 3 Then
    D$ = "00" + Entry2$
Else
    D$ = Entry2$
End If
D$ = Entry2$
GetTrackTime = Trim$(D$)
MMCOntrol1.TimeFormat = 10
CDMin = Val(Min$)
CDSec = Val(Sec$)

End Function


Function GetRunningTime(MMCOntrol1 As Object) As String
Dim E As Long, M As String, S As String, Length&, Min$, Sec$, D As Long, Entry2$
MMCOntrol1.TimeFormat = 2
Length& = MMCOntrol1.Position - MMCOntrol1.TrackPosition
Min$ = Str$(Length& And &HFF)
Sec$ = LTrim$(Str$((Length& And 65280) / 256))
If Len(Sec$) = 3 Then
    D = Val(Min$) - 1
    Min$ = LTrim$(Str$(D))
    E = Val(Right$(Sec$, 2)) + 4
    
    Sec$ = LTrim$(Str$(E))
End If

M = Min$
'If Len(Min$) = 1 Then M = "0" + Min$ Else M = Min$
'If Val(M) = 0 Then M = "00"
'If Val(M) = 1 Then M = "01"
'If Val(M) = 2 Then M = "02"
'If Val(M) = 3 Then M = "03"
'If Val(M) = 4 Then M = "04"
'If Val(M) = 5 Then M = "05"
'If Val(M) = 6 Then M = "06"
'If Val(M) = 7 Then M = "07"
'If Val(M) = 8 Then M = "08"
'If Val(M) = 9 Then M = "09"
If Len(Sec$) = 1 Then
    S = "0" + Sec$
Else
    If Len(Sec$) = 3 Then
        S = Mid$(Sec$, 2)
    Else
        S = Sec$
    End If
End If
TMin = Val(M)
TSec = Val(S)
Entry2$ = LTrim$(M) + ":" + LTrim$(S)
MMCOntrol1.TimeFormat = 10
GetRunningTime = Entry2$
End Function



Function GetRawRunningTime(MMCOntrol1 As Object) As String
Dim E As Long, M As String, S As String, Length&, Min$, Sec$, D As Long, Entry2$
MMCOntrol1.TimeFormat = 2
Length& = MMCOntrol1.Position - MMCOntrol1.TrackPosition
Min$ = Str$(Length& And &HFF)
Sec$ = LTrim$(Str$((Length& And 65280) / 256))
If Len(Sec$) = 3 Then
    D = Val(Min$) - 1
    Min$ = LTrim$(Str$(D))
    E = Val(Right$(Sec$, 2)) + 4
    
    Sec$ = LTrim$(Str$(E))
End If

M = Min$
'If Len(Min$) = 1 Then M = "0" + Min$ Else M = Min$
'If Val(M) = 0 Then M = "0"
'If Val(M) = 1 Then M = "1"
'If Val(M) = 2 Then M = "2"
'If Val(M) = 3 Then M = "3"
'If Val(M) = 4 Then M = "4"
'If Val(M) = 5 Then M = "5"
'If Val(M) = 6 Then M = "6"
'If Val(M) = 7 Then M = "7"
'If Val(M) = 8 Then M = "8"
'If Val(M) = 9 Then M = "9"
If Len(Sec$) = 1 Then
    S = "0" + Sec$
Else
    If Len(Sec$) = 3 Then
        S = Mid$(Sec$, 2)
    Else
        S = Sec$
    End If
End If
Entry2$ = M + S
MMCOntrol1.TimeFormat = 10
GetRawRunningTime = Entry2$
End Function


Function GetRawTrackTime(MMCOntrol1 As Object) As String
Dim Length&, Entry2$, Min$, Sec$, D$
MMCOntrol1.TimeFormat = 2
Length& = MMCOntrol1.TrackLength
Min$ = Str$(Length& And &HFF)
Sec$ = LTrim$(Str$((Length& And 65280) / 256))
Entry2$ = Min$ + Sec$
If Len(Sec$) = 1 Then Entry2$ = Min$ + "0" + Sec$
GetRawTrackTime = Entry2$
MMCOntrol1.TimeFormat = 10

End Function

Public Function myReadINI(inifile, inisection, inikey, iniDefault)
'Fail fracefully if no file / wrong file is specified.
'If no section (appname), default is first appname
'if no key, default is first key


Dim lpApplicationName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpReturnedString As String
Dim nSize As Long
Dim lpFileName As String
Dim retval As Long
Dim Filename As String
lpDefault = Space$(254)
lpDefault = iniDefault

lpReturnedString = Space$(254)

nSize = 254
lpFileName = inifile
lpApplicationName = inisection
lpKeyName = inikey
Filename = lpFileName
retval = GetPrivateProfileString _
(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)
myReadINI = lpReturnedString
End Function


Public Function myWriteINI(inifile As String, inisection As String, inikey As String, Info As String) As String
Dim retval As Long
retval = WritePrivateProfileString(inisection, inikey, Info, inifile)
myWriteINI = LTrim$(Str$(retval))
End Function

Public Function GenCDSerial(MMCOntrol1 As Object) As Long
Const MCI_FORMAT_MILLISECONDS = 0
Const MCI_FORMAT_MSF = 2
Const MCI_FORMAT_TMSF = 10
'MCI_Format :0,2,10 are the only supported formats for CD
Dim Compat As Integer
Dim i As Integer
Dim dwtotal&, dwtemp&
Dim byte0%, byte1%, byte2%, byte3%
'compat = 0 for EXISTING code
'Compat = 1 for use with CDPLAYER.EXE
Compat = 1
MMCOntrol1.Notify = False
MMCOntrol1.Wait = True
MMCOntrol1.Shareable = True
If MMCOntrol1.Error <> 0 Then
    MsgBox MMCOntrol1.ErrorMessage
    Exit Function
End If
MMCOntrol1.TimeFormat = MCI_FORMAT_MSF
dwtotal& = 0
For i = 1 To MMCOntrol1.Tracks
    DoEvents
    MMCOntrol1.Track = i
    dwtemp& = MMCOntrol1.TrackPosition
    byte0% = dwtemp& And &HFF&
    byte1% = (dwtemp& And &HFF00&) \ &H100
    byte2% = (dwtemp& And &HFF0000) \ &H10000
    byte3% = (dwtemp& And &H7F000000) \ &H1000000
    If (dwtemp& And &H80000000) <> 0 Then
        ' put sign bit back into byte4
        byte3 = byte3 + &H80
    End If
    dwtemp& = byte0% * &H10000 + byte1% * &H100 + byte2%
    dwtotal& = dwtotal& + dwtemp&
Next i
If MMCOntrol1.Tracks < 3 Then
    dwtotal& = dwtotal& + msf2frames(MMCOntrol1.Length) + Compat
End If
GenCDSerial = dwtotal&

End Function
Function msf2frames(msf As Long) As Long
Rem From the KnowledgeBase
Rem    byte1 = MMControl1.Position And &HFF&
Rem    byte2 = (MMControl1.Position And &HFF00&) \ &H100
Rem    byte3 = (MMControl1.Position And &HFF0000) \ &H10000
Rem    byte4 = (MMControl1.Position And &H7F000000) \ &H1000000
Rem    If (MMControl1.Position And &H80000000) <> 0 Then
Rem       ' put sign bit back into byte4
Rem       byte4 = byte4 + &H80
Rem    End If
    Dim byte0, byte1, byte2, byte3 As Integer
    Dim Min, Sec, fra As Integer
    byte0 = msf And &HFF&
    byte1 = (msf And &HFF00&) \ &H100
    byte2 = (msf And &HFF0000) \ &H10000
    byte3 = (msf And &H7F000000) \ &H1000000
    If (msf And &H80000000) <> 0 Then
       ' put sign bit back into byte4
       byte3 = byte3 + &H80
    End If
    Min = byte0
    Sec = byte1
    fra = byte2
    msf2frames = (Min * 60 + Sec) * 75 + fra
    
End Function


Function Z_Trim(String1 As String) As String
Dim A As Integer
For A = 1 To Len(String1)
    If Mid$(String1, A, 1) = Chr$(0) Then Exit For
Next A
Z_Trim = RTrim$(Left$(String1, A - 1))

End Function


