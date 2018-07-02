Attribute VB_Name = "SoundMeter"

Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32
Public i As Integer, j As Integer, msg As String * 200, hWaveIn As Long
Public Const NUM_BUFFERS = 2
Public wformat As WAVEFORMAT, hmem(NUM_BUFFERS) As Long, inHdr(NUM_BUFFERS) As WAVEHDR
Public BUFFER_SIZE
Public Const DEVICEID = 0

'=== SB Live Card ===============================================================
Public rc As Long                      ' return code
Public OK As Boolean                   ' boolean return code
Public volume As Long                      ' volume value
Public volHmem As Long                     ' handle to volume memory
Public audbytearray As AUDINPUTARRAY
Public audByteHigh As AUDINPUTARRAY
Public posval As Integer
Public tempval As Integer
Public buffaddress As Long
Public retVal As Integer

Public Const CALLBACK_FUNCTION = &H30000
Public Const MM_WIM_DATA = &H3C0
Public Const WHDR_DONE = &H1         '  done bit
Public Const GMEM_FIXED = &H0         ' Global Memory Flag used by GlobalAlloc functin
Type WAVEHDR
   lpData As Long
   dwBufferLength As Long
   dwBytesRecorded As Long
   dwUser As Long
   dwFlags As Long
   dwLoops As Long
   lpNext As Long
   Reserved As Long
End Type
Type WAVEINCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * 32
   dwFormats As Long
   wChannels As Integer
End Type
Type WAVEFORMAT
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type
Type AUDINPUTARRAY
    bytes(5000) As Byte
End Type
Public Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Public Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal Err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long

Public Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Public Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long

Public Function getVolume(pbuff As Long) As Integer
On Local Error Resume Next
Err.Clear

Dim n As Integer
Dim AQZ_H As Integer, AQZ_M As Integer, AQZ_S As Integer

AQZ_H = Hour(Now)
AQZ_M = Minute(Now)
AQZ_S = Second(Now)

Do While Not inHdr(0).dwFlags And WHDR_DONE And Not TimeSerial(Hour(Now), Minute(Now), Second(Now)) >= TimeSerial(AQZ_H, AQZ_M, AQZ_S + XA)
    ' perhaps I ought to put a time limit on this bit!
Loop

iValue.Caption = CStr(0)
iValue.Refresh
CopyStructFromPtr audbytearray, inHdr(0).lpData, inHdr(0).dwBufferLength
rc = waveInAddBuffer(hWaveIn, inHdr(0), Len(inHdr(0)))

tempval = 0
posval = 0
For n = 0 To BUFFER_SIZE - 1
    posval = audbytearray.bytes(n) - 128
    If posval < 0 Then posval = 0 - posval
    If posval > tempval Then tempval = posval
Next n

getVolume = tempval
pbuff = inHdr(0).lpData

End Function
Public Function StartInput() As Boolean
On Error GoTo Err1

wformat.wFormatTag = 1
wformat.nChannels = 2
wformat.wBitsPerSample = 8
wformat.nSamplesPerSec = 8000
wformat.nBlockAlign = wformat.nChannels * wformat.wBitsPerSample / 8
wformat.nAvgBytesPerSec = wformat.nSamplesPerSec * wformat.nBlockAlign
wformat.cbSize = 0

For i = 0 To NUM_BUFFERS - 1
    hmem(i) = GlobalAlloc(&H40, BUFFER_SIZE)
    inHdr(i).lpData = GlobalLock(hmem(i))
    inHdr(i).dwBufferLength = BUFFER_SIZE
    inHdr(i).dwFlags = 0
    inHdr(i).dwLoops = 0
Next

rc = waveInOpen(hWaveIn, DEVICEID, wformat, 0, 0, 0)
If rc <> 0 Then
    waveInGetErrorText rc, msg, Len(msg)
    MsgBox msg
    StartInput = False
    Exit Function
End If

For i = 0 To NUM_BUFFERS - 1
    rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
    If (rc <> 0) Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg
    End If
Next

For i = 0 To NUM_BUFFERS - 1
    rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
    If (rc <> 0) Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg
    End If
Next

rc = waveInStart(hWaveIn)
StartInput = True
Exit Function

Err1:
    StartInput = False
End Function
Public Function StopInput() As Integer
On Error GoTo Err1

waveInReset hWaveIn
waveInStop hWaveIn
For i = 0 To NUM_BUFFERS - 1
    waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
    GlobalFree hmem(i)
Next
waveInClose hWaveIn
GlobalFree volHmem
StopInput = 0
frmMain.scopeBox.Cls
frmMain.ProgressBar1.Value = 0
frmMain.ProgressBar2.Value = 0
Exit Function

Err1:
    StopInput = 1
End Function
