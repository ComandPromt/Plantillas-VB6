Attribute VB_Name = "Module1"
Option Explicit

Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Const CALLBACK_FUNCTION = &H30000
Public Const MMIO_READ = &H0
Public Const MMIO_FINDCHUNK = &H10
Public Const MMIO_FINDRIFF = &H20
Public Const MM_WOM_DONE = &H3BD

Type mmioinfo
   dwFlags As Long
   fccIOProc As Long
   pIOProc As Long
   wErrorRet As Long
   htask As Long
   cchBuffer As Long
   pchBuffer As String
   pchNext As String
   pchEndRead As String
   pchEndWrite As String
   lBufOffset As Long
   lDiskOffset As Long
   adwInfo(4) As Long
   dwReserved1 As Long
   dwReserved2 As Long
   hmmio As Long
End Type

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

Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type


Declare Function waveOutOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function waveOutAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal x As Long, ByVal uFlags As Long) As Long
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Declare Function mmioReadFormat Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByRef pch As WAVEFORMAT, ByVal cch As Long) As Long
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)

Dim rc As Long
Dim msg As String * 200

' variables for managing wave file
Public format As WAVEFORMAT
Dim hmmioOut As Long
Dim mmckinfoParentIn As MMCKINFO
Dim mmckinfoSubchunkIn As MMCKINFO
Dim hWaveOut As Long
Dim bufferIn As Long
Dim hmem As Long
Dim outHdr As WAVEHDR
Public numSamples As Long
Public drawFrom As Long
Public drawTo As Long
Public fFileLoaded As Boolean
Public fPlaying As Boolean

Sub waveOutProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
' Wave IO Callback function
   If (uMsg = MM_WOM_DONE) Then
      fPlaying = False
   End If
End Sub

Sub CloseWaveOut()
' Close the waveout device
    rc = waveOutReset(hWaveOut)
    rc = waveOutUnprepareHeader(hWaveOut, outHdr, Len(outHdr))
    rc = waveOutClose(hWaveOut)
End Sub

Sub LoadFile(inFile As String)
' Load wavefile into memory

   Dim hmmioIn As Long
   Dim mmioinf As mmioinfo
   
   fFileLoaded = False
   
   If (inFile = "") Then
       GlobalFree (hmem)
       Exit Sub
   End If
       
   ' Open the input file
   hmmioIn = mmioOpen(inFile, mmioinf, MMIO_READ)
   If hmmioIn = 0 Then
       MsgBox "Error opening input file, rc = " & mmioinf.wErrorRet
       Exit Sub
   End If
   
   ' Check if this is a wave file
   mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
   rc = mmioDescendParent(hmmioIn, mmckinfoParentIn, 0, MMIO_FINDRIFF)
   If (rc <> 0) Then
       rc = mmioClose(hmmioOut, 0)
       MsgBox "Not a wave file"
       Exit Sub
   End If
   
   ' Get format info
   mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("fmt", 0)
   rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
   If (rc <> 0) Then
       rc = mmioClose(hmmioOut, 0)
       MsgBox "Couldn't get format chunk"
       Exit Sub
   End If
   rc = mmioReadFormat(hmmioIn, format, Len(format))
   If (rc = -1) Then
      rc = mmioClose(hmmioOut, 0)
      MsgBox "Error reading format"
      Exit Sub
   End If
   rc = mmioAscend(hmmioIn, mmckinfoSubchunkIn, 0)
   
   ' Find the data subchunk
   mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("data", 0)
   rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
   If (rc <> 0) Then
      rc = mmioClose(hmmioOut, 0)
      MsgBox "Couldn't get data chunk"
      Exit Sub
   End If
   
   ' Allocate soundbuffer and read sound data
   GlobalFree hmem
   hmem = GlobalAlloc(&H40, mmckinfoSubchunkIn.ckSize)
   bufferIn = GlobalLock(hmem)
   rc = mmioRead(hmmioIn, bufferIn, mmckinfoSubchunkIn.ckSize)
   
   numSamples = mmckinfoSubchunkIn.ckSize / format.nBlockAlign
   
   ' Close file
   rc = mmioClose(hmmioOut, 0)
   
   fFileLoaded = True
    
End Sub

Sub play(ByVal soundcard As Integer)
' Send audio buffer to wave output

    rc = waveOutOpen(hWaveOut, soundcard, format, AddressOf waveOutProc, 0, CALLBACK_FUNCTION)
    If (rc <> 0) Then
      GlobalFree (hmem)
      waveOutGetErrorText rc, msg, Len(msg)
      MsgBox msg
      Exit Sub
    End If

    outHdr.lpData = bufferIn + (drawFrom * format.nBlockAlign)
    outHdr.dwBufferLength = (drawTo - drawFrom) * format.nBlockAlign
    outHdr.dwFlags = 0
    outHdr.dwLoops = 0

    rc = waveOutPrepareHeader(hWaveOut, outHdr, Len(outHdr))
    If (rc <> 0) Then
      waveOutGetErrorText rc, msg, Len(msg)
      MsgBox msg
    End If

    rc = waveOutWrite(hWaveOut, outHdr, Len(outHdr))
    If (rc <> 0) Then
      GlobalFree (hmem)
    Else
      fPlaying = True
      Form1.Timer1.Enabled = True
    End If
End Sub

Sub StopPlay()
   waveOutReset (hWaveOut)
End Sub

Sub GetStereo16Sample(ByVal sample As Long, ByRef leftVol As Double, ByRef rightVol As Double)
' These subs obtain a PCM sample and converts it into volume levels from (-1 to 1)

   Dim sample16 As Integer
   Dim ptr As Long
   
   ptr = sample * format.nBlockAlign + bufferIn
   CopyStructFromPtr sample16, ptr, 2
   leftVol = sample16 / 32768
   CopyStructFromPtr sample16, ptr + 2, 2
   rightVol = sample16 / 32768

End Sub

Sub GetStereo8Sample(ByVal sample As Long, ByRef leftVol As Double, ByRef rightVol As Double)

   Dim sample8 As Byte
   Dim ptr As Long
   
   ptr = sample * format.nBlockAlign + bufferIn
   CopyStructFromPtr sample8, ptr, 1
   leftVol = (sample8 - 128) / 128
   CopyStructFromPtr sample8, ptr + 1, 1
   rightVol = (sample8 - 128) / 128

End Sub

Sub GetMono16Sample(ByVal sample As Long, ByRef leftVol As Double)

   Dim sample16 As Integer
   Dim ptr As Long
   
   ptr = sample * format.nBlockAlign + bufferIn
   CopyStructFromPtr sample16, ptr, 2
   leftVol = sample16 / 32768

End Sub

Sub GetMono8Sample(ByVal sample As Long, ByRef leftVol As Double)

   Dim sample8 As Byte
   Dim ptr As Long
   
   ptr = sample * format.nBlockAlign + bufferIn
   CopyStructFromPtr sample8, ptr, 1
   leftVol = (sample8 - 128) / 128

End Sub
