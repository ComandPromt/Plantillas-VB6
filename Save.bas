Attribute VB_Name = "Module1"
Option Explicit

Public Const CALLBACK_WINDOW = &H10000
Public Const MMIO_READ = &H0
Public Const MMIO_FINDCHUNK = &H10
Public Const MMIO_FINDRIFF = &H20
Public Const MM_WOM_DONE = &H3BD
Public Const SEEK_CUR = 1
Public Const SEEK_END = 2
Public Const SEEK_SET = 0
Public Const TIME_BYTES = &H4
Public Const WHDR_DONE = &H1

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

Type MMTIME
        wType As Long
        u As Long
        x As Long
End Type

Declare Function waveOutGetPosition Lib "winmm.dll" (ByVal hWaveOut As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Declare Function waveOutOpen Lib "winmm.dll" (hWaveOut As Long, ByVal uDeviceID As Long, ByVal format As String, ByVal dwCallback As Long, ByRef fPlaying As Boolean, ByVal dwFlags As Long) As Long
Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal x As Long, ByVal uFlags As Long) As Long
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Declare Function mmioReadString Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Declare Function mmioSeek Lib "winmm.dll" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Declare Sub CopyStructFromString Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal source As String, ByVal cb As Long)
Declare Function PostWavMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef hdr As WAVEHDR) As Long

Declare Function CallWindowProc Lib "user32" Alias _
"CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal msg As Long, _
    ByVal wParam As Long, ByRef lParam As WAVEHDR) As Long

Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = -4
Dim lpPrevWndProc As Long

Const NUM_BUFFERS = 5
Const BUFFER_SECONDS = 0.1

Dim rc As Long           ' Return code
Dim hmmioIn As Long      ' file handle
Dim dataOffset As Long  ' start of audio data in wave file
Dim audioLength As Long  ' number of bytes in audio data
Dim pFormat As Long      ' pointer to wave format
Dim formatBuffer As String * 50 ' buffer to hold the wave format
Dim startPos As Long    ' sample where we started playback from
Dim format As WAVEFORMAT    ' waveformat structure
Dim i As Long       ' loop control variable
Dim j As Long       ' loop control variable
Dim hmem(1 To NUM_BUFFERS) As Long      ' memory handles
Dim pmem(1 To NUM_BUFFERS) As Long      ' memory pointers
Dim hdr(1 To NUM_BUFFERS) As WAVEHDR    ' wave headers
Dim bufferSize As Long      ' size of output buffers
Dim fPlaying As Boolean     ' is file currently playing
Dim fFileOpen As Boolean    ' is file currently open
Dim hWaveOut As Long        ' waveout handle
Dim msg As String * 250     ' message buffer
Dim hwnd As Long            ' window handle

Public Sub Initialize(hwndIn As Long)
    hwnd = hwndIn
    lpPrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    fPlaying = False
    fFileOpen = False
    startPos = 0
End Sub

Public Sub CloseFile()
    mmioClose hmmioIn, 0
    fFileOpen = False
End Sub

Public Sub OpenFile(soundfile As String)

    Dim mmckinfoParentIn As MMCKINFO
    Dim mmckinfoSubchunkIn As MMCKINFO
    Dim mmioinf As mmioinfo
    

    ' close previously open file (if any)
    CloseFile
    
    If (soundfile = "") Then
        Exit Sub
    End If
        
    ' Open the input file
    hmmioIn = mmioOpen(soundfile, mmioinf, MMIO_READ)
    If (hmmioIn = 0) Then
        MsgBox "Error opening input file, rc = " & mmioinf.wErrorRet
        Exit Sub
    End If

    ' Check if this is a wave file
    mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
    rc = mmioDescendParent(hmmioIn, mmckinfoParentIn, 0, MMIO_FINDRIFF)
    If (rc <> MMSYSERR_NOERROR) Then
        CloseFile
        MsgBox "Not a wave file"
        Exit Sub
    End If

    ' Get format info
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("fmt", 0)
    rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
    If (rc <> MMSYSERR_NOERROR) Then
        CloseFile
        MsgBox "Couldn't get format chunk"
        Exit Sub
    End If
    rc = mmioReadString(hmmioIn, formatBuffer, mmckinfoSubchunkIn.ckSize)
    If (rc = -1) Then
        CloseFile
        MsgBox "Error reading format"
        Exit Sub
    End If
    rc = mmioAscend(hmmioIn, mmckinfoSubchunkIn, 0)
    CopyStructFromString format, formatBuffer, Len(format)
    
    ' Find the data subchunk
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("data", 0)
    rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
    If (rc <> MMSYSERR_NOERROR) Then
        CloseFile
        MsgBox "Couldn't get data chunk"
        Exit Sub
    End If
    dataOffset = mmioSeek(hmmioIn, 0, SEEK_CUR)
    
    ' Get the length of the audio
    audioLength = mmckinfoSubchunkIn.ckSize
    
    ' Allocate audio buffers
    bufferSize = format.nSamplesPerSec * format.nBlockAlign * format.nChannels * BUFFER_SECONDS
    bufferSize = bufferSize - (bufferSize Mod format.nBlockAlign)
    For i = 1 To (NUM_BUFFERS)
        GlobalFree hmem(i)
        hmem(i) = GlobalAlloc(0, bufferSize)
        pmem(i) = GlobalLock(hmem(i))
    Next
    
    fFileOpen = True
    
End Sub

Public Function play() As Boolean

    If (fPlaying) Then
        play = True
        Exit Function
    End If
    
    rc = waveOutOpen(hWaveOut, 0, formatBuffer, hwnd, True, CALLBACK_WINDOW)
    
    If (rc <> MMSYSERR_NOERROR) Then
        waveOutGetErrorText rc, msg, Len(msg)
        MsgBox msg
        play = False
        Exit Function
    End If

    For i = 1 To NUM_BUFFERS
        hdr(i).lpData = pmem(i)
        hdr(i).dwBufferLength = bufferSize
        hdr(i).dwFlags = 0
        hdr(i).dwLoops = 0
        rc = waveOutPrepareHeader(hWaveOut, hdr(i), Len(hdr(i)))
        If (rc <> MMSYSERR_NOERROR) Then
            waveOutGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    fPlaying = True
    play = True

    startPos = mmioSeek(hmmioIn, 0, SEEK_CUR) - dataOffset
    
    For i = 1 To NUM_BUFFERS
        PostWavMessage hwnd, MM_WOM_DONE, 0, hdr(i)
    Next
    
End Function

Public Sub StopPlay()
    fPlaying = False
    FileSeek Position()
    waveOutReset hWaveOut
End Sub

Public Function Length() As Long
    Length = audioLength \ format.nBlockAlign
End Function

Public Function FileSeek(Position As Long) As Boolean
    Dim bytepos As Long
    FileSeek = False
    bytepos = Position * format.nBlockAlign
    If (fFileOpen = False) Or (bytepos < 0) Or (bytepos >= audioLength) Then
        Exit Function
    End If
    rc = mmioSeek(hmmioIn, bytepos + dataOffset, SEEK_SET)
    If (rc = MMSYSERR_NOERROR) Then
        FileSeek = True
    End If
    startPos = rc
End Function

Public Function Position() As Long
    Dim tm As MMTIME
    tm.wType = TIME_BYTES
    rc = waveOutGetPosition(hWaveOut, tm, Len(tm))
    If (rc = MMSYSERR_NOERROR) Then
        Position = (startPos + tm.u) \ format.nBlockAlign
    Else
        Position = (mmioSeek(hmmioIn, 0, SEEK_CUR) - dataOffset + bufferSize * NUM_BUFFERS) \ format.nBlockAlign
    End If
End Function

Public Function Playing() As Boolean
    Dim tm As MMTIME
    tm.wType = TIME_BYTES
    rc = waveOutGetPosition(hWaveOut, tm, Len(tm))
    If (rc = MMSYSERR_NOERROR) Then
        Playing = True
    Else
        Playing = False
    End If
End Function

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByRef wavhdr As WAVEHDR) As Long
    Static dataRemaining As Long
    If (uMsg = MM_WOM_DONE) Then
        If (fPlaying = True) Then
            dataRemaining = (dataOffset + audioLength - mmioSeek(hmmioIn, 0, SEEK_CUR))
            If (bufferSize < dataRemaining) Then
                rc = mmioRead(hmmioIn, wavhdr.lpData, bufferSize)
            Else
                rc = mmioRead(hmmioIn, wavhdr.lpData, dataRemaining)
                fPlaying = False
            End If
            wavhdr.dwBufferLength = rc
            rc = waveOutWrite(hWaveOut, wavhdr, Len(wavhdr))
        Else
            For i = 1 To NUM_BUFFERS
                waveOutUnprepareHeader hWaveOut, hdr(i), Len(hdr(i))
            Next
            waveOutClose hWaveOut
        End If
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, wavhdr)
End Function

