Attribute VB_Name = "Module1"
Public Const CALLBACK_WINDOW = &H10000
Public Const MMIO_READ = &H0
Public Const MMIO_FINDCHUNK = &H10
Public Const MMIO_FINDRIFF = &H20
Public Const MM_WOM_DONE = &H3BD
Public Const MMSYSERR_NOERROR = 0
Public Const SEEK_CUR = 1
Public Const SEEK_END = 2
Public Const SEEK_SET = 0
Public Const TIME_BYTES = &H4
Public Const WHDR_DONE = &H1

Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal x As Long, ByVal uFlags As Long) As Long
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Declare Function mmioReadString Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Declare Function mmioSeek Lib "winmm.dll" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long

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

Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type


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
    CopyStructFromString Format, formatBuffer, Len(Format)
    
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
    bufferSize = Format.nSamplesPerSec * Format.nBlockAlign * Format.nChannels * BUFFER_SECONDS
    bufferSize = bufferSize - (bufferSize Mod Format.nBlockAlign)
    For I = 1 To (NUM_BUFFERS)
        GlobalFree hmem(I)
        hmem(I) = GlobalAlloc(0, bufferSize)
        pmem(I) = GlobalLock(hmem(I))
    Next
    
    fFileOpen = True
    
End Sub
