Attribute VB_Name = "WAVEMIX"
Option Explicit
'------------------------------------------------------------
' WAVEMIX.BAS
' This module contains declarations for all the functions
' in the WaveMix DLL, and provides some higher-level Basic
' functions to make using WaveMix simpler.
'------------------------------------------------------------
Global hWaveMix As Long
Global lpWaveMix() As Long
Global WaveHandle As Long
Global WAVMIX_Quiet As Integer

Global Const WAVEMIX_MAXCHANNELS = 8

Type tChannelInfo
    Loops As Long
    WaveFile As String
End Type


Type WAVEMIXINFO
   wSize As Integer
   bVersionMajor As String * 1
   bVersionMinor As String * 1
   szDate(12) As String
   dwFormats As Long
End Type

Type MIXCONFIG
    wSize As Integer
    dwFlagsLo As Integer
    dwFlagsHi As Integer
    wChannels As Integer
    wSamplingRate As Integer
End Type

Private Type MIXPLAYPARAMS
    wSize         As Integer
    hMixSessionLo As Integer
    hMixSessionHi As Integer
    iChannelLo    As Integer
    iChannelHi    As Integer
    lpMixWaveLo   As Integer
    lpMixWaveHi   As Integer
    hWndNotifyLo  As Integer
    hWndNotifyHi  As Integer
    dwFlagsLo     As Integer
    dwFlagsHi     As Integer
    wLoops        As Integer
End Type

Declare Function WaveMixInit Lib "WAVMIX32.DLL" () As Long
Declare Function WaveMixConfigureInit Lib "WAVMIX32.DLL" (lpConfig As MIXCONFIG) As Long
Declare Function WaveMixActivate Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, ByVal fActivate As Integer) As Long
Declare Function WaveMixOpenWave Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, szWaveFilename As Any, ByVal hInst As Long, ByVal dwFlags As Long) As Long
Declare Function WaveMixOpenChannel Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, ByVal iChannel As Long, ByVal dwFlags As Long) As Long
Declare Function WaveMixPlay Lib "WAVMIX32.DLL" (lpMixPlayParams As Any) As Integer
Declare Function WaveMixFlushChannel Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, ByVal iChannel As Integer, ByVal dwFlags As Long) As Integer
Declare Function WaveMixCloseChannel Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, ByVal iChannel As Integer, ByVal dwFlags As Long) As Integer
Declare Function WaveMixFreeWave Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, ByVal lpMixWave As Long) As Integer
Declare Function WaveMixCloseSession Lib "WAVMIX32.DLL" (ByVal hMixSession As Long) As Integer
Declare Sub WaveMixPump Lib "WAVMIX32.DLL" ()
Declare Function WaveMixGetInfo Lib "WAVMIX32.DLL" (lpWaveMixInfo As WAVEMIXINFO) As Integer

Private Function HiWord(ByVal l As Long) As Integer
    l = l \ &H10000
    
    HiWord = Val("&H" & Hex$(l))
End Function

Private Function LoWord(ByVal l As Long) As Integer
    l = l And &HFFFF&
    
    LoWord = Val("&H" & Hex$(l))
End Function


Function WAVMIX_AddFile(FileName As String) As Integer
'------------------------------------------------------------
' Open a wave file and assign it to the next available
' channel.
'------------------------------------------------------------
Dim wRtn As Long

    WAVMIX_AddFile = False
    If WAVMIX_Quiet Then Exit Function
    If WaveHandle + 1 = WAVEMIX_MAXCHANNELS Then Exit Function

    ReDim Preserve lpWaveMix(WaveHandle)
    lpWaveMix(WaveHandle) = WaveMixOpenWave(hWaveMix, ByVal FileName, 0, 0)
    wRtn = WaveMixOpenChannel(hWaveMix, WaveHandle, 0)
    WAVMIX_AddFile = WaveHandle
    WaveHandle = WaveHandle + 1
End Function

Sub WAVMIX_SetFile(FileName As String, AChannel As Long)
'------------------------------------------------------------
' Assign a new wave file, FileName, to the specified channel,
' AChannel.  If this channel is currently assigned another
' wave file, stop playing the channel and free the active
' wave file.
'------------------------------------------------------------
Dim wRtn As Long
    
    If WAVMIX_Quiet Then Exit Sub
    
    If AChannel > UBound(lpWaveMix) Then
        ReDim Preserve lpWaveMix(AChannel)
        WaveHandle = AChannel
    End If
    
    ' If another wave is currently assigned to this
    ' channel, free it.
    If lpWaveMix(AChannel) <> 0 Then
        WAVMIX_StopChannel AChannel
        wRtn = WaveMixFreeWave(hWaveMix, lpWaveMix(AChannel))
    End If
    
    ' Open the new wave and assign it to this channel.
    lpWaveMix(AChannel) = WaveMixOpenWave(hWaveMix, ByVal FileName, 0, 0)
    wRtn = WaveMixOpenChannel(hWaveMix, AChannel, 0)
End Sub


Sub WAVMIX_Close()
'------------------------------------------------------------
' Stop playing all channels and free all wave files, then
' close down this WaveMix session.
'------------------------------------------------------------
Dim wRtn As Long
Dim i As Integer, rc As Integer
    
    If WAVMIX_Quiet Then Exit Sub

    If (hWaveMix <> 0) Then
        For i = 0 To UBound(lpWaveMix)
            If lpWaveMix(i) <> 0 Then
                WAVMIX_StopChannel CLng(i)
                rc = WaveMixFreeWave(hWaveMix, lpWaveMix(i))
            End If
        Next
        wRtn = WaveMixCloseSession(hWaveMix)
        hWaveMix = 0
    End If
End Sub

Function WAVMIX_InitMixer() As Integer
'------------------------------------------------------------
' Initialize and activate the WaveMix DLL.
'------------------------------------------------------------
Dim wRtn As Long
Dim config As MIXCONFIG

    If WAVMIX_Quiet Then Exit Function

    WaveHandle = 0
    ReDim lpWaveMix(0)
    ChDir App.Path
    
    config.wSize = Len(config)
    config.dwFlagsHi = 1
    config.dwFlagsLo = 0
    'Allow stereo sound
    config.wChannels = 2
    hWaveMix = WaveMixConfigureInit(config)
    wRtn = WaveMixActivate(hWaveMix, True)

    If (wRtn <> 0) Then
        WAVMIX_InitMixer = False
        Call WaveMixCloseSession(hWaveMix)
        hWaveMix = 0
    Else
        WAVMIX_InitMixer = True
    End If
End Function

Sub WAVMIX_StopChannel(ByVal ChannelNum As Long)
'------------------------------------------------------------
' Stop playing the specified channel.
'------------------------------------------------------------
Dim rc As Integer

    If WAVMIX_Quiet Then Exit Sub
    If (hWaveMix = 0) Then Exit Sub
    
    rc = WaveMixFlushChannel(hWaveMix, ChannelNum, 0)
End Sub

Sub WAVMIX_Activate(Activate As Long)
'------------------------------------------------------------
' Activate the WaveMix DLL.
'------------------------------------------------------------
Dim rc As Integer

    If WAVMIX_Quiet Then Exit Sub
    If (hWaveMix = 0) Then Exit Sub

    rc = WaveMixActivate(hWaveMix, Activate)
End Sub

Sub WAVMIX_PlayChannel(ChannelNum As Long, LoopWave As Long)
'------------------------------------------------------------
' Play a specified channel, and indicate whether the sound
' should be looped.
'------------------------------------------------------------
Dim params As MIXPLAYPARAMS
Dim wRtn As Long

    If WAVMIX_Quiet Then Exit Sub
    If ChannelNum > UBound(lpWaveMix) Then Exit Sub
    If (hWaveMix = 0) Then Exit Sub

    params.wSize = Len(params)
    params.hMixSessionLo = LoWord(hWaveMix)
    params.hMixSessionHi = HiWord(hWaveMix)
    params.iChannelLo = LoWord(ChannelNum)
    params.iChannelHi = HiWord(ChannelNum)
    params.lpMixWaveLo = LoWord(lpWaveMix(ChannelNum))
    params.lpMixWaveHi = HiWord(lpWaveMix(ChannelNum))
    params.hWndNotifyLo = 0
    params.hWndNotifyHi = 0
    params.dwFlagsHi = 5
    params.dwFlagsLo = 0
    params.wLoops = LoopWave
    wRtn = WaveMixPlay(params)
End Sub

