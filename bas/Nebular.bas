Attribute VB_Name = "Nebular"


Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Const HIGHEST_VOLUME_SETTING = 65535
Public Const AUX_MAPPER = -1&
Public Const MAXPNAMELEN = 32
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2& ' separate left-right volume control
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&

Public Const MIXERLINE_COMPONENTTYPE_SRC_MIDIVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)

Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)

Public Const MIXERLINE_COMPONENTTYPE_SRC_I25InVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)

Public Const MIXERLINE_COMPONENTTYPE_SRC_TADVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)

Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
                             (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
               
Public Const MIXERLINE_COMPONENTTYPE_src_AUXVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)

Public Const MIXERLINE_COMPONENTTYPE_SRC_PSPKVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)

Public Const MIXERLINE_COMPONENTTYPE_SRC_MBOOST = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)

Public Const MIXERLINE_COMPONENTTYPE_SRC_LINEVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)

Public Const MIXERLINE_COMPONENTTYPE_SRC_CDVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)

Public Const CALLBACK_FUNCTION = &H30000
Public Const MMIO_READ = &H0
Public Const MMIO_FINDCHUNK = &H10
Public Const MMIO_FINDRIFF = &H20
Public Const MM_WOM_DONE = &H3BD
Public Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Public Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Public Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Public Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Public Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Public Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Public Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000  ' /* in 10ths */
Public Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000
Public Const AUXCAPS_CDAUDIO = 1  '  audio from internal CD-ROM drive
Public Const AUXCAPS_AUXIN = 2  '  audio from auxiliary input jacks
Public Const AUXCAPS_VOLUME = &H1   '  supports volume control
Public Const AUXCAPS_LRVOLUME = &H2 '  separate left-right volume control
Public Const MIXERCONTROL_CONTROLTYPE_SLIDER = _
               (MIXERCONTROL_CT_CLASS_SLIDER Or _
                MIXERCONTROL_CT_UNITS_SIGNED)

Public Const MIXERCONTROL_CONTROLTYPE_PAN = _
               (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)


Public Const MIXERCONTROL_CONTROLTYPE_FADER = _
               (MIXERCONTROL_CT_CLASS_FADER Or _
               MIXERCONTROL_CT_UNITS_UNSIGNED)

Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = _
               (MIXERCONTROL_CONTROLTYPE_FADER + 1)

Public Const MMSYSERR_NOERROR = 0
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)

Declare Function auxGetNumDevs Lib "winmm.dll" () As Long
Declare Function auxGetDevCaps Lib "winmm.dll" Alias "auxGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As AUXCAPS, ByVal uSize As Long) As Long
Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByRef lpdwVolume As Long) As Long
Declare Function auxOutMessage Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

Declare Function mixerClose Lib "winmm.dll" _
               (ByVal hmx As Long) As Long
   
Declare Function mixerGetControlDetails Lib "winmm.dll" _
               Alias "mixerGetControlDetailsA" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
   
Declare Function mixerGetDevCaps Lib "winmm.dll" _
               Alias "mixerGetDevCapsA" _
               (ByVal uMxId As Long, _
               ByVal pmxcaps As MIXERCAPS, _
               ByVal cbmxcaps As Long) As Long
   
Declare Function mixerGetID Lib "winmm.dll" _
               (ByVal hmxobj As Long, _
               pumxID As Long, _
               ByVal fdwId As Long) As Long
               
Declare Function mixerGetLineControls Lib "winmm.dll" _
               Alias "mixerGetLineControlsA" _
               (ByVal hmxobj As Long, _
               pmxlc As MIXERLINECONTROLS, _
               ByVal fdwControls As Long) As Long
               
Declare Function mixerGetLineInfo Lib "winmm.dll" _
               Alias "mixerGetLineInfoA" _
               (ByVal hmxobj As Long, _
               pmxl As MIXERLINE, _
               ByVal fdwInfo As Long) As Long
               
Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long

Declare Function mixerMessage Lib "winmm.dll" _
               (ByVal hmx As Long, _
               ByVal uMsg As Long, _
               ByVal dwParam1 As Long, _
               ByVal dwParam2 As Long) As Long
               
Declare Function mixerOpen Lib "winmm.dll" _
               (phmx As Long, _
               ByVal uMxId As Long, _
               ByVal dwCallback As Long, _
               ByVal dwInstance As Long, _
               ByVal fdwOpen As Long) As Long
               
Declare Function mixerSetControlDetails Lib "winmm.dll" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
               
Declare Sub CopyStructFromPtr Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (struct As Any, _
               ByVal ptr As Long, ByVal cb As Long)
               
Declare Sub CopyPtrFromStruct Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (ByVal ptr As Long, _
               struct As Any, _
               ByVal cb As Long)
               
Declare Function GlobalAlloc Lib "kernel32" _
               (ByVal wFlags As Long, _
               ByVal dwBytes As Long) As Long
               
Declare Function GlobalLock Lib "kernel32" _
               (ByVal hmem As Long) As Long
               
Declare Function GlobalFree Lib "kernel32" _
               (ByVal hmem As Long) As Long
               
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
Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal X As Long, ByVal uFlags As Long) As Long
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Declare Function mmioReadFormat Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByRef pch As WAVEFORMAT, ByVal cch As Long) As Long
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long


Dim rc As Long
Dim msg As String * 200

' variables for managing wave file
Public formatA As WAVEFORMAT
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
               
               
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Type VolumeSetting
    LeftVol As Integer
    rightVol As Integer
End Type

Type AUXCAPS
       wMid As Integer
       wPid As Integer
       vDriverVersion As Long
       szPname As String * MAXPNAMELEN
       wTechnology As Integer
       dwSupport As Long
End Type

Type MIXERCAPS
   wMid As Integer                   '  manufacturer id
   wPid As Integer                   '  product id
   vDriverVersion As Long            '  version of the driver
   szPname As String * MAXPNAMELEN   '  product name
   fdwSupport As Long                '  misc. support bits
   cDestinations As Long             '  count of destinations
End Type

Type MIXERCONTROL
   cbStruct As Long           '  size in Byte of MIXERCONTROL
   dwControlID As Long        '  unique control id for mixer device
   dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
   fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
   cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
   szShortName As String * MIXER_SHORT_NAME_CHARS  ' short name of control
   szName As String * MIXER_LONG_NAME_CHARS        ' long name of control
   lMinimum As Long           '  Minimum value
   lMaximum As Long           '  Maximum value
   Reserved(10) As Long       '  reserved structure space
   End Type

Type MIXERCONTROLDETAILS
   cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
   dwControlID As Long    '  control id to get/set details on
   cChannels As Long      '  number of channels in paDetails array
   Item As Long           '  hwndOwner or cMultipleItems
   cbDetails As Long      '  size of _one_ details_XX struct
   paDetails As Long      '  pointer to array of details_XX structs
End Type

Type MIXERCONTROLDETAILS_UNSIGNED
   dwValue As Long        '  value of the control (volume level)
End Type

Type MIXERLINE
   cbStruct As Long               '  size of MIXERLINE structure
   dwDestination As Long          '  zero based destination index
   dwSource As Long               '  zero based source index (if source)
   dwLineID As Long               '  unique line id for mixer device
   fdwLine As Long                '  state/information about line
   dwUser As Long                 '  driver specific information
   dwComponentType As Long        '  component type line connects to
   cChannels As Long              '  number of channels line supports
   cConnections As Long           '  number of connections (possible)
   cControls As Long              '  number of controls at this line
   szShortName As String * MIXER_SHORT_NAME_CHARS
   szName As String * MIXER_LONG_NAME_CHARS
   dwType As Long
   dwDeviceID As Long
   wMid  As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
End Type

Type MIXERLINECONTROLS
   cbStruct As Long       '  size in Byte of MIXERLINECONTROLS
   dwLineID As Long       '  line id (from MIXERLINE.dwLineID)
                          '  MIXER_GETLINECONTROLSF_ONEBYID or
   dwControl As Long      '  MIXER_GETLINECONTROLSF_ONEBYTYPE
   cControls As Long      '  count of controls pmxctrl points to
   cbmxctrl As Long       '  size in Byte of _one_ MIXERCONTROL
   pamxctrl As Long       '  pointer to first MIXERCONTROL array
End Type

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

Function GetMixerControl(ByVal hmixer As Long, _
                        ByVal componentType As Long, _
                        ByVal ctrlType As Long, _
                        ByRef mxc As MIXERCONTROL) As Boolean
                        
' This function attempts to obtain a mixer control. Returns True if successful.
   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType
   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   If (MMSYSERR_NOERROR = rc) Then
       mxlc.cbStruct = Len(mxlc)
       mxlc.dwLineID = mxl.dwLineID
       mxlc.dwControl = ctrlType
       mxlc.cControls = 1
       mxlc.cbmxctrl = Len(mxc)
       ' Allocate a buffer for the control
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       ' Get the control
       rc = mixerGetLineControls(hmixer, _
                                 mxlc, _
                                 MIXER_GETLINECONTROLSF_ONEBYTYPE)
       If (MMSYSERR_NOERROR = rc) Then
           GetMixerControl = True
           ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       Else
           GetMixerControl = False
       End If
       GlobalFree (hmem)
       Exit Function
   End If
   GetMixerControl = False
End Function

Function SetVolumeControl(ByVal hmixer As Long, _
                        mxc As MIXERCONTROL, _
                        ByVal Volume As Long) As Boolean
'This function sets the value for a volume control. Returns True if successful
   Dim mxcd As MIXERCONTROLDETAILS
   Dim Vol As MIXERCONTROLDETAILS_UNSIGNED
   
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.Item = 0
   mxcd.cbDetails = Len(Vol)
   ' Allocate memory for the control value buffer
   ' Len(Vol) = Number of Bytes
   ' &H40     = GMEM_ZEROINIT = Initialize to zero
   ' hmem     = Handle to the newly created memory object (Buffer)
   hmem = GlobalAlloc(&H40, Len(Vol))
   ' Lock the memory object (buffer) and return a pointer to the first byte.
   mxcd.paDetails = GlobalLock(hmem)
   Vol.dwValue = Volume
   ' Copy the data into the control value buffer
   CopyPtrFromStruct mxcd.paDetails, Vol, Len(Vol)
   ' Set the control value
   rc = mixerSetControlDetails(hmixer, _
                              mxcd, _
                              MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetVolumeControl = True
   Else
       SetVolumeControl = False
   End If
End Function

Function GetVolumeControlValue(ByVal hmixer As Long, mxc As MIXERCONTROL) As Long
'This function Gets the value for a volume control. Returns True if successful
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Vol As MIXERCONTROLDETAILS_UNSIGNED

    mxcd.cbStruct = Len(mxcd)
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cChannels = 1
    mxcd.Item = 0
    mxcd.cbDetails = Len(Vol)
    mxcd.paDetails = 0
    ' Allocate memory for the control value buffer
    ' Len(Vol) = Number of Bytes
    ' &H40     = GMEM_ZEROINIT = Initialize to zero
    ' hmem     = Handle to the newly created memory object (Buffer)
    hmem = GlobalAlloc(&H40, Len(Vol))
    ' Lock the memory object (buffer) and return a pointer to the first byte.
    mxcd.paDetails = GlobalLock(hmem)
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    ' Copy the data from the control value buffer
    CopyStructFromPtr Vol, mxcd.paDetails, Len(Vol)
    GlobalFree (hmem)
    If (MMSYSERR_NOERROR = rc) Then
       GetVolumeControlValue = Vol.dwValue
    Else
        GetVolumeControlValue = -1
    End If
End Function



Sub lCrossFader()
Vol1 = 100 - sldPan.Value ' Left
Vol2 = 100 - sldPan.Value ' Right
E = CrossFader.Value
F = 100 - E
If Check4.Value = 1 Then ' Half Fader Check
    LVol = (F * Val(Vol1) / 100) * 2
    RVOL = (E * Val(Vol2) / 100) * 2
    If LVol > (50 * Val(Vol1) / 100) * 2 Then
        LVol = (50 * Val(Vol1) / 100) * 2
    End If
    If RVOL > (50 * Val(Vol2) / 100) * 2 Then
        RVOL = (50 * Val(Vol2) / 100) * 2
    End If
Else
    LVol = (F * Val(Vol1) / 100)
    RVOL = (E * Val(Vol2) / 100)
End If
Label1.Caption = "Fader: " + LTrim$(Str$(LVol)) + " x " + LTrim$(Str$(RVOL))

End Sub


Public Function lSetVolume(ByRef lLeftVol As Long, ByRef lrightVol As Long, lDeviceID As Long) As Long

    Dim bReturnValue As Boolean                     ' Return Value from Function
    Dim Volume As VolumeSetting                     ' Type structure used to convert a long to/from
                                                    ' two Integers.
    Dim lAPIReturnVal As Long                       ' Return value from API Call
    Dim lBothVolumes As Long                        ' The API passed value of the Combined Volumes
    
    Volume.LeftVol = nSigned(lLeftVol * 65535 / HIGHEST_VOLUME_SETTING)
    Volume.rightVol = nSigned(lrightVol * 65535 / HIGHEST_VOLUME_SETTING)
    
    lDataLen = Len(Volume)
    CopyMemory lBothVolumes, Volume.LeftVol, lDataLen

    lAPIReturnVal = auxSetVolume(lDeviceID, lBothVolumes)
    lSetVolume = lAPIReturnVal

End Function


Public Function lGetVolume(ByRef lLeftVol As Long, ByRef lrightVol As Long, lDeviceID As Long) As Long

    Dim bReturnValue As Boolean                     ' Return Value from Function
    Dim Volume As VolumeSetting                     ' Type structure used to convert a long to/from
                                                    ' two Integers.
    Dim lAPIReturnVal As Long                       ' Return value from API Call
    Dim lBothVolumes As Long                        ' The API Return of the Combined Volumes
    lAPIReturnVal = auxGetVolume(lDeviceID, lBothVolumes)
    lDataLen = Len(Volume)
    CopyMemory Volume.LeftVol, lBothVolumes, lDataLen
    lLeftVol = HIGHEST_VOLUME_SETTING * lUnsigned(Volume.LeftVol) / 65535
    lrightVol = HIGHEST_VOLUME_SETTING * lUnsigned(Volume.rightVol) / 65535
    lGetVolume = lAPIReturnVal
End Function

Public Function nSigned(ByVal lUnsignedInt As Long) As Integer
    Dim nReturnVal As Integer                          ' Return value from Function
    
    If lUnsignedInt > 65535 Or lUnsignedInt < 0 Then
        MsgBox "Error in conversion from Unsigned to nSigned Integer"
        nSignedInt = 0
        Exit Function
    End If

    If lUnsignedInt > 32767 Then
        nReturnVal = lUnsignedInt - 65536
    Else
        nReturnVal = lUnsignedInt
    End If
    
    nSigned = nReturnVal

End Function

Public Function lUnsigned(ByVal nSignedInt As Integer) As Long
    Dim lReturnVal As Long                          ' Return value from Function
    
    If nSignedInt < 0 Then
        lReturnVal = nSignedInt + 65536
    Else
        lReturnVal = nSignedInt
    End If
    
    If lReturnVal > 65535 Or lReturnVal < 0 Then
        MsgBox "Error in conversion from nSigned to Unsigned Integer"
        lReturnVal = 0
    End If
    
    lUnsigned = lReturnVal
End Function


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
   rc = mmioReadFormat(hmmioIn, formatA, Len(formatA))
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
   
   numSamples = mmckinfoSubchunkIn.ckSize / formatA.nBlockAlign
   
   ' Close file
   rc = mmioClose(hmmioOut, 0)
   
   fFileLoaded = True
    
End Sub

Sub play(ByVal soundcard As Integer)
' Send audio buffer to wave output

    rc = waveOutOpen(hWaveOut, soundcard, formatA, AddressOf waveOutProc, 0, CALLBACK_FUNCTION)
    If (rc <> 0) Then
      GlobalFree (hmem)
      waveOutGetErrorText rc, msg, Len(msg)
      MsgBox msg
      Exit Sub
    End If

    outHdr.lpData = bufferIn + (drawFrom * formatA.nBlockAlign)
    outHdr.dwBufferLength = (drawTo - drawFrom) * formatA.nBlockAlign
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
      Form6.Timer1.Enabled = True
    End If
End Sub

Sub StopPlay()
   waveOutReset (hWaveOut)
End Sub


Sub GetStereo16Sample(ByVal sample As Long, ByRef LeftVol As Double, ByRef rightVol As Double)
' These subs obtain a PCM sample and converts it into volume levels from (-1 to 1)

   Dim sample16 As Integer
   Dim ptr As Long
   
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample16, ptr, 2
   LeftVol = sample16 / 32768
   CopyStructFromPtr sample16, ptr + 2, 2
   rightVol = sample16 / 32768

End Sub

Sub GetStereo8Sample(ByVal sample As Long, ByRef LeftVol As Double, ByRef rightVol As Double)

   Dim sample8 As Byte
   Dim ptr As Long
   
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample8, ptr, 1
   LeftVol = (sample8 - 128) / 128
   CopyStructFromPtr sample8, ptr + 1, 1
   rightVol = (sample8 - 128) / 128

End Sub

Sub GetMono16Sample(ByVal sample As Long, ByRef LeftVol As Double)

   Dim sample16 As Integer
   Dim ptr As Long
   
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample16, ptr, 2
   LeftVol = sample16 / 32768

End Sub

Sub GetMono8Sample(ByVal sample As Long, ByRef LeftVol As Double)

   Dim sample8 As Byte
   Dim ptr As Long
   
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample8, ptr, 1
   LeftVol = (sample8 - 128) / 128

End Sub

