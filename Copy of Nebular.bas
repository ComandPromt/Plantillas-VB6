Attribute VB_Name = "Nebular1"
'This is an update from my cd/mixer I have taken the cd player away
'and put the slider/mutes/balance controls into arrays.
'this makes it alittle more harder to understand but it is better/faster coding.
'If you wish to comment or get the VU meters working for me email me at
'micracom2@hotmail.com and I shall see if I can help.

Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2& ' separate left-right volume control
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Public Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Public Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Public Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Public Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Public Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Public Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Public Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000
Public Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000
Public Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Public Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG = &H1000& + 10
Public Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = &H1000& + 9
Public Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = &H1000& + 5
Public Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = &H1000& + 1
Public Const MIXERLINE_COMPONENTTYPE_SRC_LAST = &H1000& + 10
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = &H1000& + 2
Public Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = &H1000& + 4
Public Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = &H1000& + 0
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = &H1000& + 8

Public Const MIXERLINE_COMPONENTTYPE_SRC_I25InVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINEVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MIDIVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_CDVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_src_AUXVol = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)

Public Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 0)
Public Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_DST_LINE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_DST_LAST = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)

Public Const MMIO_READ = &H0
Public Const MMIO_FINDCHUNK = &H10
Public Const MMIO_FINDRIFF = &H20
' Mixer control types

Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Public Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Public Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_BUTTON = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
Public Const MIXERCONTROL_CONTROLTYPE_DECIBELS = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
Public Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MICROTIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MILLITIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Public Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_STEREOENH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_PERCENT = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
Public Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)
' Mixer line flags
Public Const MIXERLINE_LINEF_ACTIVE = &H1&
Public Const MIXERLINE_LINEF_DISCONNECTED = &H8000&
Public Const MIXERLINE_LINEF_SOURCE = &H80000000
' Mixer line target types
Public Const MIXERLINE_TARGETTYPE_AUX = 5
Public Const MIXERLINE_TARGETTYPE_MIDIIN = 4
Public Const MIXERLINE_TARGETTYPE_MIDIOUT = 3
Public Const MIXERLINE_TARGETTYPE_UNDEFINED = 0
Public Const MIXERLINE_TARGETTYPE_WAVEIN = 2
Public Const MIXERLINE_TARGETTYPE_WAVEOUT = 1
' Error constants
Public Const MIXERR_INVALLINE = 1024 + 0
Public Const MIXERR_BASE = 1024
Public Const MIXERR_INVALCONTROL = 1024 + 1
Public Const MIXERR_INVALVALUE = 1024 + 2
Public Const MIXERR_LASTERROR = 1024 + 2
Type MIXERCAPS
     wMid As Integer                   '  manufacturer id
     wPid As Integer                   '  product id
     vDriverVersion As Long            '  version of the driver
     szPname As String * MAXPNAMELEN   '  product name
     fdwSupport As Long             '  misc. support bits
     cDestinations As Long          '  count of destinations
End Type

' Mixer line types
Type Target
     dwType As Long                 '  MIXERLINE_TARGETTYPE_xxxx
     dwDeviceID As Long             '  target device ID of device type
     wMid As Integer                   '  of target device
     wPid As Integer                   '       "
     vDriverVersion As Long            '       "
     szPname As String * MAXPNAMELEN
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
     lpTarget As Target
End Type

' MM Control types
Type MIXERLINECONTROLS
     cbStruct As Long         '  size in Byte of MIXERLINECONTROLS
     dwLineID As Long         '  line id (from MIXERLINE.dwLineID)
     dwControl As Long        '  used with MIXER_GETLINECONTROLSF_ONEBYTYPE or MIXER_GETLINECONTROLSF_ONEBYID
     cControls As Long        '  count of controls pmxctrl points to
     cbmxctrl As Long         '  size in Byte of _one_ MIXERCONTROL
     pamxctrl As Long         '  pointer to first MIXERCONTROL array
End Type

Type MIXERCONTROL
     cbStruct As Long           '  size in Byte of MIXERCONTROL
     dwControlID As Long        '  unique control id for mixer device
     dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
     fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
     cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
     szShortName(1 To MIXER_SHORT_NAME_CHARS) As Byte
     szName(1 To MIXER_LONG_NAME_CHARS) As Byte
     Bounds(1 To 6) As Long
     Metrics(1 To 6) As Long
End Type

Type MIXERCONTROLDETAILS
     cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
     dwControlID As Long    '  control id to get/set details on
     cChannels As Long      '  number of channels in paDetails array
     item As Long                           ' hwndOwner or cMultipleItems
     cbDetails As Long      '  size of _one_ details_XX struct
     paDetails As Long      '  pointer to array of details_XX structs
End Type

Type MIXERCONTROLDETAILS_BOOLEAN
     fValue As Long
End Type

Type MIXERCONTROLDETAILS_LISTTEXT
     dwParam1 As Long
     dwParam2 As Long
     szName As String * MIXER_LONG_NAME_CHARS
End Type

Type MIXERCONTROLDETAILS_SIGNED
     lValue As Long
End Type

Type MIXERCONTROLDETAILS_UNSIGNED
     dwValue As Long
End Type

Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, pumxID As Long, ByVal fdwId As Long) As Long
Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, ByVal uMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Public Declare Function RegisterDLL Lib "Regist10.dll" Alias "REGISTERDLL" (ByVal DllPath As String, bRegister As Boolean) As Boolean
Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal X As Long, ByVal uFlags As Long) As Long
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Declare Function mmioReadFormat Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByRef pch As waveFormat, ByVal cch As Long) As Long
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)

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

Dim rc As Long

' variables for managing wave file
Public formatA As waveFormat
Dim hmmioOut As Long
Dim mmckinfoParentIn As MMCKINFO
Dim mmckinfoSubchunkIn As MMCKINFO
Dim bufferIn As Long
Dim hmem As Long
Public numSamples As Long
Public drawFrom As Long
Public drawTo As Long
Public fFileLoaded As Boolean

Type waveFormat
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
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

Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer ' e.g. = &h0031 = .31
    dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
    dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
    dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
    dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
    dwFileFlagsMask As Long ' = &h3F For version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
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
       rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
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

Function SetVolumeControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volume As Long) As Boolean
  Dim mxcd As MIXERCONTROLDETAILS
   Dim vol As MIXERCONTROLDETAILS_UNSIGNED
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.item = 0
   mxcd.cbDetails = Len(vol)
   hmem = GlobalAlloc(&H40, Len(vol))
   mxcd.paDetails = GlobalLock(hmem)
   vol.dwValue = volume
   ' Copy the data into the control value buffer
   CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetVolumeControl = True
   Else
       SetVolumeControl = False
   End If
End Function

Function SetPANControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volL As Long, ByVal volR As Long) As Boolean
'This function sets the value for a volume control. Returns True if successful
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol(1) As MIXERCONTROLDETAILS_UNSIGNED
   mxcd.item = mxc.cMultipleItems
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cbStruct = Len(mxcd)
   mxcd.cbDetails = Len(vol(1))
   ' Allocate a buffer for the control value buffer
   mxcd.cChannels = 2
   hmem = GlobalAlloc(&H40, Len(vol(1)))
   mxcd.paDetails = GlobalLock(hmem)
   vol(1).dwValue = volR
   vol(0).dwValue = volL
   ' Copy the data into the control value buffer
   CopyPtrFromStruct mxcd.paDetails, vol(1).dwValue, Len(vol(0)) * mxcd.cChannels
   CopyPtrFromStruct mxcd.paDetails, vol(0).dwValue, Len(vol(1)) * mxcd.cChannels
   ' Set the control value
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetPANControl = True
   Else
       SetPANControl = False
   End If
   
End Function

Function unSetMuteControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal unmute As Long) As Boolean
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol As MIXERCONTROLDETAILS_UNSIGNED
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.item = 0
   mxcd.cbDetails = Len(vol)
   hmem = GlobalAlloc(&H40, Len(vol))
   mxcd.paDetails = GlobalLock(hmem)
   vol.dwValue = unmute
   CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       unSetMuteControl = True
   Else
       unSetMuteControl = False
   End If
End Function

Function SetMuteControl(ByVal hmixer As Long, mxc As MIXERCONTROL, mute As Boolean) As Boolean
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol As MIXERCONTROLDETAILS_UNSIGNED
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.item = 0
   mxcd.cbDetails = Len(vol)
   hmem = GlobalAlloc(&H40, Len(vol))
   mxcd.paDetails = GlobalLock(hmem)
   vol.dwValue = volume
   CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetMuteControl = True
   Else
       SetMuteControl = False
   End If
End Function

Function GetVolumeControlValue(ByVal hmixer As Long, mxc As MIXERCONTROL) As Long
'This function Gets the value for a volume control. Returns True if successful
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    mxcd.cbStruct = Len(mxcd)
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cChannels = 1
    mxcd.item = 0
    mxcd.cbDetails = Len(vol)
    mxcd.paDetails = 0
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr vol, mxcd.paDetails, Len(vol)
    GlobalFree (hmem)
    If (MMSYSERR_NOERROR = rc) Then
       GetVolumeControlValue = vol.dwValue
    Else
        GetVolumeControlValue = -1
    End If
End Function

