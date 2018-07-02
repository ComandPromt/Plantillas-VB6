Attribute VB_Name = "WAVEMIX"
'
'Declaration for the WaveMix32.dll file
'

Option Explicit


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
Declare Function WaveMixGetInfo Lib "WAVMIX32.DLL" (lpWaveMixInfo As WAVEMIXINFO) As Integer
Declare Sub WaveMixPump Lib "WAVMIX32.DLL" ()

' Flag values for MIXPLAYPARAMS
Public Const WMIX_QUEUEWAVE = &H0
Public Const WMIX_CLEARQUEUE = &H1
Public Const WMIX_USELRUCHANNEL = &H2
Public Const WMIX_HIGHPRIORITY = &H4
Public Const WMIX_WAIT = &H8

Type MIXPLAYPARAMS
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

'Flags for MIXCONFIG
Public Const WMIX_CONFIG_CHANNELS = &H1
Public Const WMIX_CONFIG_SAMPLINGRATE = &H2

Type MIXCONFIG
    wSize As Integer
    dwFlags As Long
    wChannels As Integer
    wSamplingRate As Integer
End Type


Public Const WMIX_FILE = &H1
Public Const WMIX_RESOURCE = &H2
Public Const WMIX_MEMORY = &H4

Public Const WMIX_OPENSINGLE = 0
Public Const WMIX_OPENALL = 1
Public Const WMIX_OPENCOUNT = 2

Public Const WMIX_ALL = &H1
Public Const WMIX_NOREMIX = &H2


Public Const WAVERR_BASE = 32

Public Const MMSYSERR_INVALHANDLE = 5
Public Const MMSYSERR_BADDEVICEID = 2
Public Const MMSYSERR_ALLOCATED = 4
Public Const MMSYSERR_NOMEM = 7
Public Const MMSYSERR_ERROR = 1
Public Const WAVERR_BADFORMAT = 32

'Returns the higher Word value of the passed value
Function HighWord(ByVal Value As Long) As Integer
    Value = Value \ &H10000
    HighWord = Val("&H" & Hex$(Value))
End Function

'Returns the Lower Word of the passed value
Function LowerWord(ByVal Value As Long) As Integer
    Value = Value And &HFFFF&
    LowerWord = Val("&H" & Hex$(Value))
End Function

