Attribute VB_Name = "WAVEMIX32"
Option Explicit

''Declarations for the WaveMix32.dll library
Declare Function WaveMixInit Lib "WAVMIX32.DLL" () As Long
Declare Function WaveMixConfigureInit Lib "WAVMIX32.DLL" (NewConfig As MIXCONFIG) As Long
Declare Function WaveMixActivate Lib "WAVMIX32.DLL" (ByVal MixSession As Long, ByVal Activate As Integer) As Long

Declare Function WaveMixOpenWave Lib "WAVMIX32.DLL" (ByVal MixSession As Long, WaveFilename As Any, ByVal hInst As Long, ByVal Flags As Long) As Long

Declare Function WaveMixOpenChannel Lib "WAVMIX32.DLL" (ByVal MixSession As Long, ByVal Channel As Long, ByVal Flags As Long) As Long
Declare Function WaveMixPlay Lib "WAVMIX32.DLL" (PlayParameters As MIXPLAYPARAMS) As Integer
Declare Function WaveMixFlushChannel Lib "WAVMIX32.DLL" (ByVal MixSession As Long, ByVal Channel As Long, ByVal Flags As Long) As Integer
Declare Function WaveMixCloseChannel Lib "WAVMIX32.DLL" (ByVal MixSession As Long, ByVal Channel As Long, ByVal Flags As Long) As Integer
Declare Function WaveMixFreeWave Lib "WAVMIX32.DLL" (ByVal MixSession As Long, ByVal MixWave As Long) As Integer
Declare Function WaveMixCloseSession Lib "WAVMIX32.DLL" (ByVal MixSession As Long) As Integer
Declare Sub WaveMixPump Lib "WAVMIX32.DLL" ()
Declare Function WaveMixGetInfo Lib "WAVMIX32.DLL" (WaveMixInfo As WaveMixInfo) As Integer

'Types for the WaveMix32.dll library
Public Type WaveMixInfo
   Size As Integer
   VersionMajor As Byte
   VersionMinor As Byte
   Date(12) As String
   Formats As Long
End Type

Public Type MIXCONFIG
    Size As Integer
    Flags As Long
    Channels As Integer
    SamplingRate As Integer
End Type

Public Type MIXPLAYPARAMS
    Size         As Integer
    MixSessionLo As Integer
    MixSessionHi As Integer
    ChannelLo    As Integer
    ChannelHi    As Integer
    MixWaveLo   As Integer
    MixWaveHi   As Integer
    hWndNotifyLo  As Integer
    hWndNotifyHi  As Integer
    FlagsLo     As Integer
    FlagsHi     As Integer
    wLoops        As Integer
End Type

'Constants for the WaveMix32.dll library
Public Const WMIX_QUEUEWAVE As Long = &H0
Public Const WMIX_CLEARQUEUE As Long = &H1
Public Const WMIX_USELRUCHANNEL As Long = &H2
Public Const WMIX_HIGHPRIORITY As Long = &H4
Public Const WMIX_WAIT As Long = &H8

Public Const WMIX_CONFIG_CHANNELS As Long = &H1
Public Const WMIX_CONFIG_SAMPLINGRATE As Long = &H2

Public Const WMIX_FILE As Long = &H0
Public Const WMIX_RESOURCE As Long = &H2

Public Const WMIX_OPENSINGLE As Long = &H0
Public Const WMIX_OPENALL As Long = &H1
Public Const WMIX_OPENCOUNT As Long = &H2

Public Const WMIX_ALL As Long = &H1
Public Const WMIX_NOREMIX As Long = &H2

Public Const WMIX_ACTIVATE As Integer = 1
Public Const WMIX_DEACTIVATE As Integer = 0
''''''''''''''''''''''''''''''''''''''''''''''

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
