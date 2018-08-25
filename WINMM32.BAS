Attribute VB_Name = "basWinMM"
' Module Name               basWinMM
' Module Description        Module which contains ALL Declarations
'                           for API calls in the WinMM Library
'
Option Explicit

Declare Function auxGetDevCaps _
    Lib "winmm.dll" _
    Alias "auxGetDevCapsA" ( _
        ByVal uDeviceID As Long, _
        lpCaps As AUXCAPS, _
        ByVal uSize As Long) _
    As Long

Declare Function auxGetNumDevs _
    Lib "winmm.dll" () _
    As Long

Declare Function auxGetVolume _
    Lib "winmm.dll" ( _
        ByVal uDeviceID As Long, _
        lpdwVolume As Long) _
    As Long

Declare Function auxOutMessage _
    Lib "winmm.dll" ( _
        ByVal uDeviceID As Long, _
        ByVal msg As Long, _
        ByVal dw1 As Long, _
        ByVal dw2 As Long) _
    As Long

Declare Function auxSetVolume _
    Lib "winmm.dll" ( _
        ByVal uDeviceID As Long, _
        ByVal dwVolume As Long) _
    As Long

Declare Function CloseDriver _
    Lib "winmm.dll" ( _
        ByVal hDriver As Long, _
        ByVal lParam1 As Long, _
        ByVal lParam2 As Long) _
    As Long

Declare Function DefDriverProc _
    Lib "winmm.dll" ( _
        ByVal dwDriverIdentifier As Long, _
        ByVal hdrvr As Long, _
        ByVal uMsg As Long, _
        ByVal lParam1 As Long, _
        ByVal lParam2 As Long) _
    As Long

Declare Function DrvGetModuleHandle _
    Lib "winmm.dll" ( _
        ByVal hDriver As Long) _
    As Long

Declare Function GetDriverModuleHandle _
    Lib "winmm.dll" ( _
        ByVal hDriver As Long) _
    As Long

Declare Function joyGetDevCaps _
    Lib "winmm.dll" _
    Alias "joyGetDevCapsA" ( _
        ByVal id As Long, _
        lpCaps As JOYCAPS, _
        ByVal uSize As Long) _
    As Long

Declare Function joyGetNumDevs _
    Lib "winmm.dll" _
    Alias "joyGetNumDev" ( _
        ) _
    As Long

Declare Function joyGetPos _
    Lib "winmm.dll" ( _
        ByVal uJoyID As Long, _
        pji As JOYINFO) _
    As Long

Declare Function joyGetPosEx _
    Lib "winmm.dll" ( _
        ByVal uJoyID As Long, _
        pji As JOYINFOEX) _
    As Long

Declare Function joyGetThreshold _
    Lib "winmm.dll" ( _
        ByVal id As Long, _
        lpuThreshold As Long) _
    As Long

Declare Function joyReleaseCapture _
    Lib "winmm.dll" ( _
        ByVal id As Long) _
    As Long

Declare Function joySetCapture _
    Lib "winmm.dll" ( _
        ByVal hWnd As Long, _
        ByVal uID As Long, _
        ByVal uPeriod As Long, _
        ByVal bChanged As Long) _
    As Long

Declare Function joySetThreshold _
    Lib "winmm.dll" ( _
        ByVal id As Long, _
        ByVal uThreshold As Long) _
    As Long

Declare Function mciExecute _
    Lib "winmm.dll" ( _
        ByVal lpstrCommand As String) _
    As Long

Declare Function mciGetCreatorTask _
    Lib "winmm.dll" ( _
        ByVal wDeviceID As Long) _
    As Long

Declare Function mciGetDeviceID _
    Lib "winmm.dll" _
    Alias "mciGetDeviceIDA" ( _
        ByVal lpstrName As String) _
    As Long

Declare Function mciGetDeviceIDFromElementID _
    Lib "winmm.dll" _
    Alias "mciGetDeviceIDFromElementIDA" ( _
        ByVal dwElementID As Long, _
        ByVal lpstrType As String) _
    As Long

Declare Function mciGetErrorString _
    Lib "winmm.dll" _
    Alias "mciGetErrorStringA" ( _
        ByVal dwError As Long, _
        ByVal lpstrBuffer As String, _
        ByVal uLength As Long) _
    As Long

Declare Function mciGetYieldProc _
    Lib "winmm" ( _
        ByVal mciId As Long, _
        pdwYieldData As Long) _
    As Long

Declare Function mciSendCommand _
    Lib "winmm.dll" _
    Alias "mciSendCommandA" ( _
        ByVal wDeviceID As Long, _
        ByVal uMessage As Long, _
        ByVal dwParam1 As Long, _
        ByVal dwParam2 As Long) _
    As Long

Declare Function mciSendString _
    Lib "winmm.dll" _
    Alias "mciSendStringA" ( _
        ByVal lpstrCommand As String, _
        ByVal lpstrReturnString As String, _
        ByVal uReturnLength As Long, _
        ByVal hwndCallback As Long) _
    As Long

Declare Function mciSetYieldProc _
    Lib "winmm" ( _
        ByVal mciId As Long, _
        ByVal fpYieldProc As Long, _
        ByVal dwYieldData As Long) _
    As Boolean

Declare Function midiConnect _
    Lib "winmm.dll" ( _
        ByVal hmi As Long, _
        ByVal hmo As Long, _
        pReserved As Any) _
    As Long

Declare Function midiDisconnect _
    Lib "winmm.dll" ( _
        ByVal hmi As Long, _
        ByVal hmo As Long, _
        pReserved As Any) _
    As Long

Declare Function midiInAddBuffer _
    Lib "winmm.dll" ( _
        ByVal hMidiIn As Long, _
        lpMidiInHdr As MIDIHDR, _
        ByVal uSize As Long) _
    As Long

Declare Function midiInClose _
    Lib "winmm.dll" ( _
        ByVal hMidiIn As Long) _
    As Long

Declare Function midiInGetDevCaps _
    Lib "winmm.dll" _
    Alias "midiInGetDevCapsA" ( _
        ByVal uDeviceID As Long, _
        lpCaps As MIDIINCAPS, _
        ByVal uSize As Long) _
    As Long

Declare Function midiInGetErrorText _
    Lib "winmm.dll" _
    Alias "midiInGetErrorTextA" ( _
        ByVal err As Long, _
        ByVal lpText As String, _
        ByVal uSize As Long) _
    As Long

Declare Function midiInGetID _
    Lib "winmm.dll" ( _
        ByVal hMidiIn As Long, _
        lpuDeviceID As Long) _
    As Long

Declare Function midiInGetNumDevs _
    Lib "winmm.dll" ( _
        ) _
    As Long

Declare Function midiInMessage _
    Lib "winmm.dll" ( _
        ByVal hMidiIn As Long, _
        ByVal msg As Long, _
        ByVal dw1 As Long, _
        ByVal dw2 As Long) _
    As Long

Declare Function midiInOpen _
    Lib "winmm.dll" ( _
        lphMidiIn As Long, _
        ByVal uDeviceID As Long, _
        ByVal dwCallback As Long, _
        ByVal dwInstance As Long, _
        ByVal dwFlags As Long) _
    As Long

Declare Function midiInPrepareHeader _
    Lib "winmm.dll" ( _
        ByVal hMidiIn As Long, _
        lpMidiInHdr As MIDIHDR, _
        ByVal uSize As Long) _
    As Long

Declare Function midiInReset _
    Lib "winmm.dll" ( _
        ByVal hMidiIn As Long) _
    As Long

Declare Function midiInStart _
    Lib "winmm.dll" ( _
        ByVal hMidiIn As Long) _
    As Long

Declare Function midiInStop _
    Lib "winmm.dll" ( _
        ByVal hMidiIn As Long) _
    As Long

Declare Function midiInUnprepareHeader _
    Lib "winmm.dll" ( _
        ByVal hMidiIn As Long, _
        lpMidiInHdr As MIDIHDR, _
        ByVal uSize As Long) _
    As Long

Declare Function midiOutCacheDrumPatches _
    Lib "winmm.dll" ( _
        ByVal hMidiOut As Long, _
        ByVal uPatch As Long, _
        lpKeyArray As Long, _
        ByVal uFlags As Long) _
    As Long

Declare Function midiOutCachePatches _
    Lib "winmm.dll" ( _
        ByVal hMidiOut As Long, _
        ByVal uBank As Long, _
        lpPatchArray As Long, _
        ByVal uFlags As Long) _
    As Long

Declare Function midiOutClose _
    Lib "winmm.dll" ( _
        ByVal hMidiOut As Long) _
    As Long

Declare Function midiOutGetDevCaps _
    Lib "winmm.dll" _
    Alias "midiOutGetDevCapsA" ( _
        ByVal uDeviceID As Long, _
        lpCaps As MIDIOUTCAPS, _
        ByVal uSize As Long) _
    As Long

Declare Function midiOutGetErrorText _
    Lib "winmm.dll" _
    Alias "midiOutGetErrorTextA" ( _
        ByVal err As Long, _
        ByVal lpText As String, _
        ByVal uSize As Long) _
    As Long

Declare Function midiOutGetID _
    Lib "winmm.dll" ( _
        ByVal hMidiOut As Long, _
        lpuDeviceID As Long) _
    As Long

Declare Function midiOutGetNumDevs _
    Lib "winmm" () _
    As Integer

Declare Function midiOutGetVolume _
    Lib "winmm.dll" ( _
        ByVal uDeviceID As Long, _
        lpdwVolume As Long) _
    As Long

Declare Function midiOutLongMsg _
    Lib "winmm.dll" ( _
        ByVal hMidiOut As Long, _
        lpMidiOutHdr As MIDIHDR, _
        ByVal uSize As Long) _
    As Long

Declare Function midiOutMessage _
    Lib "winmm.dll" ( _
        ByVal hMidiOut As Long, _
        ByVal msg As Long, _
        ByVal dw1 As Long, _
        ByVal dw2 As Long) _
    As Long

Declare Function midiOutOpen _
    Lib "winmm.dll" ( _
        lphMidiOut As Long, _
        ByVal uDeviceID As Long, _
        ByVal dwCallback As Long, _
        ByVal dwInstance As Long, _
        ByVal dwFlags As Long) _
    As Long

Declare Function midiOutPrepareHeader _
    Lib "winmm.dll" ( _
        ByVal hMidiOut As Long, _
        lpMidiOutHdr As MIDIHDR, _
        ByVal uSize As Long) _
    As Long

Declare Function midiOutReset _
    Lib "winmm.dll" ( _
        ByVal hMidiOut As Long) _
    As Long

Declare Function midiOutSetVolume _
    Lib "winmm.dll" ( _
        ByVal uDeviceID As Long, _
        ByVal dwVolume As Long) _
    As Long

Declare Function midiOutShortMsg _
    Lib "winmm.dll" ( _
        ByVal hMidiOut As Long, _
        ByVal dwMsg As Long) _
    As Long

Declare Function midiOutUnprepareHeader _
    Lib "winmm.dll" ( _
        ByVal hMidiOut As Long, _
        lpMidiOutHdr As MIDIHDR, _
        ByVal uSize As Long) _
    As Long

Declare Function midiStreamClose _
    Lib "winmm.dll" ( _
        ByVal hms As Long) _
    As Long

Declare Function midiStreamOpen _
    Lib "winmm.dll" ( _
        phms As Long, puDeviceID As Long, _
        ByVal cMidi As Long, ByVal dwCallback As Long, _
        ByVal dwInstance As Long, ByVal fdwOpen As Long) _
    As Long

Declare Function midiStreamOut _
    Lib "winmm.dll" ( _
        ByVal hms As Long, _
        pmh As MIDIHDR, _
        ByVal cbmh As Long) _
    As Long

Declare Function midiStreamPause _
    Lib "winmm.dll" ( _
        ByVal hms As Long) _
    As Long

Declare Function midiStreamPosition _
    Lib "winmm.dll" ( _
        ByVal hms As Long, _
        lpmmt As MMTIME, _
        ByVal cbmmt As Long) _
    As Long

Declare Function midiStreamProperty _
    Lib "winmm.dll" ( _
        ByVal hms As Long, _
        lppropdata As Byte, _
        ByVal dwProperty As Long) _
    As Long

Declare Function midiStreamRestart _
    Lib "winmm.dll" ( _
        ByVal hms As Long) _
    As Long

Declare Function midiStreamStop _
    Lib "winmm.dll" ( _
        ByVal hms As Long) _
    As Long

Declare Function mixerClose _
    Lib "winmm.dll" ( _
        ByVal hmx As Long) _
    As Long

Declare Function mixerGetControlDetails _
    Lib "winmm.dll" _
    Alias "mixerGetControlDetailsA" ( _
        ByVal hmxobj As Long, _
        pmxcd As MIXERCONTROLDETAILS, _
        ByVal fdwDetails As Long) _
    As Long

Declare Function mixerGetDevCaps _
    Lib "winmm.dll" _
    Alias "mixerGetDevCapsA" ( _
        ByVal uMxId As Long, _
        ByVal pmxcaps As MIXERCAPS, _
        ByVal cbmxcaps As Long) _
    As Long

Declare Function mixerGetID _
    Lib "winmm.dll" ( _
        ByVal hmxobj As Long, _
        pumxID As Long, _
        ByVal fdwId As Long) _
    As Long

Declare Function mixerGetLineControls _
    Lib "winmm.dll" _
    Alias "mixerGetLineControlsA" ( _
        ByVal hmxobj As Long, _
        pmxlc As MIXERLINECONTROLS, _
        ByVal fdwControls As Long) _
    As Long

Declare Function mixerGetLineInfo _
    Lib "winmm.dll" _
    Alias "mixerGetLineInfoA" ( _
        ByVal hmxobj As Long, _
        pmxl As MIXERLINE, _
        ByVal fdwInfo As Long) _
    As Long

Declare Function mixerGetNumDevs _
    Lib "winmm.dll" () _
    As Long

Declare Function mmioOpen _
    Lib "winmm.dll" _
    Alias "mmioOpenA" ( _
        ByVal szFileName As String, _
        lpmmioinfo As MMIOINFO, _
        ByVal dwOpenFlags As Long) _
    As Long

Declare Function mmioRead _
    Lib "winmm.dll" ( _
        ByVal hmmio As Long, _
        ByVal pch As String, _
        ByVal cch As Long) _
    As Long

Declare Function mmioRename _
    Lib "winmm.dll" _
    Alias "mmioRenameA" ( _
        ByVal szFileName As String, _
        ByVal SzNewFileName As String, _
        lpmmioinfo As MMIOINFO, _
        ByVal dwRenameFlags As Long) _
    As Long

Declare Function mmioSeek _
    Lib "winmm.dll" ( _
        ByVal hmmio As Long, _
        ByVal lOffset As Long, _
        ByVal iOrigin As Long) _
    As Long

Declare Function mmioSendMessage _
    Lib "winmm.dll" ( _
        ByVal hmmio As Long, _
        ByVal uMsg As Long, _
        ByVal lParam1 As Long, _
        ByVal lParam2 As Long) _
    As Long

Declare Function mmioSetBuffer _
    Lib "winmm.dll" ( _
        ByVal hmmio As Long, _
        ByVal pchBuffer As String, _
        ByVal cchBuffer As Long, _
        ByVal uFlags As Long) _
    As Long

Declare Function mmioSetInfo _
    Lib "winmm.dll" ( _
        ByVal hmmio As Long, _
        lpmmioinfo As MMIOINFO, _
        ByVal uFlags As Long) _
    As Long

Declare Function mmioStringToFOURCC _
    Lib "winmm.dll" _
    Alias "mmioStringToFOURCCA" ( _
        ByVal sz As String, _
        ByVal uFlags As Long) _
    As Long

Declare Function mmioWrite _
    Lib "winmm.dll" ( _
        ByVal hmmio As Long, _
        ByVal pch As String, _
        ByVal cch As Long) _
    As Long

Declare Function mmsystemGetVersion _
    Lib "winmm.dll" () _
    As Long

Declare Function OpenDriver _
    Lib "winmm.dll" ( _
        ByVal szDriverName As String, _
        ByVal szSectionName As String, _
        ByVal lParam2 As Long) _
    As Long

Declare Sub OutputDebugStr _
    Lib "winmm.dll" ( _
        ByVal lpszOutputString As String)

Declare Function PlaySound _
    Lib "winmm.dll" _
    Alias "PlaySoundA" ( _
        ByVal lpszName As String, _
        ByVal hModule As Long, _
        ByVal dwFlags As Long) _
    As Long

Declare Function SendDriverMessage _
    Lib "winmm.dll" ( _
        ByVal hDriver As Long, _
        ByVal message As Long, _
        ByVal lParam1 As Long, _
        ByVal lParam2 As Long) _
    As Long

Declare Function sndPlaySound _
    Lib "winmm.dll" _
    Alias "sndPlaySoundA" ( _
        ByVal lpszSoundName As String, _
        ByVal uFlags As Long) _
    As Long

Declare Function timeBeginPeriod _
    Lib "winmm.dll" ( _
        ByVal uPeriod As Long) _
    As Long

Declare Function timeEndPeriod _
    Lib "winmm.dll" ( _
        ByVal uPeriod As Long) _
    As Long

Declare Function timeGetDevCaps _
    Lib "winmm.dll" ( _
        lpTimeCaps As TIMECAPS, _
        ByVal uSize As Long) _
    As Long

Declare Function timeGetSystemTime _
    Lib "winmm.dll" ( _
        lpTime As MMTIME, _
        ByVal uSize As Long) _
    As Long

Declare Function timeGetTime _
    Lib "winmm.dll" () _
    As Long

Declare Function timeKillEvent _
    Lib "winmm.dll" ( _
        ByVal uID As Long) _
    As Long

Declare Function timeSetEvent _
    Lib "winmm.dll" ( _
        ByVal uDelay As Long, _
        ByVal uResolution As Long, _
        ByVal lpFunction As Long, _
        ByVal dwUser As Long, _
        ByVal uFlags As Long) _
    As Long

Declare Function waveInAddBuffer _
    Lib "winmm.dll" ( _
        ByVal hWaveIn As Long, _
        lpWaveInHdr As WAVEHDR, _
        ByVal uSize As Long) _
    As Long

Declare Function waveInClose _
    Lib "winmm.dll" ( _
        ByVal hWaveIn As Long) _
    As Long

Declare Function waveInGetDevCaps _
    Lib "winmm.dll" _
    Alias "waveInGetDevCapsA" ( _
        ByVal uDeviceID As Long, _
        lpCaps As WAVEINCAPS, _
        ByVal uSize As Long) _
    As Long

Declare Function waveInGetErrorText _
    Lib "winmm.dll" _
    Alias "waveInGetErrorTextA" ( _
        ByVal err As Long, _
        ByVal lpText As String, _
        ByVal uSize As Long) _
    As Long

Declare Function waveInGetID _
    Lib "winmm.dll" ( _
        ByVal hWaveIn As Long, _
        lpuDeviceID As Long) _
    As Long

Declare Function waveInGetNumDevs _
    Lib "winmm.dll" () _
    As Long

Declare Function waveInGetPosition _
    Lib "winmm.dll" ( _
        ByVal hWaveIn As Long, _
        lpInfo As MMTIME, _
        ByVal uSize As Long) _
    As Long

Declare Function waveInMessage _
    Lib "winmm.dll" ( _
        ByVal hWaveIn As Long, _
        ByVal msg As Long, _
        ByVal dw1 As Long, _
        ByVal dw2 As Long) _
    As Long

Declare Function waveInOpen _
    Lib "winmm.dll" ( _
        lphWaveIn As Long, ByVal uDeviceID As Long, _
        lpFormat As WAVEFORMAT, ByVal dwCallback As Long, _
        ByVal dwInstance As Long, ByVal dwFlags As Long) _
    As Long

Declare Function waveInPrepareHeader _
    Lib "winmm.dll" ( _
        ByVal hWaveIn As Long, _
        lpWaveInHdr As WAVEHDR, _
        ByVal uSize As Long) _
    As Long

Declare Function waveInReset _
    Lib "winmm.dll" ( _
        ByVal hWaveIn As Long) _
    As Long

Declare Function waveInStart _
    Lib "winmm.dll" ( _
        ByVal hWaveIn As Long) _
    As Long

Declare Function waveInStop _
    Lib "winmm.dll" ( _
        ByVal hWaveIn As Long) _
    As Long

Declare Function waveInUnprepareHeader _
    Lib "winmm.dll" ( _
        ByVal hWaveIn As Long, _
        lpWaveInHdr As WAVEHDR, _
        ByVal uSize As Long) _
    As Long

Declare Function waveOutBreakLoop _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long) _
    As Long

Declare Function waveOutClose _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long) _
    As Long

Declare Function waveOutGetDevCaps _
    Lib "winmm.dll" _
    Alias "waveOutGetDevCapsA" ( _
        ByVal uDeviceID As Long, _
        lpCaps As WAVEOUTCAPS, _
        ByVal uSize As Long) _
    As Long

Declare Function waveOutGetErrorText _
    Lib "winmm.dll" _
    Alias "waveOutGetErrorTextA" ( _
        ByVal err As Long, _
        ByVal lpText As String, _
        ByVal uSize As Long) _
    As Long

Declare Function waveOutGetID _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long, _
        lpuDeviceID As Long) _
    As Long

Declare Function waveOutGetNumDevs _
    Lib "winmm.dll" ( _
        ) _
    As Long

Declare Function waveOutGetPitch _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long, _
        lpdwPitch As Long) _
    As Long

Declare Function waveOutGetPlaybackRate _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long, _
        lpdwRate As Long) _
    As Long

Declare Function waveOutGetPosition _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long, _
        lpInfo As MMTIME, _
        ByVal uSize As Long) _
    As Long

Declare Function waveOutGetVolume _
    Lib "winmm.dll" ( _
        ByVal uDeviceID As Long, _
        lpdwVolume As Long) _
    As Long

Declare Function waveOutMessage _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long, _
        ByVal msg As Long, _
        ByVal dw1 As Long, _
        ByVal dw2 As Long) _
    As Long

Declare Function waveOutOpen _
    Lib "winmm.dll" ( _
        lphWaveOut As Long, ByVal uDeviceID As Long, _
        lpFormat As WAVEFORMAT, ByVal dwCallback As Long, _
        ByVal dwInstance As Long, ByVal dwFlags As Long) _
    As Long

Declare Function waveOutPause _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long) _
    As Long

Declare Function waveOutPrepareHeader _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long, _
        lpWaveOutHdr As WAVEHDR, _
        ByVal uSize As Long) _
    As Long

Declare Function waveOutReset _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long) _
    As Long

Declare Function waveOutRestart _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long) _
    As Long

Declare Function waveOutSetPitch _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long, _
        ByVal dwPitch As Long) _
    As Long

Declare Function waveOutSetPlaybackRate _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long, _
        ByVal dwRate As Long) _
    As Long

Declare Function waveOutSetVolume _
    Lib "winmm.dll" ( _
        ByVal uDeviceID As Long, _
        ByVal dwVolume As Long) _
    As Long

Declare Function waveOutUnprepareHeader _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long, _
        lpWaveOutHdr As WAVEHDR, _
        ByVal uSize As Long) _
    As Long

Declare Function waveOutWrite _
    Lib "winmm.dll" ( _
        ByVal hWaveOut As Long, _
        lpWaveOutHdr As WAVEHDR, _
        ByVal uSize As Long) _
    As Long


