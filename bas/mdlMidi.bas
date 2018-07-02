Attribute VB_Name = "mdlMidi"
'All Constants Types and Declarations for MIDI.
'This is a general module I use over seperate projects
Dim h_midiout As Long
Const MaxPNameLen = 32
Public Const MHDR_VALID = &H7
Public Const MHDR_DONE = &H1
Public Const MHDR_INQUEUE = &H4
Public Const MHDR_PREPARED = &H2
Public Const MIDI_IO_STATUS = &H20&
Public Const MIDI_MAPPER = -1&
Public Const MIDI_UNCACHE = 4
Public Const MIDICAPS_CACHE = &H4
Public Const MIDICAPS_LRVOLUME = &H2
Public Const MIDICAPS_STREAM = &H8
Public Const MIDICAPS_VOLUME = &H1
Public Const MIDIERR_BASE = 64
Public Const MIDIERR_INVALIDSETUP = (MIDIERR_BASE + 5)
Public Const MIDIERR_LASTERROR = (MIDIERR_BASE + 5)
Public Const MIDIERR_NODEVICE = (MIDIERR_BASE + 4)
Public Const MIDIERR_NOMAP = (MIDIERR_BASE + 2)
Public Const MIDIERR_NOTREADY = (MIDIERR_BASE + 3)
Public Const MIDIERR_STILLPLAYING = (MIDIERR_BASE + 1)
Public Const MIDIERR_UNPREPARED = (MIDIERR_BASE + 0)
Public Const MIDIMAPPER = (-1)
Public Const MIDIPROP_GET = &H40000000
Public Const MIDIPROP_SET = &H80000000
Public Const MIDIPROP_TEMPO = &H2&
Public Const MIDIPROP_TIMEDIV = &H1&
Public Const MIDISTRM_ERROR = -2

Type MIDIOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MaxPNameLen
        wTechnology As Integer
        wVoices As Integer
        wNotes As Integer
        wChannelMask As Integer
        dwSupport As Long
End Type
Type MIDIEVENT
        dwDeltaTime As Long          '  Ticks since last event
        dwStreamID As Long           '  Reserved; must be zero
        dwEvent As Long              '  Event type and parameters
        dwParms(1) As Long           '  Parameters if this is a long event
End Type
Type MIDIHDR
        lpData As String
        dwBufferLength As Long
        dwBytesRecorded As Long
        dwUser As Long
        dwFlags As Long
        lpNext As Long
        Reserved As Long
End Type
Type MIDIINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MaxPNameLen
End Type
Type MIDIPROPTEMPO
        cbStruct As Long
        dwTempo As Long
End Type
Type MIDIPROPTIMEDIV
        cbStruct As Long
        dwTimeDiv As Long
End Type
Type MIDISTRMBUFFVER
        dwVersion As Long                  '  Stream buffer format version
        dwMid As Long                      '  Manufacturer ID as defined in MMREG.H
        dwOEMVersion As Long               '  Manufacturer version for custom ext
End Type
Type MMTIME
        wType As Long
        u As Long
End Type
Declare Function midiConnect Lib "winmm.dll" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
Declare Function midiDisconnect Lib "winmm.dll" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIINCAPS, ByVal uSize As Long) As Long
Declare Function midiInGetErrorText Lib "winmm.dll" Alias "midiInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function midiInGetID Lib "winmm.dll" (ByVal hMidiIn As Long, lpuDeviceID As Long) As Long
Declare Function midiInGetNumDevs Lib "winmm.dll" () As Long
Declare Function midiInMessage Lib "winmm.dll" (ByVal hMidiIn As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function midiInOpen Lib "winmm.dll" (lphMidiIn As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiInPrepareHeader Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInUnprepareHeader Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function midiOutGetID Lib "winmm.dll" (ByVal hMidiOut As Long, lpuDeviceID As Long) As Long
Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Declare Function midiOutLongMsg Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutMessage Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiOutPrepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutReset Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Declare Function midiOutUnprepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiStreamClose Lib "winmm.dll" (ByVal hms As Long) As Long
Declare Function midiStreamOpen Lib "winmm.dll" (phms As Long, puDeviceID As Long, ByVal cMidi As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function midiStreamOut Lib "winmm.dll" (ByVal hms As Long, pmh As MIDIHDR, ByVal cbmh As Long) As Long
Declare Function midiStreamPause Lib "winmm.dll" (ByVal hms As Long) As Long
Declare Function midiStreamPosition Lib "winmm.dll" (ByVal hms As Long, lpmmt As MMTIME, ByVal cbmmt As Long) As Long
Declare Function midiStreamProperty Lib "winmm.dll" (ByVal hms As Long, lppropdata As Byte, ByVal dwProperty As Long) As Long
Declare Function midiStreamRestart Lib "winmm.dll" (ByVal hms As Long) As Long
Declare Function midiStreamStop Lib "winmm.dll" (ByVal hms As Long) As Long



