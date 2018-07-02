Attribute VB_Name = "VBSYXGBL"
Option Explicit
' MIDI output device capabilities structure
Type MIDIOUTCAPS
    wMid As Integer                ' Manufacturer ID
    wPid As Integer                ' Product ID
    vDriverVersion As Long         ' Driver version
    szPname As String * 32         ' Product name (NULL terminated string)
    wTechnology As Integer         ' Device type
    wVoices As Integer             ' n. of voices (internal synth only)
    wNotes As Integer              ' max n. of notes (internal synth only)
    wChannelMask As Integer        ' n. of Midi channels (internal synth only)
    dwSupport As Long              ' Supported extra controllers (volume, etc)
End Type
' MIDI data block header
Type MIDIHDR
    lpData As Long ' was String! '  pointer to locked data block
    dwBufferLength As Long       '  length of data in data block
    dwBytesRecorded As Long      '  used for input only
    dwUser As Long               '  for client's use
    dwFlags As Long              '  assorted flags (see defines)
    lpNext As Long               '  reserved for driver
    reserved As Long             '  reserved for driver
End Type
'return value from API Functions, handles etc
Global vntRet As Variant
Global hMidi As Long
Global mDev As Long
' VBSYX Variables
Global MidiSYX, outSYX, L1, L2, L3, L4, syxchg
