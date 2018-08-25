Attribute VB_Name = "PIANO1"
#If Win32 Then
   Public Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
   Public Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
   Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
   Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
   Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
   Public Declare Function midiOutOpen Lib "mmsystem.dll" Alias "MidiOutOpen" (hMidiOut As Long, ByVal DeviceId As Integer, ByVal C As Long, ByVal I As Long, ByVal F As Long) As Integer
   Public Declare Function midiOutShortMsg Lib "mmsystem.dll" Alias "MidiOutShortMsg" (ByVal hMidiOut As Integer, ByVal MidiMessage As Long) As Integer
   Public Declare Function midiOutClose Lib "mmsystem.dll" Alias "MidiOutClose" (ByVal hMidiOut As Integer) As Integer
   Public Declare Function GetPrivateProfileString Lib "kernel" (ByVal Sname$, ByVal Kname$, ByVal Def$, ByVal Ret$, ByVal Size%, ByVal Fname$) As Integer
   Public Declare Function sndPlaySound Lib "mmsystem" (ByVal lpsSound As String, ByVal wFlag As Integer) As Integer
#End If

Global MidiEventOut, MidiNoteOut, MidiVelOut As Long
Global hMidiOut As Long
Global hMidiOutCopy As Long
Global MidiOpenError As String

Global Const MODAL = 1
Global Const ShiftKey = 1

' The Patch number array used for current patch for each midi channel
' Then Volume array used for each channels volume setting
' TrackChannel is array for the current midi channel that that Track on the mixi is set to.
Global MidiPatch(16), MidiVolume(16), TrackChannel(16), MidiPan(16), Octave(16) As Integer

' The current Midi Channel out set on Piano form
Global MidiChannelOut As Integer

' The Velocity (Volume) of notes for current midi channel
Global MidiVelocity As Integer

'Boolean for it CapsLock has been pressed or not
Global CapsLock As Integer

' NoteRepeat used to stop the same key from repeating.  CapsLock detects if it is down.
Global NoteRepeat As Integer


Sub MidiOutOpenPort()
   MidiOpenError = Str$(midiOutOpen(hMidiOut, -1, 0, 0, 0))
   hMidiOutCopy = hMidiOut
End Sub

Sub ReadPatch()
Dim Sname As String, Ret As String, Ext As String
    
   Ret = String$(255, 0)
   Default1$ = Ret
   Sname = "General MIDI"
   Ext = Str$(MidiPatch(MidiChannelOut))
   filename$ = App.Path & "\PATCH.INI"
   nSize = GetPrivateProfileString(Sname, Ext, Default1$, Ret, Len(Ret), filename$)
   Piano.PatchLabel.Caption = Ret
End Sub

Sub SendMidiOut()
Dim MidiMessage As Long
Dim lowint As Long
Dim highint As Long
    
   lowint = (MidiNoteOut * 256) + MidiEventOut
   highint = (MidiVelOut * 256) * 256

   MidiMessage = lowint + highint
   X% = midiOutShortMsg(hMidiOutCopy, MidiMessage)
End Sub

