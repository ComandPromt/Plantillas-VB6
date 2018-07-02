Attribute VB_Name = "basSound"
Option Explicit

Private Const SND_ALIAS = &H10000
Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000
Private Const SND_LOOP = &H8
Private Const SND_NODEFAULT = &H2
Private Const SND_NOSTOP = &H10
Private Const SND_NOWAIT = &H2000
Private Const SND_SYNC = &H0

Private Declare Function PlaySound Lib "winmm.dll" Alias _
   "PlaySoundA" (ByVal lpszName As String, _
   ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long
Private Declare Function midiOutGetNumDevs Lib "winmm" () As Integer

Private Declare Function midiOutGetVolume Lib "winmm.dll" _
   (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function waveOutGetVolume Lib "winmm.dll" _
   (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
   
Public Declare Function midiOutSetVolume Lib "winmm.dll" _
   (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function waveOutSetVolume Lib "winmm.dll" _
   (ByVal uDeviceID As Long, lpdwVolume As Long) As Long

Private Declare Function mciSendString Lib "winmm.dll" Alias _
   "mciSendStringA" (ByVal lpstrCommand As String, _
   ByVal lpstrReturnString As String, ByVal uReturnLength As Long, _
   ByVal hwndCallback As Long) As Long
   
   
Private Const MMSYSERR_NOERROR = 0



Public Const AUDIO_NONE = 0
Public Const AUDIO_WAVE = 1
Public Const AUDIO_MIDI = 2

'
' Returns 1 if wave output
' Returns 2 if midi output
' Returns 3 if both
'
Public Function CanPlaySound() As Integer
   Dim i As Integer

   i = AUDIO_NONE
   
   If waveOutGetNumDevs > 0 Then
      i = AUDIO_WAVE
   End If
   
   If midiOutGetNumDevs > 0 Then
      i = i + AUDIO_MIDI
   End If
   
   CanPlaySound = i
End Function

'
' Bug: Does not work correctly
Public Function GetVolume(Optional rt As Variant, Optional lt As Variant, Optional audiotype As Variant) As Integer
   Dim i As Long
   Dim k As Integer
   
   rt = 0
   lt = 0
   k = 0
   
   If IsMissing(audiotype) Then
      audiotype = AUDIO_MIDI + AUDIO_WAVE
   End If
   
   If (audiotype And AUDIO_MIDI) = AUDIO_MIDI Then
      midiOutGetVolume 0, i
      rt = ((i And &HFFFF0000) \ &HFFFF&) And &HFFFF&
      lt = i And &HFFFF&
      k = 1
   End If
   
   If (audiotype And AUDIO_WAVE) = AUDIO_WAVE Then
      waveOutGetVolume 0, i
      rt = rt + ((i And &HFFFF0000) / &H10000) And &HFFFF&
      lt = lt + (i And &HFFFF&)
      k = k + 1
   End If

   If k = 0 Then
      GetVolume = 0
   Else
      GetVolume = (rt + lt) / (k * 2)
      rt = rt / k
      lt = lt / k
   End If
End Function


'
'
' Bug: Does not work correctly
Public Sub SetVolume(ByVal rt As Integer, ByVal lt As Integer, Optional audiotype As Variant)
   If IsMissing(audiotype) Then
      audiotype = AUDIO_MIDI + AUDIO_WAVE
   End If
   
   If (audiotype And AUDIO_MIDI) = AUDIO_MIDI Then
      midiOutSetVolume 0, (rt * &HFFFF&) + lt
   End If
   
   If (audiotype And AUDIO_WAVE) = AUDIO_WAVE Then
      waveOutSetVolume 0, (rt * &HFFFF&) + lt
   End If
End Sub


'
' Typical system sounds constant across all windows platforms
'
'    SystemQuestion
'    SystemStart
'    SystemAsterisk
'    SystemExclamation
'    SystemExit
'    SystemHand
'
'  Returns true if success, false if failed.
'  async assumes true
'  loop assumes false
Public Function SoundPlay(filename As String, Optional async As Variant, Optional sLoop As Variant) As Boolean
   Dim i As Integer
   Dim f As String
   Dim j As Long
         
   i = Len(filename)
   f = UCase(filename)
   
   If IsMissing(async) Then
      j = SND_ASYNC
   Else
      If async Then
         j = SND_ASYNC
      Else
         j = SND_SYNC
      End If
   End If
   
   If Not IsMissing(sLoop) Then
      If sLoop And (j = SND_ASYNC) Then
         j = j + SND_LOOP
      End If
   End If
   
   j = j + SND_NOSTOP + SND_NOWAIT
   
   If InStr(f, ".WAV") = i - 3 Then
      If CanPlaySound And AUDIO_WAVE = AUDIO_WAVE Then
         j = j + SND_FILENAME + SND_NODEFAULT
         i = PlaySound(filename, 0, j)
         SoundPlay = IIf(i = 0, False, True)
      Else
         Beep
         SoundPlay = True
      End If
      
   'Assume media player for other file names   .MID .RMI etc..
   ElseIf InStr(f, ".") = i - 3 Then
      If CanPlaySound And AUDIO_MIDI = AUDIO_MIDI Then
         i = mciSendString("open " & filename & " type sequencer alias filename", 0&, 0, 0)
         'Note the true/false order is supposed to be opposite of the others.
         SoundPlay = IIf(i = 0, True, False)
         If (j And SND_ASYNC) = SND_ASYNC Then
            If (j And SND_LOOP) = SND_LOOP Then
               'Bug: repeat doesn't work.
               mciSendString "play filename repeat", 0&, 0, 0
            Else
               mciSendString "play filename", 0&, 0, 0
            End If
         Else
            mciSendString "play filename wait", 0&, 0, 0
            mciSendString "close filename", 0&, 0, 0
         End If
      Else
         Beep
         SoundPlay = True
      End If
   Else
      j = j + SND_ALIAS
      i = PlaySound(filename, 0, j)
      SoundPlay = IIf(i = 0, False, True)
   End If
End Function

Public Function SoundStop(Optional audiotype As Variant)
   If IsMissing(audiotype) Then
      mciSendString "close filename", 0&, 0, 0
      SoundPlay vbNullString, 0, 0
   Else
      If (audiotype And AUDIO_MIDI) = AUDIO_MIDI Then
         mciSendString "close filename", 0&, 0, 0
      End If
      If (audiotype And AUDIO_WAVE) = AUDIO_WAVE Then
         SoundPlay vbNullString, 0, 0
      End If
   End If
End Function

