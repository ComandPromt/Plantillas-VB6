Attribute VB_Name = "BeatThi"
Option Explicit
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function waveOutGetNumDevs Lib "winmm" () As Long

Global Const SND_SYNC = &H0 'just after the sound is ended exit function
Global Const SND_ASYNC = &H1 'just after the beginning of the sound exit function
Global Const SND_NODEFAULT = &H2 'if the sound cannot be found no error message
Global Const SND_LOOP = &H8 'repeat the sound until the function is called again
Global Const SND_NOSTOP = &H10 'if currently a sound is played the function will return without playing the selected sound

Global Const Flags& = SND_ASYNC Or SND_NODEFAULT
