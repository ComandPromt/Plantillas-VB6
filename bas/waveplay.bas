Attribute VB_Name = "Module1"
Option Explicit
'Basic Wave Playing by J W Lehman
Declare Function sndPlaySound Lib "MMSYSTEM.DLL" (ByVal lpszSoundName$, ByVal wFlags%) As Integer
   Global Const SND_SYNC = &H0
   Global Const SND_ASYNC = &H1
   Global Const SND_NODEFAULT = &H2
   Global Const SND_LOOP = &H8
   Global Const SND_NOSTOP = &H10

