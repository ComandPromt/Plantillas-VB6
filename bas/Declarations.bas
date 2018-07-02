Attribute VB_Name = "Declarations"
Option Explicit

'Taken from WIN32API.TXT

Public Const MAXPNAMELEN = 32       '  max product name length (including NULL)

Type MIXERCAPS
    wMid As Integer                 '  manufacturer id
    wPid As Integer                 '  product id
    vDriverVersion As Long          '  version of the driver
    szPname As String * MAXPNAMELEN '  product name
    fdwSupport As Long              '  misc. support bits
    cDestinations As Long           '  count of destinations
End Type

Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, ByRef pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long

