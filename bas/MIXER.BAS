Attribute VB_Name = "basMixer"
Option Explicit

Declare Function agGetAddressForObject Lib "apigid32.dll" (object As Any) As Long

Public Const MAXPNAMELEN = 32

' ---------
' Mixer errors
' ---------
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)
Public Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)
Public Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)
Public Const MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12)
Public Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10)
Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)
Public Const MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13)
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)
Public Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 13)
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)
Public Const MMSYSERR_NOERROR = 0
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)
Public Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)
Public Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)

Public Const MIDIERR_BASE = 64
Public Const MIDIERR_NODEVICE = (MIDIERR_BASE + 4)

'*******************
' Mixer
'*******************

Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_LONG_NAME_CHARS = 64

' Mixer control get/set details
Public Const MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1&
Public Const MIXER_GETCONTROLDETAILSF_QUERYMASK = &HF&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&

Public Const MIXER_SETCONTROLDETAILSF_CUSTOM = &H1&
Public Const MIXER_SETCONTROLDETAILSF_QUERYMASK = &HF&
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&

' Mixer get line controls
Public Const MIXER_GETLINECONTROLSF_ALL = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Public Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&

' Mixer get line info of
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETLINEINFOF_DESTINATION = &H0&
Public Const MIXER_GETLINEINFOF_LINEID = &H2&
Public Const MIXER_GETLINEINFOF_QUERYMASK = &HF&
Public Const MIXER_GETLINEINFOF_SOURCE = &H1&
Public Const MIXER_GETLINEINFOF_TARGETTYPE = &H4&

'
Public Const MIXER_OBJECTF_AUX = &H50000000
Public Const MIXER_OBJECTF_HANDLE = &H80000000
Public Const MIXER_OBJECTF_MIDIIN = &H40000000
Public Const MIXER_OBJECTF_MIDIOUT = &H30000000
Public Const MIXER_OBJECTF_MIXER = &H0&
Public Const MIXER_OBJECTF_WAVEIN = &H20000000
Public Const MIXER_OBJECTF_WAVEOUT = &H10000000

Public Const MIXER_OBJECTF_HMIDIIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIIN)
Public Const MIXER_OBJECTF_HMIDIOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)
Public Const MIXER_OBJECTF_HMIXER = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)
Public Const MIXER_OBJECTF_HWAVEIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)
Public Const MIXER_OBJECTF_HWAVEOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)

' Control flag
Public Const MIXERCONTROL_CONTROLF_DISABLED = &H80000000
Public Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2&
Public Const MIXERCONTROL_CONTROLF_UNIFORM = &H1&

' Mixer control classes

Public Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Public Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Public Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Public Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Public Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Public Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Public Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000
Public Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Public Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000

' Mixer control types
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_BUTTON = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
Public Const MIXERCONTROL_CONTROLTYPE_DECIBELS = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
Public Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Public Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
Public Const MIXERCONTROL_CONTROLTYPE_MICROTIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MILLITIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Public Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_STEREOENH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
Public Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Public Const MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_ONOFF = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
Public Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_PERCENT = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
Public Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)

' Mixer line flags
Public Const MIXERLINE_LINEF_ACTIVE = &H1&
Public Const MIXERLINE_LINEF_DISCONNECTED = &H8000&
Public Const MIXERLINE_LINEF_SOURCE = &H80000000

' Mixer line target types
Public Const MIXERLINE_TARGETTYPE_AUX = 5
Public Const MIXERLINE_TARGETTYPE_MIDIIN = 4
Public Const MIXERLINE_TARGETTYPE_MIDIOUT = 3
Public Const MIXERLINE_TARGETTYPE_UNDEFINED = 0
Public Const MIXERLINE_TARGETTYPE_WAVEIN = 2
Public Const MIXERLINE_TARGETTYPE_WAVEOUT = 1

' Destination and source component types
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = 0
Public Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = 1
Public Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = 5
Public Const MIXERLINE_COMPONENTTYPE_DST_LAST = 8
Public Const MIXERLINE_COMPONENTTYPE_DST_LINE = 2
Public Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = 3
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = 4
Public Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = 6
Public Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = 0
Public Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN = 8
Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = 7
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG = &H1000& + 10
Public Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = &H1000& + 9
Public Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = &H1000& + 5
Public Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = &H1000& + 1
Public Const MIXERLINE_COMPONENTTYPE_SRC_LAST = &H1000& + 10
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = &H1000& + 2
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = &H1000& + 3
Public Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = &H1000& + 7
Public Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = &H1000& + 4
Public Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = &H1000& + 6
Public Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = &H1000& + 0
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = &H1000& + 8

' Error constants
Public Const MIXERR_INVALLINE = 1024 + 0
Public Const MIXERR_BASE = 1024
Public Const MIXERR_INVALCONTROL = 1024 + 1
Public Const MIXERR_INVALVALUE = 1024 + 2
Public Const MIXERR_LASTERROR = 1024 + 2

' Product IDs

' -----
' TYPES
' -----
' Mixer type
Type MIXERCAPS
     wMid As Integer                   '  manufacturer id
     wPid As Integer                   '  product id
     vDriverVersion As Long            '  version of the driver
     szPname As String * MAXPNAMELEN   '  product name
     fdwSupport As Long             '  misc. support bits
     cDestinations As Long          '  count of destinations
End Type

' Mixer line types
Type Target
     dwType As Long                 '  MIXERLINE_TARGETTYPE_xxxx
     dwDeviceID As Long             '  target device ID of device type
     wMid As Integer                   '  of target device
     wPid As Integer                   '       "
     vDriverVersion As Long            '       "
     szPname As String * MAXPNAMELEN
End Type

Type MIXERLINE
     cbStruct As Long               '  size of MIXERLINE structure
     dwDestination As Long          '  zero based destination index
     dwSource As Long               '  zero based source index (if source)
     dwLineID As Long               '  unique line id for mixer device
     fdwLine As Long                '  state/information about line
     dwUser As Long                 '  driver specific information
     dwComponentType As Long        '  component type line connects to
     cChannels As Long              '  number of channels line supports
     cConnections As Long           '  number of connections (possible)
     cControls As Long              '  number of controls at this line
     szShortName As String * MIXER_SHORT_NAME_CHARS
     szName As String * MIXER_LONG_NAME_CHARS
     lpTarget As Target
End Type

' MM Control types
Type MIXERLINECONTROLS
     cbStruct As Long         '  size in Byte of MIXERLINECONTROLS
     dwLineID As Long         '  line id (from MIXERLINE.dwLineID)
     dwControl As Long        '  used with MIXER_GETLINECONTROLSF_ONEBYTYPE or MIXER_GETLINECONTROLSF_ONEBYID
     cControls As Long        '  count of controls pmxctrl points to
     cbmxctrl As Long         '  size in Byte of _one_ MIXERCONTROL
     pamxctrl As Long         '  pointer to first MIXERCONTROL array
End Type

Type MIXERCONTROL
     cbStruct As Long           '  size in Byte of MIXERCONTROL
     dwControlID As Long        '  unique control id for mixer device
     dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
     fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
     cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
     szShortName(1 To MIXER_SHORT_NAME_CHARS) As Byte
     szName(1 To MIXER_LONG_NAME_CHARS) As Byte
     Bounds(1 To 6) As Long
     Metrics(1 To 6) As Long
End Type

Type MIXERCONTROLDETAILS
     cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
     dwControlID As Long    '  control id to get/set details on
     cChannels As Long      '  number of channels in paDetails array
     item As Long                           ' hwndOwner or cMultipleItems
     cbDetails As Long      '  size of _one_ details_XX struct
     paDetails As Long      '  pointer to array of details_XX structs
End Type

Type MIXERCONTROLDETAILS_BOOLEAN
     fValue As Long
End Type

Type MIXERCONTROLDETAILS_LISTTEXT
     dwParam1 As Long
     dwParam2 As Long
     szName As String * MIXER_LONG_NAME_CHARS
End Type

Type MIXERCONTROLDETAILS_SIGNED
     lValue As Long
End Type

Type MIXERCONTROLDETAILS_UNSIGNED
     dwValue As Long
End Type

' --------
' DECLARES
' --------
Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, pumxID As Long, ByVal fdwId As Long) As Long
Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, ByVal uMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long

Public Function Trim0(sName As String) As String

' Keep left portion of string sName up to first 0.
' Useful with Win API null terminated strings.

Dim x As Integer
x = InStr(sName, Chr$(0))
If x > 0 Then Trim0 = Left$(sName, x - 1) Else Trim0 = sName

End Function

Function GetLineFlag(vValue As Variant) As String

' Function returns name of constant for given value.

Dim sName As String

sName = ""
If vValue And &H1& Then sName = sName & " Active"
If vValue And &H8000& Then sName = sName & " Disconnected"
If vValue And &H80000000 Then sName = sName & " Source"

GetLineFlag = Trim(sName)

End Function
Function GetComponentType(vValue As Variant) As String

' Function returns name of constant for given value.

Dim sName As String

Select Case vValue

   Case 0
      sName = "MIXERLINE_COMPONENTTYPE_DST_FIRST"

   Case 1
      sName = "MIXERLINE_COMPONENTTYPE_DST_DIGITAL"

   Case 5
      sName = "MIXERLINE_COMPONENTTYPE_DST_HEADPHONES"

   Case 8
      sName = "MIXERLINE_COMPONENTTYPE_DST_LAST"

   Case 2
      sName = "MIXERLINE_COMPONENTTYPE_DST_LINE"

   Case 3
      sName = "MIXERLINE_COMPONENTTYPE_DST_MONITOR"

   Case 4
      sName = "MIXERLINE_COMPONENTTYPE_DST_SPEAKERS"

   Case 6
      sName = "MIXERLINE_COMPONENTTYPE_DST_TELEPHONE"

   Case 0
      sName = "MIXERLINE_COMPONENTTYPE_DST_UNDEFINED"

   Case 8
      sName = "MIXERLINE_COMPONENTTYPE_DST_VOICEIN"

   Case 7
      sName = "MIXERLINE_COMPONENTTYPE_DST_WAVEIN"

   Case &H1000&
      sName = "MIXERLINE_COMPONENTTYPE_SRC_FIRST"

   Case &H1000& + 10
      sName = "MIXERLINE_COMPONENTTYPE_SRC_ANALOG"

   Case &H1000& + 9
      sName = "MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY"

   Case &H1000& + 5
      sName = "MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC"

   Case &H1000& + 1
      sName = "MIXERLINE_COMPONENTTYPE_SRC_DIGITAL"

   Case &H1000& + 10
      sName = "MIXERLINE_COMPONENTTYPE_SRC_LAST"

   Case &H1000& + 2
      sName = "MIXERLINE_COMPONENTTYPE_SRC_LINE"

   Case &H1000& + 3
      sName = "MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE"

   Case &H1000& + 7
      sName = "MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER"

   Case &H1000& + 4
      sName = "MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER"

   Case &H1000& + 6
      sName = "MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE"

   Case &H1000& + 0
      sName = "MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED"

   Case &H1000& + 8
      sName = "MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT"

   Case Else
      sName = "<invalid>"

End Select

GetComponentType = sName

End Function
Function GetTargetType(vValue As Variant) As String

' Function returns name of constant for given value.

Dim sName As String

Select Case vValue

   Case 5
      sName = "MIXERLINE_TARGETTYPE_AUX"

   Case 4
      sName = "MIXERLINE_TARGETTYPE_MIDIIN"

   Case 3
      sName = "MIXERLINE_TARGETTYPE_MIDIOUT"

   Case 0
      sName = "MIXERLINE_TARGETTYPE_UNDEFINED"

   Case 2
      sName = "MIXERLINE_TARGETTYPE_WAVEIN"

   Case 1
      sName = "MIXERLINE_TARGETTYPE_WAVEOUT"

   Case Else
      sName = "<invalid>"

End Select

GetTargetType = sName

End Function
Function GetControlType(vValue As Variant) As String

' Function returns name of constant for given value.

Dim sName As String

Select Case vValue

   Case (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_FADER"

   Case (MIXERCONTROL_CONTROLTYPE_FADER + 2)
      sName = "MIXERCONTROL_CONTROLTYPE_BASS"

   Case (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
      sName = "MIXERCONTROL_CONTROLTYPE_BOOLEAN"

   Case (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
      sName = "MIXERCONTROL_CONTROLTYPE_BOOLEANMETER"

   Case (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
      sName = "MIXERCONTROL_CONTROLTYPE_BUTTON"

   Case (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
      sName = "MIXERCONTROL_CONTROLTYPE_CUSTOM"

   Case (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
      sName = "MIXERCONTROL_CONTROLTYPE_DECIBELS"

   Case (MIXERCONTROL_CONTROLTYPE_FADER + 4)
      sName = "MIXERCONTROL_CONTROLTYPE_EQUALIZER"

   Case (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
      sName = "MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT"

   Case (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
      sName = "MIXERCONTROL_CONTROLTYPE_LOUDNESS"

   Case (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_MICROTIME"

   Case (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_MILLITIME"

   Case (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_MIXER"

   Case (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
      sName = "MIXERCONTROL_CONTROLTYPE_MONO"

   Case (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_SLIDER"

   Case (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
      sName = "MIXERCONTROL_CONTROLTYPE_STEREOENH"

   Case (MIXERCONTROL_CONTROLTYPE_FADER + 3)
      sName = "MIXERCONTROL_CONTROLTYPE_TREBLE"

   Case (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_UNSIGNED"

   Case (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER"

   Case (MIXERCONTROL_CONTROLTYPE_FADER + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_VOLUME"

   Case (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_SIGNEDMETER"

   Case (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
      sName = "MIXERCONTROL_CONTROLTYPE_SINGLESELECT"

   Case (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
      sName = "MIXERCONTROL_CONTROLTYPE_MUTE"

   Case (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_MUX"

   Case (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_ONOFF"

   Case (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_PAN"

   Case (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
      sName = "MIXERCONTROL_CONTROLTYPE_PEAKMETER"

   Case (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
      sName = "MIXERCONTROL_CONTROLTYPE_PERCENT"

   Case (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
      sName = "MIXERCONTROL_CONTROLTYPE_QSOUNDPAN"

   Case (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)
      sName = "MIXERCONTROL_CONTROLTYPE_SIGNED"

   Case Else
      sName = "<invalid>"

End Select

GetControlType = sName

End Function
Function GetControlFlag(vValue As Variant) As String

' Function returns name of constant for given value.

Dim sName As String

sName = ""
If (vValue And &H1&) > 0 Then sName = sName & " Uniform"
If (vValue And &H2&) > 0 Then sName = sName & " Multiple"
If (vValue And &H80000000) > 0 Then sName = sName & " Disabled"

GetControlFlag = Trim(sName)

End Function
