Attribute VB_Name = "mMixer"

Option Explicit

'/****************************************************************************
'
'mMixer bas module for VB5
'Copyright (c) 1998 by Ray Mercer, all rights reserved
'Additions by Dave Snyder
'defines and declares for Win32 audio mixer functions
'Exclusively from http://i.am/shrinkwrapvb
'
'no warranty expressed or implied
'
'****************************************************************************/
'

Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MAXPNAMELEN = 32


Public Const HIGHEST_VOLUME_SETTING = 65535

'/****************************************************************************
'
'            General error return values
'
'****************************************************************************/
Public Const MMSYSERR_BASE = 0
Public Const WAVERR_BASE = 32
'#define MIDIERR_BASE           64
'#define TIMERR_BASE            96
'#define JOYERR_BASE            160
'#define MCIERR_BASE            256
Public Const MIXERR_BASE = 1024
'#define MCI_STRING_OFFSET      512
'#define MCI_VD_OFFSET          1024
'#define MCI_CD_OFFSET          1088
'#define MCI_WAVE_OFFSET        1152
'#define MCI_SEQ_OFFSET         1216
'
'/* general error return values */
Public Enum vbMMSYSERRORS
    MMSYSERR_NOERROR = 0                        '/* no error */
    MMSYSERR_ERROR = (MMSYSERR_BASE + 1)        '/* unspecified error */
    MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)  '/* device ID out of range */
    MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)   '/* driver failed enable */
    MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)    '/* device already allocated */
    MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)  '/* device handle is invalid */
    MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)     '/* no device driver present */
    MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)        '/* memory allocation error */
    MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8) '/* function isn't supported */
    MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)    '/* error value out of range */
    MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10)   '/* invalid flag passed */
    MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)  '/* invalid parameter passed */
    MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12)  '/* handle being used simultaneously on another thread (eg callback) */
    MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13) '/* specified alias not found */
    MMSYSERR_BADDB = (MMSYSERR_BASE + 14)       '/* bad registry database */
    MMSYSERR_KEYNOTFOUND = (MMSYSERR_BASE + 15) '/* registry key not found */
    MMSYSERR_READERROR = (MMSYSERR_BASE + 16)   '/* registry read error */
    MMSYSERR_WRITEERROR = (MMSYSERR_BASE + 17)  '/* registry write error */
    MMSYSERR_DELETEERROR = (MMSYSERR_BASE + 18) '/* registry delete error */
    MMSYSERR_VALNOTFOUND = (MMSYSERR_BASE + 19) '/* registry value not found */
    MMSYSERR_NODRIVERCB = (MMSYSERR_BASE + 20)  '/* driver does not call DriverCallback */
'/* */
'/*  MMRESULT error return values specific to the mixer API */
'/* */
'/* */
    MIXERR_INVALLINE = (MIXERR_BASE + 0)
    MIXERR_INVALCONTROL = (MIXERR_BASE + 1)
    MIXERR_INVALVALUE = (MIXERR_BASE + 2)
End Enum


'/* */
'/*  MIXERCONTROL_CONTROLTYPE_xxx building block defines */
'/* */
'/* */
Public Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Public Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Public Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Public Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
'
Public Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000
'
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
'
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
'
Public Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
'
Public Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Public Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
'
'Public Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Public Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000  ' /* in 10ths */
Public Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000          '/* in 10ths */

'/* */
'/*  Commonly used control types for specifying MIXERCONTROL.dwControlType */
'/* */
'
Public Const MIXERCONTROL_CONTROLTYPE_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_ONOFF = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Public Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Public Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
Public Const MIXERCONTROL_CONTROLTYPE_STEREOENH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
Public Const MIXERCONTROL_CONTROLTYPE_BUTTON = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_DECIBELS = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PERCENT = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
Public Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Public Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Public Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MICROTIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MILLITIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)

Public Enum MIXERCONTROL_TYPE
    mcFADER_FADER = &H50030000
    mcVOLUME_FADER = &H50030001
    mcBASS_FADER = &H50030002
    mcTREBLE_FADER = &H50030003
    mcEQUALIZER_FADER = &H50030004
    mcGENERIC_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
    mcBOOLEAN_METER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
    mcBOOLEAN_SWITCH = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
    mcONOFF_SWITCH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
    mcMUTE_SWITCH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
    mcMONO_SWITCH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
    mcLOUDNESS_SWITCH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
    mcSTEREOENH_SWITCH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
    mcBUTTON_SWITCH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 6)
    mcDECIBELS_NUMBER = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
    mcSIGNED_NUMBER = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)
    mcUNSIGNED_NUMBER = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
    mcPERCENT_NUMBER = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
    mcSLIDER_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
    mcPAN_SLIDER = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
    mcQSOUNDPAN_SLIDER = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
    mcSINGLESELECT_LIST = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
    mcMUX_LIST = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
    mcMULTIPLESELECT_LIST = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
    mcMIXER_LIST = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
    mcMICROTIME_TIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
    mcMILLITIME_TIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
    mcPEAK_METER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
End Enum


'/* */
'/*  MIXERLINE.fdwLine */
'/* */
'/* */
Public Const MIXERLINE_LINEF_ACTIVE = &H1&
Public Const MIXERLINE_LINEF_DISCONNECTED = &H8000&
Public Const MIXERLINE_LINEF_SOURCE = &H80000000

'/* */
'/*  MIXERCONTROL.fdwControl */
'/* */
'/* */
Public Const MIXERCONTROL_CONTROLF_UNIFORM = &H1&
Public Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2&
Public Const MIXERCONTROL_CONTROLF_DISABLED = &H80000000

'/* */
'/*  MIXERLINE.dwComponentType */
'/* */
'/*  component types for destinations and sources */
'/* */
'/* */
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 0)
Public Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_DST_LINE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
'Public Const MIXERLINE_COMPONENTTYPE_DST_LAST     =   (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)

Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 0)
Public Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
Public Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
'#define MIXERLINE_COMPONENTTYPE_SRC_LAST        (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)

Public Enum MIXER_LINE_TYPE
    dstUNDEFINED = &H0&
    dstDIGITAL = &H1&
    dstline = &H2&
    dstMONITOR = &H3&
    dstSPEAKERS = &H4&
    dstHEADPHONES = &H5&
    dstTELEPHONE = &H6&
    dstWAVEIN = &H7&
    dstVOICEIN = &H8&
    srcUNDEFINED = &H1000&
    srcDIGITAL = &H1001&
    srcLINE = &H1002&
    srcMICROPHONE = &H1003&
    srcSYNTHESIZER = &H1004&
    srcCOMPACTDISC = &H1005&
    srcTELEPHONE = &H1006&
    srcPCSPEAKER = &H1007&
    srcWAVEOUT = &H1008&
    srcAUXILIARY = &H1009&
    srcANALOG = &H100A&
End Enum

'
'/* */
'/*  MIXERLINE.Target.dwType */
'/* */
'/* */
Public Const MIXERLINE_TARGETTYPE_UNDEFINED = 0&
Public Const MIXERLINE_TARGETTYPE_WAVEOUT = 1&
Public Const MIXERLINE_TARGETTYPE_WAVEIN = 2&
Public Const MIXERLINE_TARGETTYPE_MIDIOUT = 3&
Public Const MIXERLINE_TARGETTYPE_MIDIIN = 4&
Public Const MIXERLINE_TARGETTYPE_AUX = 5&

Public Enum TARGET_TYPE
    ttUNDEFINED = 0&
    ttWAVEOUT = 1&
    ttWAVEIN = 2&
    ttMIDIOUT = 3&
    ttMIDIIN = 4&
    ttAUX = 5&
End Enum


Public Const MIXER_OBJECTF_HANDLE = &H80000000
Public Const MIXER_OBJECTF_MIXER = &H0&
Public Const MIXER_OBJECTF_HMIXER = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)
Public Const MIXER_OBJECTF_WAVEOUT = &H10000000
Public Const MIXER_OBJECTF_HWAVEOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)
Public Const MIXER_OBJECTF_WAVEIN = &H20000000
Public Const MIXER_OBJECTF_HWAVEIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)
Public Const MIXER_OBJECTF_MIDIOUT = &H30000000
Public Const MIXER_OBJECTF_HMIDIOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)
Public Const MIXER_OBJECTF_MIDIIN = &H40000000
Public Const MIXER_OBJECTF_HMIDIIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIIN)
Public Const MIXER_OBJECTF_AUX = &H50000000


'/**/
'/*  MIXER MESSAGES */
'/**/
Public Const MM_MIXM_LINE_CHANGE = &H3D0     '     /* mixer line change notify */
Public Const MM_MIXM_CONTROL_CHANGE = &H3D1      '  /* mixer control change notify */

'/****************************************************************************
'
'              Driver callback support
'
'****************************************************************************/
'
'/* flags used with waveOutOpen(), waveInOpen(), midiInOpen(), and */
'/* midiOutOpen() to specify the type of the dwCallback parameter. */
'
'#define CALLBACK_TYPEMASK   0x00070000l    /* callback type mask */
Public Const CALLBACK_NULL = &H0&           '/* no callback */
Public Const CALLBACK_WINDOW = &H10000      '/* dwCallback is a HWND */
'#define CALLBACK_TASK       0x00020000l    /* dwCallback is a HTASK */
Public Const CALLBACK_FUNCTION = &H30000    '/* dwCallback is a FARPROC */
'#define CALLBACK_THREAD     (CALLBACK_TASK)/* thread ID replaces 16 bit task */
'#define CALLBACK_EVENT      0x00050000l    /* dwCallback is an EVENT Handle */

Public Type MIXERCAPS
    wMid As Integer         '  manufacturer id
    wPid As Integer         '  product id
    vDriverVersion As Long  '  version of the driver
    szPname As String * MAXPNAMELEN   '  product name
    fdwSupport As Long      '  misc. support bits
    cDestinations As Long   '  count of destinations
End Type

Public Type MIXERLINE_TARGET
    dwType As Long          ' Target media device type associated with the audio
                           ' line described in the MIXERLINE structure.
    dwDeviceID As Long      ' Current device identifier of the target media device
                           ' when the dwType member is a target type other than
                           ' MIXERLINE_TARGETTYPE_UNDEFINED.
    wMid  As Integer        ' Manufacturer identifier of the target media device
                           ' when the dwType member is a target type other than
                           ' MIXERLINE_TARGETTYPE_UNDEFINED.
    wPid As Integer         ' Product identifier of the target media device when
                           ' the dwType member is a target type other than
                           ' MIXERLINE_TARGETTYPE_UNDEFINED.
    vDriverVersion As Long  ' Driver version of the target media device when the
                           ' dwType member is a target type other than
                           ' MIXERLINE_TARGETTYPE_UNDEFINED.
    szPname As String * MAXPNAMELEN  ' Product name of the target media device when
                                    ' the dwType member is a target type other than
                                    ' MIXERLINE_TARGETTYPE_UNDEFINED.
End Type
    
Public Type MIXERLINE
    cbStruct As Long      '  size of MIXERLINE structure
    dwDestination As Long '  zero based dest. index
    dwSource As Long      '  zero based source index
    dwLineID As Long      '  unique line id
    fdwLine As Long       '  information about line
    dwUser As Long        '  driver specific information
    dwComponentType As Long '  component type
    cChannels As Long     '  # of channels line supports
    cConnections As Long  '  # of connections possible
    cControls As Long     '  # of controls at this line
    szShortName As String * MIXER_SHORT_NAME_CHARS
    szName As String * MIXER_LONG_NAME_CHARS
    Target As MIXERLINE_TARGET
End Type

Public Type MIXERLINECONTROLS
    cbStruct As Long  ' size in bytes of MIXERLINECONTROLS
    dwLineID As Long  ' line id (from MIXERLINE.dwLineID)
    dwControl As Long 'UNION of next two commented lines
    'dwControlID As Long 'MIXER_GETLINECONTROLSF_ONEBYID
    'dwControlType As Long 'MIXER_GETLINECONTROLSF_ONEBYTYPE
    cControls As Long ' count of controls pmxctrl points to
    cbmxctrl As Long  ' size in bytes of _one_ MIXERCONTROL
    pamxctrl As Long 'string  ' ptr to MIXERCONTROL array
End Type

Public Type MIXERCONTROL
    cbStruct As Long           '  size in Byte of MIXERCONTROL
    dwControlID As Long        '  unique control id for mixer device
    dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
    fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
    cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
    szShortName As String * MIXER_SHORT_NAME_CHARS  ' short name of control
    szName As String * MIXER_LONG_NAME_CHARS        ' long name of control
    lMinimum As Long           '  Minimum value
    lMaximum As Long           '  Maximum value
    reserved(10) As Long       '  reserved structure space
End Type

Type MIXERCONTROLDETAILS
' The MIXERCONTROLDETAILS user defined type refers to control-detail structures,
' retrieving or setting state information of an audio mixer control. All members of this
' user-defined type must be initialized before calling the mixerGetControlDetails and
' mixerSetControlDetails functions.
   cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
   dwControlID As Long    '  control id to get/set details on
   cChannels As Long      '  number of channels in paDetails array
   Item As Long           '  hwndOwner or cMultipleItems
   cbDetails As Long      '  size of _one_ details_XX struct
   paDetails As Long      '  pointer to array of details_XX structs
End Type
      
Public Type MIXERCONTROLDETAILS_BOOLEAN
   fValue As Long
End Type
  
Public Type MIXERCONTROLDETAILS_LISTTEXT
   dwParam1 As Long
   dwParam2 As Long
   szName As String * MIXER_LONG_NAME_CHARS
End Type

  
Public Type MIXERCONTROLDETAILS_UNSIGNED
   dwValue As Long        '  value of the control
End Type

Public Type MIXERCONTROLDETAILS_SIGNED
' The MIXERCONTROLDETAILS_SIGNED user-defined type retrieves and sets signed type control
' properties for an audio mixer control.
   Lvalue As Long
End Type


Public Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long

Public Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" ( _
                                                        ByVal uMxId As Long, _
                                                        pmxcaps As MIXERCAPS, _
                                                        ByVal cbmxcaps As Long) As Long

Public Declare Function mixerOpen Lib "winmm.dll" ( _
                                        phmx As Long, _
                                        ByVal uMxId As Long, _
                                        ByVal dwCallback As Long, _
                                        ByVal dwInstance As Long, _
                                        ByVal fdwOpen As Long) As Long 'returns MMRESULT

Public Declare Function mixerClose Lib "winmm.dll" ( _
                                        ByVal hmx As Long) As Long  'returns MMRESULT

Public Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" ( _
                                                        ByVal hmxobj As Long, _
                                                        pmxl As MIXERLINE, _
                                                        ByVal fdwInfo As Long) As Long
' flags for mixerGetLineInfo()
Public Const MIXER_GETLINEINFOF_DESTINATION = &H0&
Public Const MIXER_GETLINEINFOF_SOURCE = &H1&
Public Const MIXER_GETLINEINFOF_LINEID = &H2&
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETLINEINFOF_TARGETTYPE = &H4&


Public Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" ( _
                                                        ByVal hmxobj As Long, _
                                                        pmxlc As MIXERLINECONTROLS, _
                                                        ByVal fdwControls As Long) As Long
' flags for mixerGetLineControls()
Public Const MIXER_GETLINECONTROLSF_ALL = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Public Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&

Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" _
            (ByVal hmxobj As Long, _
            pmxcd As MIXERCONTROLDETAILS, _
            ByVal fdwDetails As Long) As Long
' flags for mixerGetControlDetails()
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1&
Public Const MIXER_GETCONTROLDETAILSF_QUERYMASK = &HF&
' The mixerGetControlDetails function retrieves details about a single control associated
' with an audio line. the function uses the following parameters:
'     hmxobj-     a long value that is the handle to the mixer device object being queried.
'     pMxcd-      the variable defined as the MIXERCONTROLDETAILS user-defined type.
'     fdwDetails- Flags for retrieving control details. The following values are defined:
'                    MIXER_GETCONTROLDETAILSF_LISTTEXT-The paDetails member of the MIXERCONTROLDETAILS
'                       user-defined variable points to one or more MIXERCONTROLDETAILS_LISTTEXT user-defined
'                       variables to receive text labels for multiple-item controls. An application must get all list
'                       text items for a multiple-item control at once. This flag cannot be
'                       used with MIXERCONTROL_CONTROLTYPE_CUSTOM controls.
'                    MIXER_GETCONTROLDETAILSF_VALUE-Current values for a control are
'                       retrieved. The paDetails member of the MIXERCONTROLDETAILS user-defined
'                       variable points to one or more details appropriate for the control class.
'                    MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-
'                       The hmxobj parameter is the handle of a MIDI (Musical Instrument Digital
'                       Interface) input device. This handle must have been returned by the
'                       midiInOpen function.
'                    MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       hmxobj parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen function.
'                    MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The
'                       hmxobj parameter is a mixer device handle returned by the mixerOpen
'                       function. This flag is optional.
'                    MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       hmxobj parameter is a waveform-audio input handle returned by the
'                       waveInOpen function.
'                    MIXER_OBJECTF_HWAVEOUT ((MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       hmxobj parameter is a waveform-audio output handle returned by the
'                       waveOutOpen function.
'                    MIXER_OBJECTF_MIDIIN (&H40000000L)-The hmxobj parameter is the identifier
'                       of a MIDI input device. This identifier must be in the range of zero to
'                       one less than the number of devices returned by the midiInGetNumDevs
'                       function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier
'                       of a MIDI output device. This identifier must be in the range of zero
'                       to one less than the number of devices returned by the midiOutGetNumDevs
'                       function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is the identifier of a
'                       mixer device in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs function.

'Set Mixer Control Details call
Declare Function mixerSetControlDetails Lib "winmm.dll" _
                     (ByVal hmxobj As Long, _
                     pmxcd As MIXERCONTROLDETAILS, _
                     ByVal fdwDetails As Long) As Long


'set control details flags
Public Const MIXER_SETCONTROLDETAILSF_CUSTOM = &H1&
Public Const MIXER_SETCONTROLDETAILSF_QUERYMASK = &HF&
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&

Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                             ByVal dwBytes As Long) As Long
' The GlobalAlloc function allocates the specified number of bytes from the heap.
' Win32 memory management does not provide a separate local heap and global heap.
' This function is provided only for compatibility with 16-bit versions of Windows. The function
' uses the following parameters:
'     wFlags-     a long value that specifies how to allocate memory. If zero is specified,
'                 the default is allocate fixed memory. This value can be one or more of the
'                 following flags:
'                    GMEM_FIXED (&H0)- Allocates fixed memory. The return value is a pointer.
'                    GMEM_MOVEABLE (&H2)- Allocates movable memory. In Win32, memory blocks are
'                       never moved in physical memory, but they can be moved within the default .
'                       The return value is the handle of the memory object. To translate the
'                       heap handle into a pointer, use the GlobalLock function. This flag
'                       cannot be combined with the GMEM_FIXED flag.
'                    GPTR (GMEM_FIXED Or GMEM_ZEROINIT)-Combines the GMEM_FIXED and GMEM_ZEROINIT
'                       flags.
'                    GHND (GMEM_MOVEABLE Or GMEM_ZEROINIT)- Combines the GMEM_MOVEABLE and
'                       GMEM_ZEROINIT flags.
'                    GMEM_ZEROINIT (&H4)-Initializes memory contents to zero.
'     dwBytes-    Specifies the number of bytes to allocate. If this parameter is zero and
'                 the uFlags parameter specifies the GMEM_MOVEABLE flag, the function returns
'                 a handle to a memory object that is marked as discarded.

Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
' The GlobalLock function locks a global memory object and returns a pointer to the first
' byte of the object's memory block. This function is provided only for compatibility with
' 16-bit versions of Windows. The function requires a handle to the global memory object. This
' handle is returned by either the GlobalAlloc or GlobalReAlloc function.

Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
' The GlobalFree function frees the specified global memory object and invalidates its handle.
' This function is provided only for compatibility with 16-bit versions of Windows. The function
' requires a h andle to the global memory object. This handle is returned by either the
' GlobalAlloc or GlobalReAlloc function.
Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
' The CopyStructFromPtr and CopyPtrFromStruct functions are user-defined versions of the
' RtlMoveMemory function. RtlMoveMemory moves memory either forward or backward, aligned or
' unaligned, in 4-byte blocks, followed by any remaining bytes. The function requires the
' following parameters:
'     Destination-   Pointer to the starting address of the copied block's destination.
'     Source-        Pointer to the starting address of the block of memory to copy.
'     Length-        Specifies the size, in bytes, of the block of memory to copy.





Public Function LoWord(ByVal DWord As Long) As Integer
    If DWord And &H8000& Then
        LoWord = DWord Or &HFFFF0000
    Else
        LoWord = DWord And &HFFFF&
    End If
End Function
Public Function HiWord(ByVal DWord As Long) As Integer
    
    HiWord = (DWord And &HFFFF0000) \ &H10000

End Function

Public Function LoByte(ByVal w As Integer) As Byte
    LoByte = w And &HFF
End Function

Public Function HiByte(ByVal w As Integer) As Byte
    HiByte = (w And &HFF00&) \ 256
End Function

Public Function mmsysGetErrorString(ByVal error As vbMMSYSERRORS) As String
    
    Select Case error
        Case MIXERR_INVALCONTROL
            mmsysGetErrorString = "invalid control"
        Case MIXERR_INVALLINE
            mmsysGetErrorString = "invalid line"
        Case MIXERR_INVALVALUE
            mmsysGetErrorString = "invalid value"
        Case MMSYSERR_ALLOCATED
            mmsysGetErrorString = "device already allocated"
        Case MMSYSERR_BADDB
            mmsysGetErrorString = "bad registry database"
        Case MMSYSERR_BADDEVICEID
            mmsysGetErrorString = "device ID out of range"
        Case MMSYSERR_BADERRNUM
            mmsysGetErrorString = "error value out of range"
        Case MMSYSERR_DELETEERROR
            mmsysGetErrorString = "registry delete error"
        Case MMSYSERR_ERROR
            mmsysGetErrorString = "unspecified error"
        Case MMSYSERR_HANDLEBUSY
            mmsysGetErrorString = "handle being used"
        Case MMSYSERR_INVALFLAG
            mmsysGetErrorString = "invalid flag passed"
        Case MMSYSERR_INVALHANDLE
            mmsysGetErrorString = "device handle is invalid"
        Case MMSYSERR_INVALIDALIAS
            mmsysGetErrorString = "specified alias not found"
        Case MMSYSERR_INVALPARAM
            mmsysGetErrorString = "invalid parameter passed"
        Case MMSYSERR_KEYNOTFOUND
            mmsysGetErrorString = "registry key not found"
        Case MMSYSERR_NODRIVER
            mmsysGetErrorString = "no device driver present"
        Case MMSYSERR_NODRIVERCB
            mmsysGetErrorString = "driver does not call DriverCallback"
        Case MMSYSERR_NOMEM
            mmsysGetErrorString = "memory allocation error"
        Case MMSYSERR_NOTENABLED
            mmsysGetErrorString = "driver failed enable"
        Case MMSYSERR_NOTSUPPORTED
            mmsysGetErrorString = "function isn't supported"
        Case MMSYSERR_READERROR
            mmsysGetErrorString = "registry read error"
        Case MMSYSERR_VALNOTFOUND
            mmsysGetErrorString = "registry value not found"
        Case MMSYSERR_WRITEERROR
            mmsysGetErrorString = "registry write error"
        Case MMSYSERR_NOERROR
            mmsysGetErrorString = "no error"
        Case Else
            mmsysGetErrorString = "unknown error"
    End Select
    

End Function

Public Function GetControlTypeString(ByVal ctrlType As MIXERCONTROL_TYPE) As String
    Select Case ctrlType
        Case mcBASS_FADER
            GetControlTypeString = "BASS FADER CONTROL"
        Case mcBOOLEAN_METER
            GetControlTypeString = "BOOLEAN METER CONTROL"
        Case mcBOOLEAN_SWITCH
            GetControlTypeString = "BOOLEAN SWITCH CONTROL"
        Case mcBUTTON_SWITCH
            GetControlTypeString = "BUTTON SWITCH CONTROL"
        Case mcDECIBELS_NUMBER
            GetControlTypeString = "DECIBELS NUMBER CONTROL"
        Case mcEQUALIZER_FADER
            GetControlTypeString = "EQUALIZER FADER CONTROL"
        Case mcFADER_FADER
            GetControlTypeString = "FADER CONTROL"
        Case mcGENERIC_CUSTOM
            GetControlTypeString = "CUSTOM CONTROL"
        Case mcLOUDNESS_SWITCH
            GetControlTypeString = "LOUDNESS SWITCH CONTROL"
        Case mcMICROTIME_TIME
            GetControlTypeString = "MICROTIME TIME CONTROL"
        Case mcMILLITIME_TIME
            GetControlTypeString = "MILLITIME TIME CONTROL"
        Case mcMIXER_LIST
            GetControlTypeString = "MIXER LIST CONTROL"
        Case mcMONO_SWITCH
            GetControlTypeString = "MONO SWITCH CONTRL"
        Case mcMULTIPLESELECT_LIST
            GetControlTypeString = "MULTIPLESELECT LIST CONTROL"
        Case mcMUTE_SWITCH
            GetControlTypeString = "MUTE SWITCH CONTROL"
        Case mcMUX_LIST
            GetControlTypeString = "MUX LIST CONTROL"
        Case mcONOFF_SWITCH
            GetControlTypeString = "ONOFF SWITCH CONTROL"
        Case mcPAN_SLIDER
            GetControlTypeString = "PAN SLIDER CONTROL"
        Case mcPERCENT_NUMBER
            GetControlTypeString = "PERCENT NUMBER CONTROL"
        Case mcQSOUNDPAN_SLIDER
            GetControlTypeString = "QSOUND PAN SLIDER CONTROL"
        Case mcSIGNED_NUMBER
            GetControlTypeString = "SIGNED NUMBER CONTROL"
        Case mcSINGLESELECT_LIST
            GetControlTypeString = "SINGLESELECT LIST CONTROL"
        Case mcSLIDER_SLIDER
            GetControlTypeString = "SLIDER CONTROL"
        Case mcSTEREOENH_SWITCH
            GetControlTypeString = "STEREO ENHANCE SWITCH CONTROL"
        Case mcTREBLE_FADER
            GetControlTypeString = "TREBLE FADER CONTROL"
        Case mcUNSIGNED_NUMBER
            GetControlTypeString = "UNSIGNED NUMBER CONTROL"
        Case mcVOLUME_FADER
            GetControlTypeString = "VOLUME FADER CONTROL"
        Case mcPEAK_METER
            GetControlTypeString = "SIGNED PEAK METER"
        Case Else
            GetControlTypeString = "UNKNOWN"
    End Select
End Function

Public Function GetLineTypeString(ByVal lineType As MIXER_LINE_TYPE) As String
    Select Case lineType
        Case dstUNDEFINED
            GetLineTypeString = "UNDEFINED LINE TYPE"
        Case dstDIGITAL
            GetLineTypeString = "DIGITAL DESTINATION"
        Case dstline
            GetLineTypeString = "LINE LEVEL DESTINATION"
        Case dstMONITOR
            GetLineTypeString = "MONITOR DESTINATION"
        Case dstSPEAKERS
            GetLineTypeString = "ADJUSTABLE LEVEL SPEAKER DESTINATION"
        Case dstHEADPHONES
            GetLineTypeString = "ADJUSTABLE LEVEL HEADPHONES DESTINATION"
        Case dstTELEPHONE
            GetLineTypeString = "TELEPHONE LINE DESTINATION"
        Case dstWAVEIN
            GetLineTypeString = "WAVEIN (ADC) RECORDING DESTINATION"
        Case dstVOICEIN
            GetLineTypeString = "VOICE RECORDING DESTINATION"
        Case srcUNDEFINED
            GetLineTypeString = "NON-STANDARD SOURCE"
        Case srcDIGITAL
            GetLineTypeString = "DIGITAL SOURCE"
        Case srcLINE
            GetLineTypeString = "LINE LEVEL SOURCE"
        Case srcMICROPHONE
            GetLineTypeString = "MICROPHONE RECORDING SOURCE"
        Case srcSYNTHESIZER
            GetLineTypeString = "INTERNAL SYNTHESIZER SOURCE"
        Case srcCOMPACTDISC
            GetLineTypeString = "INTERNAL AUDIO CD SOURCE"
        Case srcTELEPHONE
            GetLineTypeString = "TELEPHONE LINE SOURCE"
        Case srcPCSPEAKER
            GetLineTypeString = "PC SPEAKER SOURCE"
        Case srcWAVEOUT
            GetLineTypeString = "WAVEOUT (DAC) SOURCE"
        Case srcAUXILIARY
            GetLineTypeString = "AUXILIARY AUDIO SOURCE"
        Case srcANALOG
            GetLineTypeString = "ANALOG SOURCE"
        Case Else
            GetLineTypeString = "UNKNOWN"
    End Select

End Function

Public Function GetTargetTypeString(ByVal targetType As TARGET_TYPE) As String
    Select Case targetType
        Case ttUNDEFINED
            GetTargetTypeString = "UNDEFINED"
        Case ttWAVEOUT
            GetTargetTypeString = "WAVE OUT"
        Case ttWAVEIN
            GetTargetTypeString = "WAVE IN"
        Case ttMIDIOUT
            GetTargetTypeString = "MIDI OUT"
        Case ttMIDIIN
            GetTargetTypeString = "MIDI IN"
        Case ttAUX
            GetTargetTypeString = "AUX AUDIO"
        Case Else
            GetTargetTypeString = "UNKNOWN"
    End Select

End Function



Public Sub DBPrintMIXERLINE(mxln As MIXERLINE, ByVal extradata As String)
   Debug.Print "* * * " & extradata & " * * *"
   With mxln
        Debug.Print "==============MIXERLINE============="
        Debug.Print "cbStruct: " & .cbStruct
        Debug.Print "dwDestination: " & .dwDestination
        Debug.Print "dwSource: " & .dwSource
        Debug.Print "dwLineID: " & .dwLineID
        Debug.Print "fdwLine: " & .fdwLine
        Debug.Print "dwUser: " & .dwUser
        Debug.Print "dwComponentType: " & .dwComponentType
        Debug.Print "cChannels: " & .cChannels
        Debug.Print "cConnections: " & .cConnections
        Debug.Print "cControls: " & .cControls
        Debug.Print "szShortName: " & .szShortName
        Debug.Print "szName: " & .szName
        Debug.Print "TARGET.dwType: " & .Target.dwType
        Debug.Print "TARGET.dwDeviceID: " & .Target.dwDeviceID
        Debug.Print "TARGET.wMid: " & .Target.wMid
        Debug.Print "TARGET.wPid: " & .Target.wPid
        Debug.Print "TARGET.vDriverVersion: " & .Target.vDriverVersion
        Debug.Print "szName: " & .szName
        Debug.Print "===================================="
   End With
End Sub

Public Sub DBPrintMIXERCONTROL(mxctl As MIXERCONTROL, ByVal extradata As String)
    Debug.Print "* * * " & extradata & " * * *"
    With mxctl
        Debug.Print "========mxctl========="
        Debug.Print "szShortName:" & .szShortName
        Debug.Print "szName: " & .szName
        'Debug.Print "lMinimum: " & .lMinimum
        'Debug.Print "lMaximum: " & .lMaximum
        'Debug.Print "cSteps: " & .cSteps
        Debug.Print "fdwControl: " & .fdwControl
        Debug.Print "dwControlID: " & .dwControlID
        Debug.Print "dwControlType: " & .dwControlType
        Debug.Print "cMultipleItems: " & .cMultipleItems
        Debug.Print "==============================="
    End With
End Sub
Public Sub DBPrintMIXERCONTROLDETAILS(mxcdtls As MIXERCONTROLDETAILS, ByVal extradata As String)
    Debug.Print "* * * " & extradata & " * * *"
    With mxcdtls
        Debug.Print "========mxcd========="
        Debug.Print "cbStruct:" & .cbStruct
        Debug.Print "dwControlID:" & .dwControlID
        Debug.Print "cChannels:" & .cChannels
        Debug.Print "Item:" & .Item
        Debug.Print "cbDetails:" & .cbDetails
        Debug.Print "paDetails:" & .paDetails
    End With
     
End Sub


