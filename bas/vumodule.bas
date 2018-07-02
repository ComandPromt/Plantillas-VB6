Attribute VB_Name = "Module1"


   ' Global Memory Flag used by GlobalAlloc functin

Type WAVEHDR
' The WAVEHDR user-defined type defines the header used to identify a waveform-audio buffer.
   lpData As Long          ' Address of the waveform buffer.
   dwBufferLength As Long  ' Length, in bytes, of the buffer.
   dwBytesRecorded As Long ' When the header is used in input, this member specifies how much
                           ' data is in the buffer.

   dwUser As Long          ' User data.
   dwFlags As Long         ' Flags supplying information about the buffer. Set equal to zero.
   dwLoops As Long         ' Number of times to play the loop. Set equal to zero.
   lpNext As Long          ' Not used
   Reserved As Long        ' Not used
End Type

Type WAVEINCAPS
' The WAVEINCAPS user-defined variable describes the capabilities of a waveform-audio input
' device.
   wMid As Integer         ' Manufacturer identifier for the device driver for the
                           ' waveform-audio input device. Manufacturer identifiers
                           ' are defined in Manufacturer and Product Identifiers in
                           ' the Platform SDK product documentation.
   wPid As Integer         ' Product identifier for the waveform-audio input device.
                           ' Product identifiers are defined in Manufacturer and Product
                           ' Identifiers in the Platform SDK product documentation.
   vDriverVersion As Long  ' Version number of the device driver for the
                           ' waveform-audio input device. The high-order byte
                           ' is the major version number, and the low-order byte
                           ' is the minor version number.
   szPname As String * 32  ' Product name in a null-terminated string.
   dwFormats As Long       ' Standard formats that are supported. See the Platform
                           ' SDK product documentation for more information.
   wChannels As Integer    ' Number specifying whether the device supports
                           ' mono (1) or stereo (2) input.
End Type

Type WAVEFORMAT
' The WAVEFORMAT user-defined type describes the format of waveform-audio data. Only
' format information common to all waveform-audio data formats is included in this
' user-defined type.
   wFormatTag As Integer      ' Format type. Use the constant WAVE_FORMAT_PCM Waveform-audio data
                              ' to define the data as PCM.
   nChannels As Integer       ' Number of channels in the waveform-audio data. Mono data uses one
                              ' channel and stereo data uses two channels.
   nSamplesPerSec As Long     ' Sample rate, in samples per second.
   nAvgBytesPerSec As Long    ' Required average data transfer rate, in bytes per second. For
                              ' example, 16-bit stereo at 44.1 kHz has an average data rate of
                              ' 176,400 bytes per second (2 channels — 2 bytes per sample per
                              ' channel — 44,100 samples per second).
   nBlockAlign As Integer     ' Block alignment, in bytes. The block alignment is the minimum atomic unit of data. For PCM data, the block alignment is the number of bytes used by a single sample, including data for both channels if the data is stereo. For example, the block alignment for 16-bit stereo PCM is 4 bytes (2 channels — 2 bytes per sample).
   wBitsPerSample As Integer  ' For buffer estimation
   cbSize As Integer          ' Block size of the data.
End Type

Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, _
                                             ByVal uDeviceID As Long, _
                                             lpFormat As WAVEFORMAT, _
                                             ByVal dwCallback As Long, _
                                             ByVal dwInstance As Long, _
                                             ByVal dwFlags As Long) As Long
' The waveInOpen function opens the given waveform-audio input device for recording. The function
' uses the following parameters
'     lphWaveIn-  a long value that is the handle identifying the open waveform-audio input
'                 device. Use this handle to identify the device when calling other
'                 waveform-audio input functions. This parameter can be NULL if WAVE_FORMAT_QUERY
'                 is specified for fdwOpen.
'     uDeviceID-  a long value that identifies the waveform-audio input device to open.
'                 This parameter can be either a device identifier or a handle of an open
'                 waveform-audio input device.
'     lpFormat-   the WAVEFORMAT user-defined typed that identifies the desired format for
'                 recording waveform-audio data.
'     dwCallback- a long value that is an event handle, a handle to a window, or the identifier
'                 of a thread to be called during waveform-audio recording to process messages
'                 related to the progress of recording. If no callback function is required,
'                 this value can be zero. For more information on the callback function,
'                 see waveInProc.
'     dwCallback- a long value that is the user-instance data passed to the callback mechanism.
'                 This parameter is not used with the window callback mechanism.
'     dwFlags-    Flags for opening the device. The following values are defined:
'                 CALLBACK_EVENT (&H50000)-event handle.
'                 CALLBACK_FUNCTION (&H30000)-callback procedure address.
'                 CALLBACK_NULL (&H00000)-No callback mechanism. This is the default setting.
'                 CALLBACK_THREAD (&H20000)-thread identifier.
'                 CALLBACK_WINDOW (&H10000)-window handle.
'                 WAVE_FORMAT_DIRECT (&H8)-ACM driver does not perform conversions on the
'                                            audio data.
'                 WAVE_FORMAT_QUERY (&H1)-queries the device to determine whether it supports
'                                         the given format, but it does not open the device.
'                 WAVE_MAPPED (&H4)-The uDeviceID parameter specifies a waveform-audio device
'                                   to be mapped to by the wave mapper.

Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                      lpWaveInHdr As WAVEHDR, _
                                                      ByVal uSize As Long) As Long
' The waveInPrepareHeader function prepares a buffer for waveform-audio input. The function
' uses the following parameters:
'     hWaveIn-    a long value that is the handle of the waveform-audio input device.
'     lpWaveInHdr-the WAVEHDR user-defined type variable.
'     uSize-      the size in bytes of the WAVEHDR user-defined type variable. Use the
'                 results of the Len function for this parameter.


Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInReset function stops input on the given waveform-audio input device and resets
' the current position to zero. All pending buffers are marked as done and returned to
' the application. The function requires the handle to the waveform-audio input device.

Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInStart function starts input on the given waveform-audio input device. The function
' requires the handle of the waveform-audio input device.

Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInStop function stops waveform-audio input. The function requires the handle of
' the waveform-audio input device.

Declare Function waveInUnprepareHeader Lib "winmm.dll" _
                                          (ByVal hWaveIn As Long, _
                                          lpWaveInHdr As WAVEHDR, _
                                          ByVal uSize As Long) As Long
' The waveInUnprepareHeader function cleans up the preparation performed by the
' waveInPrepareHeader function. This function must be called after the device driver
' fills a buffer and returns it to the application. You must call this function before
' freeing the buffer. The function uses the following parameters:
'     hWaveIn-       a long value that is the handle of the waveform-audio input device.
'     lpWaveInHdr-   the variable typed as the WAVEHDR user-defined type identifying the
'                    buffer to be cleaned up.
'     uSize-         a long value that is the size in bytes, of the WAVEHDR varaible. Use
'                    the Len function with the WAVEHDR variable as the argument to get this
'                    value.

Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInClose function closes the given waveform-audio input device. The function
' requires the handle of the waveform-audio input device. If the function succeeds,
' the handle is no longer valid after this call.

Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" _
                  (ByVal uDeviceID As Long, _
                  lpCaps As WAVEINCAPS, _
                  ByVal uSize As Long) As Long
' This function retrieves the capabilities of a given waveform-audio input device. You can
' use this function to determine the number of waveform-audio input devices present in the
' system. If the value specified by the uDeviceID parameter is a device identifier,
' it can vary from zero to one less than the number of devices present. The function uses
' the following parameters
'     uDeviceID-     long value that identifies waveform-audio output device. This value can be
'                    either a device identifier or a handle of an open waveform-audio input device.
'     lpCaps-user-   defined variable containing information about the capabilities of the device.
'     uSize-         the size in bytes of the user-defined variable used as the lpCaps parameter.
'                    Use the Len function to get this value.

Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
' The waveInGetNumDevs function returns the number of waveform-audio input devices present in the system.

Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" _
                     (ByVal err As Long, _
                     ByVal lpText As String, _
                     ByVal uSize As Long) As Long
'The waveInGetErrorText function retrieves a textual description of the error identified by
' the given error number. The function uses the following parameters:
'     Err-     a long value that is the error number.
'     lpText-  a string variable that contains the textual error description.
'     uSize-   the size in characters of the lpText string variable.

Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                   lpWaveInHdr As WAVEHDR, _
                                                   ByVal uSize As Long) As Long
' The waveInAddBuffer function sends an input buffer to the given waveform-audio input device.
' The function uses the following parameters:
'     hWaveIn-       a long value that is the handle of the waveform-audio input device.
'     lpWaveInHdr-   the variable typed as the WAVEHDR user-defined type.
'     uSize-         a long value that is the size in bytes of the variable typed as the
'                    WAVEHDR user-defined variable. Use the Len function with the WAVEHDR
'                    variable as the argument to get this value.


Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32

Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&

Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)

Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)

Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
' The mixerClose function closes the specified mixer device. The function requires the
' handle of the mixer device. This handle must have been returned successfully by the
' mixerOpen function. If mixerClose is successful, the handle is no longer valid.

Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" _
            (ByVal hmxobj As Long, _
            pMxcd As MIXERCONTROLDETAILS, _
            ByVal fdwDetails As Long) As Long
' The mixerGetControlDetails function retrieves details about a single control associated
' with an audio line. the function uses the following parameters:
'     hmxobj-     a long value that is the handle to the mixer device object being queried.
'     pMxcd-      the variable defined as the MIXERCONTROLDETAILS user-defined type.
'     fdwDetails- Flags for retrieving control details. The following values are defined:
'                    MIXER_GETCONTROLDETAILSF_LISTTEXT-The paDetails member of the
'                       MIXERCONTROLDETAILS user-defined variable points to one or more
'                       MIXERCONTROLDETAILS_LISTTEXT user-defined variables to receive text
'                       labels for multiple-item controls. An application must get all list
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

Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" _
                  (ByVal uMxId As Long, _
                  ByVal pmxcaps As MIXERCAPS, _
                  ByVal cbmxcaps As Long) As Long
' The mixerGetDevCaps function queries a specified mixer device to determine its capabilities.
' The function uses the following parameters:
'     uMxId-      a long value that is the handle of an open mixer device.
'     pmxcaps-    a variable defined as the MIXERCAPS user-defined type to contain information
'                 about the capabilities of the device.
'     cbmxcaps-   a long value that is the size in bytes, of the variable defined as the
'                 MIXERCAPS user-defined type. Use the Len functions with the MIXERCAPS variable
'                 as the argument to get this value.

Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, _
                                             pumxID As Long, _
                                             ByVal fdwId As Long) As Long
' The mixerGetID function retrieves the device identifier for a mixer device associated
' with a specified device handle.The function uses the following parameters:
'     hmxobj-  a long value that is the handle of the audio mixer object to map to a
'              mixer device identifier.
'     pumxID-  the long value to contain the mixer device identifier. If no mixer device
'              is available for the hmxobj object, the value – 1 is placed in this location
'              and the MMSYSERR_NODRIVER error value is returned.
'     fdwId-   Flags for mapping the mixer object hmxobj. The following values are defined:
'                 MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs function.
'                 MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-
'                       the hmxobj parameter is the handle of a MIDI input device. This handle
'                       must have been returned by the midiInOpen function.
'                 MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       hmxobj parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen function.
'                 MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The hmxobj
'                       parameter is a mixer device handle returned by the mixerOpen function.
'                       This flag is optional.
'                 MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       hmxobj parameter is a waveform-audio input handle returned by the
'                       waveInOpen function.
'                 MIXER_OBJECTF_HWAVEOUT ((MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       hmxobj parameter is a waveform-audio output handle returned by the
'                       waveOutOpen function.
'                 MIXER_OBJECTF_MIDIIN (&H40000000L)-The hmxobj parameter is the identifier of
'                       a MIDI input device. This identifier must be in the range of zero to
'                       one less than the number of devices returned by the midiInGetNumDevs
'                       function.
'                 MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier of
'                       a MIDI output device. This identifier must be in the range of zero to
'                       one less than the number of devices returned by the midiOutGetNumDevs
'                       function.
'                 MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is the identifier of
'                       a mixer device in the range of zero to one less than the number of
'                       devices returned by the mixerGetNumDevs function. This flag is optional.
'                 MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs function.
'                 MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs function.

Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" _
                  (ByVal hmxobj As Long, _
                  pmxlc As MIXERLINECONTROLS, _
                  ByVal fdwControls As Long) As Long
' The mixerGetLineControls function retrieves one or more controls associated with an audio
' line. The function uses the following parameters:
'     hmxobj-        a long value that is the handle of the mixer device object that is being
'                    queried.
'     pmxlc-         the variable defined as the MIXERLINECONTROLS user-defined type used to
'                    reference one or more variables defined as theMIXERCONTROL user-defined
'                    types to be filled with information about the controls associated with
'                    an audio line. The cbStruct member of the MIXERLINECONTROLS variable
'                    must always be initialized to be the size, in bytes, of the
'                    MIXERLINECONTROLS variable.
'     fdwControls-   Flags for retrieving information about one or more controls associated w
'                    with an audio line. The following values are defined:
'                    MIXER_GETLINECONTROLSF_ALL-The pmxlc parameter references a list of
'                       MIXERCONTROL variables that will receive information on all controls
'                       associated with the audio line identified by the dwLineID member of
'                       the MIXERLINECONTROLS structure. The cControls member must be initialized
'                       to the number of controls associated with the line. This number is
'                       retrieved from the cControls member of the MIXERLINE structure returned
'                       by the mixerGetLineInfo function. The cbmxctrl member must be
'                       initialized to the size, in bytes, of a single MIXERCONTROL variable.
'                       The pamxctrl member must point to the first MIXERCONTROL variable to be
'                       filled. The dwControlID and dwControlType members are ignored for this
'                       query.
'                    MIXER_GETLINECONTROLSF_ONEBYID-The pmxlc parameter references a single
'                       MIXERCONTROL variable that will receive information on the control
'                       identified by the dwControlID member of the MIXERLINECONTROLS variable.
'                       The cControls member must be initialized to 1. The cbmxctrl member must
'                       be initialized to the size, in bytes, of a single MIXERCONTROL variable.
'                       The pamxctrl member must point to a MIXERCONTROL structure to be filled.
'                       The dwLineID and dwControlType members are ignored for this query. This
'                       query is usually used to refresh a control after receiving a
'                       MM_MIXM_CONTROL_CHANGE control change notification message by the
'                       user-defined callback (see mixerOpen).
'                    MIXER_GETLINECONTROLSF_ONEBYTYPE-The mixerGetLineControls function
'                       retrieves information about the first control of a specific class for
'                       the audio line that is being queried. The pmxlc parameter references a
'                       single MIXERCONTROL structure that will receive information about the
'                       specific control. The audio line is identified by the dwLineID member.
'                       The control class is specified in the dwControlType member of the
'                       MIXERLINECONTROLS variable. The dwControlID member is ignored for this
'                       query. This query can be used by an application to get information on
'                       a single control associated with a line. For example, you might want
'                       your application to use a peak meter only from a waveform-audio output
'                       line.
'                    MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-The
'                       hmxobj parameter is the handle of a MIDI input device. This handle must
'                       have been returned by the midiInOpen function.
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
'                    MIXER_OBJECTF_MIDIIN (&H40000000L)-The hmxobj parameter is the identifier of
'                       a MIDI input device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiInGetNumDevs function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier of
'                       a MIDI output device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiOutGetNumDevs function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is the identifier of a
'                       mixer device in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs function.

Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" _
                     (ByVal hmxobj As Long, _
                     pmxl As MIXERLINE, _
                     ByVal fdwInfo As Long) As Long
' The mixerGetLineInfo function retrieves information about a specific line of a mixer device.
' Uses the same parameters and constants as the mixerGetLineControls function.

Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
' The mixerGetNumDevs function retrieves the number of mixer devices present in the system.

Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, _
                                                ByVal uMsg As Long, _
                                                ByVal dwParam1 As Long, _
                                                ByVal dwParam2 As Long) As Long
' The mixerMessage function sends a custom mixer driver message directly to a mixer driver.
' The function uses the following parameters:
'     hmx-     a long value that is the handle of an open instance of a mixer device. This
'              value is the result of the mixerOpen function.
'     uMsg-    Custom mixer driver message to send to the mixer driver. This message must
'              be above or equal to the MXDM_USER constant.
'     dwParam1 and dwParam2-Arguments associated with the message being sent.

Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, _
                                             ByVal uMxId As Long, _
                                             ByVal dwCallback As Long, _
                                             ByVal dwInstance As Long, _
                                             ByVal fdwOpen As Long) As Long
' The mixerOpen function opens a specified mixer device and ensures that the device will
' not be removed until the application closes the handle. the function uses the following
' parameters:
'     phmx-       a long value that is the handle identifying the opened mixer device. Use
'                 this handle to identify the device when calling other audio mixer functions.
'                 This parameter cannot be NULL.
'     uMxId-      a long value that identifies the mixer device to open. Use a valid device
'                 identifier or any HMIXEROBJ (see the mixerGetID function for a description of
'                 mixer object handles). A "mapper" for audio mixer devices does not currently
'                 exist, so a mixer device identifier of – 1 is not valid.
'     dwCallback- Handle of a window called when the state of an audio line and/or control
'                 associated with the device being opened is changed. Specify zero for this
'                 parameter if no callback mechanism is to be used.
'     dwInstance- User instance data passed to the callback function. This parameter is not
'                 used with window callback functions.
'     fdwOpen-    Flags for opening the device. The following values are defined:
'                    CALLBACK_WINDOW-  The dwCallback parameter is assumed to be a window handle.
'                    MIXER_OBJECTF_AUX (&H50000000)-The uMxId parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-the
'                       uMxId parameter is the handle of a MIDI input device. This handle must
'                       have been returned by the midiInOpen function.
'                    MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       uMxId parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen function.
'                    MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The uMxId
'                       parameter is a mixer device handle returned by the mixerOpen function.
'                       This flag is optional.
'                    MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       uMxId parameter is a waveform-audio input handle returned by the
'                       waveInOpen function.
'                    MIXER_OBJECTF_HWAVEOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       uMxId parameter is a waveform-audio output handle returned by the
'                       waveOutOpen function.
'                    MIXER_OBJECTF_MIDIIN (&H40000000)-The uMxId parameter is the identifier of
'                       a MIDI input device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiInGetNumDevs function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The uMxId parameter is the identifier of
'                       a MIDI output device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiOutGetNumDevs function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The uMxId parameter is a mixer device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The uMxId parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The uMxId parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs function.

Declare Function mixerSetControlDetails Lib "winmm.dll" _
         (ByVal hmxobj As Long, _
         pMxcd As MIXERCONTROLDETAILS, _
         ByVal fdwDetails As Long) As Long
' The mixerSetControlDetails function sets properties of a single control associated with an
' audio line. The function uses the following parameters
'     hmxobj-        a long value that is the handle of the mixer device object for which
'                    properties are being set.
'     pMxcd-         the variable declares as the MIXERCONTROLDETAILS user-defined type.
'                    This variable references the control detail structures that contain the
'                    desired state for the control.
'     fdwDetails-    Flags for setting properties for a control. The following values are
'                    defined:
'                    MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-
'                       The hmxobj parameter is the handle of a MIDI input device. This handle
'                       must have been returned by the midiInOpen function.
'                    MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       hmxobj parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen function.
'                    MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The hmxobj
'                       parameter is a mixer device handle returned by the mixerOpen function.
'                       This flag is optional.
'                    MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       hmxobj parameter is a waveform-audio input handle returned by the
'                       waveInOpen function.
'                    MIXER_OBJECTF_HWAVEOUT ((MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       hmxobj parameter is a waveform-audio output handle returned by the
'                       waveOutOpen function.
'                    MIXER_OBJECTF_MIDIIN (&H40000000)-The hmxobj parameter is the identifier
'                       of a MIDI inputdevice. This identifier must be in the range of zero to
'                        one less than the number of devices returned by the midiInGetNumDevs
'                        function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier
'                       of a MIDI output device. This identifier must be in the range of zero
'                       to one less than the number of devices returned by the midiOutGetNumDevs
'                       function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is a mixer device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs function.
'                    MIXER_SETCONTROLDETAILSF_CUSTOM-A custom dialog box for the specified
'                       custom mixer control is displayed. The mixer device gathers the required
'                       information from the user and returns the data in the specified buffer.
'                       The handle for the owning window is specified in the hwndOwner member
'                       of the MIXERCONTROLDETAILS structure. (This handle can be set to NULL.)
'                       The application can then save the data from the dialog box and use it
'                       later to reset the control to the same state by using the
'                       MIXER_SETCONTROLDETAILSF_VALUE flag.
'                    MIXER_SETCONTROLDETAILSF_VALUE (&H00000000)-The current value(s) for a control
'                       are set. The paDetails member of the MIXERCONTROLDETAILS structure points
'                       to one or more mixer-control details structures of the appropriate class for
'                       the control.

Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
' The CopyStructFromPtr and CopyPtrFromStruct functions are user-defined versions of the
' RtlMoveMemory function. RtlMoveMemory moves memory either forward or backward, aligned or
' unaligned, in 4-byte blocks, followed by any remaining bytes. The function requires the
' following parameters:
'     Destination-   Pointer to the starting address of the copied block's destination.
'     Source-        Pointer to the starting address of the block of memory to copy.
'     Length-        Specifies the size, in bytes, of the block of memory to copy.

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

Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
' The GlobalLock function locks a global memory object and returns a pointer to the first
' byte of the object's memory block. This function is provided only for compatibility with
' 16-bit versions of Windows. The function requires a handle to the global memory object. This
' handle is returned by either the GlobalAlloc or GlobalReAlloc function.

Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
' The GlobalFree function frees the specified global memory object and invalidates its handle.
' This function is provided only for compatibility with 16-bit versions of Windows. The function
' requires a h andle to the global memory object. This handle is returned by either the
' GlobalAlloc or GlobalReAlloc function.

Type MIXERCAPS
' The MIXERCAPS user-defined type contains information about the capabilites of the mixer device.
   wMid As Integer                   '  manufacturer id
   wPid As Integer                   '  product id
   vDriverVersion As Long            '  version of the driver
   szPname As String * MAXPNAMELEN   '  product name
   fdwSupport As Long                '  misc. support bits
   cDestinations As Long             '  count of destinations
End Type



Type MIXERCONTROLDETAILS
' The MIXERCONTROLDETAILS user defined type refers to control-detail structures,
' retrieving or setting state information of an audio mixer control. All members of this
' user-defined type must be initialized before calling the mixerGetControlDetails and
' mixerSetControlDetails functions.
   cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
   dwControlID As Long    '  control id to get/set details on
   cChannels As Long      '  number of channels in paDetails array
   item As Long           '  hwndOwner or cMultipleItems
   cbDetails As Long      '  size of _one_ details_XX struct
   paDetails As Long      '  pointer to array of details_XX structs
End Type

Type MIXERCONTROLDETAILS_SIGNED
' The MIXERCONTROLDETAILS_SIGNED user-defined type retrieves and sets signed type control
' properties for an audio mixer control.
   lValue As Long
End Type

Type MIXERLINE
' The MIXERLINE user-defined type describes the state and metrics of an audio line.
   cbStruct As Long        ' Size of MIXERLINE structure
   dwDestination As Long   ' Zero based destination index
   dwSource As Long        ' Zero based source index (if source)
   dwLineID As Long        ' Unique line id for mixer device
   fdwLine As Long         ' State/information about line
   dwUser As Long          ' Driver specific information
   dwComponentType As Long ' Component type for this audio line.
   cChannels As Long       ' Maximum number of separate channels that can be
                           ' manipulated independently for the audio line.
   cConnections As Long    ' Number of connections that are associated with the
                           ' audio line.
   cControls As Long       ' Number of controls associated with the audio line.
   szShortName As String * MIXER_SHORT_NAME_CHARS  ' Short string that describes
                                                   ' the audio mixer line specified
                                                   ' in the dwLineID member.
   szName As String * MIXER_LONG_NAME_CHARS  ' String that describes the audio
                                             ' mixer line specified in the dwLineID
                                             ' member. This description should be
                                             ' appropriate as a complete description
                                             ' for the line.
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

Type MIXERLINECONTROLS
' The MIXERLINECONTROLS user-defined type contains information about the controls
' of an audio line.
   cbStruct As Long     ' size in Byte of MIXERLINECONTROLS
   dwLineID As Long     ' Line identifier for which controls are being queried.
   dwControl As Long    ' Control identifier of the desired control
   cControls As Long    ' Number of MIXERCONTROL structure elements to retrieve.
   cbmxctrl As Long     ' Size, in bytes, of a single MIXERCONTROL structure.
   pamxctrl As Long     ' Address of one or more MIXERCONTROL structures to receive
                        '  the properties of the requested audio line controls.
End Type

Public i As Integer
Public j As Integer
Public rc As Long
Public msg As String * 200
Public hWaveIn As Long
Public format As WAVEFORMAT

Public Const NUM_BUFFERS = 2
Public Const BUFFER_SIZE = 8192
Public Const DEVICEID = 0
Public hmem(NUM_BUFFERS) As Long
Public inHdr(NUM_BUFFERS) As WAVEHDR

Public fRecording As Boolean

Function GetControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Boolean
' This function attempts to obtain a mixer control. Returns True if successful.

   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType
   
   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
   If (MMSYSERR_NOERROR = rc) Then
      mxlc.cbStruct = Len(mxlc)
      mxlc.dwLineID = mxl.dwLineID
      mxlc.dwControl = ctrlType
      mxlc.cControls = 1
      mxlc.cbmxctrl = Len(mxc)
      
      ' Allocate a buffer for the control
      'hmem = GlobalAlloc(&H40, Len(mxc))
      hmem = GlobalAlloc(GMEM_FIXED, Len(mxc))
      mxlc.pamxctrl = GlobalLock(hmem)
      mxc.cbStruct = Len(mxc)
      
      ' Get the control
      rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
            
      If (MMSYSERR_NOERROR = rc) Then
         GetControl = True
         
         ' Copy the control into the destination structure
         CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
      Else
         GetControl = False
      End If
      GlobalFree (hmem)
      Exit Function
   End If
   
   GetControl = False
End Function

' Function to process the wave recording notifications.
Sub waveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
   If (uMsg = MM_WIM_DATA) Then
      If fRecording Then
         rc = waveInAddBuffer(hwi, hdr, Len(hdr))
      End If
   End If
End Sub

' This function starts recording from the soundcard. The soundcard must be recording in order to
' monitor the input level. Without starting the recording from this application, input level
' can still be monitored if another application is recording audio
Function StartInput() As Boolean

    If fRecording Then
        StartInput = True
        Exit Function
    End If
    
    format.wFormatTag = 1
    format.nChannels = 1
    format.wBitsPerSample = 8
    format.nSamplesPerSec = 8000
    format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8
    format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign
    format.cbSize = 0
    
    For i = 0 To NUM_BUFFERS - 1
        hmem(i) = GlobalAlloc(&H40, BUFFER_SIZE)
        inHdr(i).lpData = GlobalLock(hmem(i))
        inHdr(i).dwBufferLength = BUFFER_SIZE
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
    Next

    rc = waveInOpen(hWaveIn, DEVICEID, format, 0, 0, 0)
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg
        StartInput = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
End Function

' Stop receiving audio input on the soundcard
Sub StopInput()

    fRecording = False
    waveInReset hWaveIn
    waveInStop hWaveIn
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    waveInClose hWaveIn
End Sub
