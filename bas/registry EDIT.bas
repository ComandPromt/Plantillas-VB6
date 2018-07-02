Attribute VB_Name = "Module1"
Option Explicit

Public Const NO_ERROR = &H0

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
      (ByVal hKey As Long, _
      ByVal lpSubKey As String, _
      ByVal ulOptions As Long, _
      ByVal samDesired As Long, _
      phkResult As Long) As Long
'////////////////////////////////////////////////////////////////////////////////////////////
' This Windows API function opens the specified key in the registry. This function requires
' the following parameters:
'     hKey-       a long value that identifies a currently open key or any of the following
'                 predefined reserved registry keys values:
'                    HKEY_CLASSES_ROOT
'                    HKEY_CURRENT_CONFIG
'                    HKEY_CURRENT_USER
'                    HKEY_LOCAL_MACHINE
'                    HKEY_USERS
'                    Windows NT only: HKEY_PERFORMANCE_DATA
'                    Windows 95 only: HKEY_DYN_DATA
'     lpSubKey-   a null-terminated string containing the name of the subkey to open.
'                 If this parameter is NULL or a pointer to an empty string, the function
'                 will open a new handle of the key identified by the hKey parameter. In
'                 this case, the function will not close the handles previously opened.
'     ulOptions-  set to zero.
'     samDesired- a long value that specifies an access mask that describes the desired
'                 security access for the new key. See the Platform SDK documentation or
'                 the winnt.h header file shipped with Microsoft Visual C++ for the
'                 appropriate values.
'     phkResult-  a long variable that receives the handle of the opened key. When you no
'                 longer need the returned handle, call the RegCloseKey function to close it.
'
' Unlike the RegCreateKeyEx function, the RegOpenKeyEx function does not create the specified
' key if the key does not exist in the registry.
'////////////////////////////////////////////////////////////////////////////////////////////
      
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
      (ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal lpReserved As Long, _
      lpType As Long, _
      lpData As Any, _
      lpcbData As Long) As Long
'////////////////////////////////////////////////////////////////////////////////////////////
' This Windows API function retrieves the type and data for a specified value name associated
' with an open registry key. The function has the following parameters:
'     hKey-          a long value that identifies a currently open key or any of the following
'                    predefined reserved handle values:
'                       HKEY_CLASSES_ROOT
'                       HKEY_CURRENT_CONFIG
'                       HKEY_CURRENT_USER
'                       HKEY_LOCAL_MACHINE
'                       HKEY_USERS
'                       Windows NT only: HKEY_PERFORMANCE_DATA
'                       Windows 95 only: HKEY_DYN_DATA
'     lpValueName-   a null-terminated string containing the name of the value to query.
'                    If this parameter is an empty string, the function retrieves the
'                    type and data for the key's unnamed or default value, if any.
'     lpReserved-    set to null.
'     lpType-        a long variable that receives the type of data associated with the
'                    specified value.
'     lpData-        a long variable to receive the value's data. This parameter can be
'                    NULL if the data is not required.
'     lpcbData-      the size of the string buffer. Use the Len function to get this value.
'////////////////////////////////////////////////////////////////////////////////////////////
      
Declare Function RegQueryValueString Lib "advapi32.dll" Alias "RegQueryValueExA" _
      (ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal lpReserved As Long, _
      lpType As Long, _
      ByVal buf As String, _
      lpcbData As Long) As Long
'/////////////////////////////////////////////////////////////////////////////////////////////
' This Windows API function is similar to the previous function except one of the parameters
' accepts a string (buf) while the previous functions accepts any variable type (lpData).
'////////////////////////////////////////////////////////////////////////////////////////////
      
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'/////////////////////////////////////////////////////////////////////////////////////////////
' This Windows API function closes the handle of the specified open key.
'/////////////////////////////////////////////////////////////////////////////////////////////

Public Const HKEY_CURRENT_USER = &H80000001
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Public Const MCI_SET = &H80D
Public Const MCI_WAVE_OUTPUT = &H800000
Public Const MCI_WAVE_INPUT = &H400000
Public Const MCI_WAVE_SET_AVGBYTESPERSEC = &H80000
Public Const MCI_WAVE_SET_BITSPERSAMPLE = &H200000
Public Const MCI_WAVE_SET_BLOCKALIGN = &H100000
Public Const MCI_WAVE_SET_CHANNELS = &H20000
Public Const MCI_WAVE_SET_FORMATTAG = &H10000
Public Const MCI_WAVE_SET_SAMPLESPERSEC = &H40000

Public Const MAXPNAMELEN = 32

Type WAVEFORMAT
'//////////////////////////////////////////////////////////////////////////////////////
' The WAVEFORMAT user-defined variable describes the format of waveform-audio data.
' Only format information common to all waveform-audio data formats is this variable.
'//////////////////////////////////////////////////////////////////////////////////////
   wFormatTag As Integer      ' Format type.
   
   nChannels As Integer       ' Number of channels in the waveform-audio data.
                              ' Mono data uses one channel and stereo data uses two channels.

   nSamplesPerSec As Long     ' Sample rate in samples per second.

   nAvgBytesPerSec As Long    ' Required average data transfer rate in bytes per second.
                              ' For example, 16-bit stereo at 44.1 kHz has an average data rate
                              ' of 176,400 bytes per second (2 channels — 2 bytes per sample
                              ' per channel — 44,100 samples per second).
                              
   nBlockAlign As Integer     ' Block alignment in bytes. The block alignment is the minimum
                              ' atomic unit of data. For PCM data, the block alignment is the
                              ' number of bytes used by a single sample, including data for
                              ' both channels if the data is stereo. For example, the block
                              ' alignment for 16-bit stereo PCM is 4 bytes (2 channels — 2
                              ' bytes per sample).

   wBitsPerSample As Integer  ' Number of bits per sample
   
   cbSize As Integer
End Type

Type MCI_WAVE_SET_PARMS
'//////////////////////////////////////////////////////////////////////////////////////////
' The MCI_WAVE_SET_PARMS structure user-defined variable contains information for
' the MCI_SET command for waveform-audio devices.
'//////////////////////////////////////////////////////////////////////////////////////////
   dwCallback As Long         ' Window handle used for the MCI_NOTIFY flag.
   
   dwTimeFormat As Long       ' Time format of the device.
   
   dwAudio As Long            ' Channel number for audio output. Typically used
                              ' when turning a channel on or off.
                               
   wInput As Long             ' Audio input channel.
   
   wOutput As Long            ' Output device to use. For example, this value could be
                              ' 2 if a system had two installed sound cards.
                               
   wFormatTag As Integer      ' Format of the waveform-audio data. See the Platform SDK
                              ' product documentation for more information about formats.
                               
   wReserved2 As Integer      ' Reserved
   nChannels As Integer       ' 1 for Mono or 2 for stereo.
   wReserved3 As Integer      ' Reserved.
   nSamplesPerSec As Long     ' Samples per second.
   nAvgBytesPerSec As Long    ' Sample rate in bytes per second.
   nBlockAlign As Integer     ' Block alignment of the data.
   wReserved4 As Integer      ' Reserved.
   wBitsPerSample As Integer  ' Bits per sample.
   wReserved5 As Integer      ' Reserved
End Type

Type WAVEOUTCAPS
'///////////////////////////////////////////////////////////////////////////////////////////
' The WAVEOUTCAPS user-defined variable describes the capabilities of a waveform-audio
' output device.
'///////////////////////////////////////////////////////////////////////////////////////////
   wMid As Integer                  ' Manufacturer identifier for the device driver for
                                    ' the device. Manufacturer identifiers are defined
                                    ' in Manufacturer and Product Identifiers in the
                                    ' Platform SDK product documentation.

   wPid As Integer                  ' Product identifier for the device. Product
                                    ' identifiers are defined in Manufacturer and Product
                                    ' Identifiers in the Platform SDK product documentation.

   vDriverVersion As Long           ' Version number of the device driver for the device.
                                    ' The high-order byte is the major version number, and
                                    ' the low-order byte is the minor version number.

   szPname As String * MAXPNAMELEN  ' Product name in a null-terminated string.

   dwFormats As Long                ' Standard formats that are supported. See the Platform
                                    ' SDK product documentation for more information.
                                    
   wChannels As Integer             ' 1 if the device supports mono or 2 if the device supports
                                    ' stereo output.

   dwSupport As Long                ' Optional functionality supported by the device. See the
                                    ' Platform SDK product documentation for more information.
End Type

Type WAVEINCAPS
'////////////////////////////////////////////////////////////////////////////////////////////
' The WAVEINCAPS user-defined variable describes the capabilities of a waveform-audio input
' device.
'////////////////////////////////////////////////////////////////////////////////////////////
   wMid As Integer                  ' Manufacturer identifier for the device driver for the
                                    ' waveform-audio input device. Manufacturer identifiers
                                    ' are defined in Manufacturer and Product Identifiers in
                                    ' the Platform SDK product documentation.
   
   wPid As Integer                  ' Product identifier for the waveform-audio input device.
                                    ' Product identifiers are defined in Manufacturer and Product
                                    ' Identifiers in the Platform SDK product documentation.
   
   vDriverVersion As Long           ' Version number of the device driver for the
                                    ' waveform-audio input device. The high-order byte
                                    ' is the major version number, and the low-order byte
                                    ' is the minor version number.
   
   szPname As String * MAXPNAMELEN  ' Product name in a null-terminated string.
   dwFormats As Long                ' Standard formats that are supported. See the Platform
                                    ' SDK product documentation for more information.
                                     
   wChannels As Integer              ' Number specifying whether the device supports
                                     ' mono (1) or stereo (2) input.
End Type

Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" _
   (ByVal dwError As Long, _
   ByVal lpstrBuffer As String, _
   ByVal uLength As Long) As Long
'/////////////////////////////////////////////////////////////////////////////////////////////
' This Windows API function retrieves a string that describes the specified MCI error code.
' The unction requires the following parameters
'
'     dwError-       the long return value from the mciSendString function.
'     lpstrBuffer-   a string variable that is the buffer that receives a null-terminated
'                    string describing the specified error. Show this variable in a
'                    message box to determine the error.
'     uLength-       a long value that represents the length in characters of the buffer.
'                    Use the results of the Len function on the lpstrBuffer string variable
'                    for this value.
'
' The function returns TRUE if successful or FALSE if the error code is not known.
'/////////////////////////////////////////////////////////////////////////////////////////////

Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" _
   (ByVal wDeviceID As Long, _
   ByVal uMessage As Long, _
   ByVal dwParam1 As Long, _
   ByRef dwParam2 As Any) As Long
'/////////////////////////////////////////////////////////////////////////////////////////////
' This Wincows API function sends a command message to the specified MCI device. The function
' uses the following parameters:
'     wDeviceID-  a long value identifying the MCI device to receive the command message.
'                 This parameter is not used with the MCI_OPEN command message.
'     uMessage-   a long value representing the command message. See Command Messages in
'                 the Platform SDK documentation for more information.
'     dwParam1-   a long value that represents the flags for the command message.
'     dwParam2-   parameters for the command message.
'/////////////////////////////////////////////////////////////////////////////////////////////

Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" _
   (ByVal uDeviceID As Long, _
   lpCaps As WAVEOUTCAPS, _
   ByVal uSize As Long) As Long
'/////////////////////////////////////////////////////////////////////////////////////////////
' This Windows API function retrieves the capabilities of a given waveform-audio output device.
' The function uses the following parameters:
'     uDeviceID-  a long value identifying the waveform-audio output device.
'     lpCaps-     the user defined variable with information about the capabilities of the device.
'     uSize-      the size of the user-defined variable used to store information about the
'                 capabilities of the device.
'////////////////////////////////////////////////////////////////////////////////////////////

Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
'////////////////////////////////////////////////////////////////////////////////////////////
' This Windows API function returns the number of waveform-audio input devices present in the
' system.
'////////////////////////////////////////////////////////////////////////////////////////////

Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" _
   (ByVal uDeviceID As Long, _
   lpCaps As WAVEINCAPS, _
   ByVal uSize As Long) As Long
'/////////////////////////////////////////////////////////////////////////////////////////////
' This Windows API function retrieves the capabilities of a given waveform-audio input device.
' You can use this function to determine the number of waveform-audio input devices present in
' the system. If the value specified by the uDeviceID parameter is a device identifier, it can
' vary from zero to one less than the number of devices present. The function uses the
' following parameters:
'     uDeviceID-     long value that identifies waveform-audio output device. This value can
'                    be either a device identifier or a handle of an open waveform-audio
'                    input device.
'     lpCaps-        user-defined variable containing information about the capabilities of
'                    the device.
'     uSize-         the size in bytes of the user-defined variable used as the lpCaps parameter.
'                    Use the Len function to get this value.
'//////////////////////////////////////////////////////////////////////////////////////////////

Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
'//////////////////////////////////////////////////////////////////////////////////////////////
' This Windows API function returns the number of waveform-audio input devices present in the
' system.
'//////////////////////////////////////////////////////////////////////////////////////////////
