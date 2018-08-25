Attribute VB_Name = "Declarations"

Public Const MMSYSERR_NOERROR = 0
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)
Public Const MAXPNAMELEN = 32
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&

Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
               
Public Const MIXERLINE_COMPONENTTYPE_SRC_TREBLE = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
                             
Public Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
               
Public Const MIXERLINE_COMPONENTTYPE_SRC_TADIN = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
               
Public Const MIXERLINE_COMPONENTTYPE_SRC_MIDI = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
               
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
               (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
               
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)

Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)

Public Const MIXERLINE_COMPONENTTYPE_SRC_I25 = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)

Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000

Public Const MIXERCONTROL_CONTROLTYPE_FADER = _
               (MIXERCONTROL_CT_CLASS_FADER Or _
               MIXERCONTROL_CT_UNITS_UNSIGNED)

Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = _
               (MIXERCONTROL_CONTROLTYPE_FADER + 1)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)


'****************************************************************************
'* Use this structure to break the Long into two Integers                   *
'*********************************************************************
Public Type VolumeSetting
    LeftVol As Integer
    RightVol As Integer
End Type


'  flags for wTechnology field in AUXCAPS structure
Public Const AUXCAPS_CDAUDIO = 1  '  audio from internal CD-ROM drive
Public Const AUXCAPS_AUXIN = 2  '  audio from auxiliary input jacks

'  flags for dwSupport field in AUXCAPS structure
Public Const AUXCAPS_VOLUME = &H1         '  supports volume control
Public Const AUXCAPS_LRVOLUME = &H2         '  separate left-right volume control

Declare Function auxGetNumDevs Lib "winmm.dll" () As Long
Declare Function auxGetDevCaps Lib "winmm.dll" Alias "auxGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As AUXCAPS, ByVal uSize As Long) As Long
Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, ByRef pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByRef lpdwVolume As Long) As Long
Declare Function auxOutMessage Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
  
Declare Function mixerGetControlDetails Lib "winmm.dll" _
               Alias "mixerGetControlDetailsA" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
   

   
Declare Function mixerGetID Lib "winmm.dll" _
               (ByVal hmxobj As Long, _
               pumxID As Long, _
               ByVal fdwId As Long) As Long
               
Declare Function mixerGetLineControls Lib "winmm.dll" _
               Alias "mixerGetLineControlsA" _
               (ByVal hmxobj As Long, _
               pmxlc As MIXERLINECONTROLS, _
               ByVal fdwControls As Long) As Long
               
Declare Function mixerGetLineInfo Lib "winmm.dll" _
               Alias "mixerGetLineInfoA" _
               (ByVal hmxobj As Long, _
               pmxl As MIXERLINE, _
               ByVal fdwInfo As Long) As Long
               


Declare Function mixerMessage Lib "winmm.dll" _
               (ByVal hmx As Long, _
               ByVal uMsg As Long, _
               ByVal dwParam1 As Long, _
               ByVal dwParam2 As Long) As Long
                              
Declare Function mixerSetControlDetails Lib "winmm.dll" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
               
Declare Sub CopyStructFromPtr Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (struct As Any, _
               ByVal ptr As Long, ByVal cb As Long)
               
Declare Sub CopyPtrFromStruct Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (ByVal ptr As Long, _
               struct As Any, _
               ByVal cb As Long)
               
'****************************************************************************
'* Use the CopyMemory function from the Windows API                         *
'****************************************************************************
              
Declare Function GlobalAlloc Lib "kernel32" _
               (ByVal wFlags As Long, _
               ByVal dwBytes As Long) As Long
               
Declare Function GlobalLock Lib "kernel32" _
               (ByVal hmem As Long) As Long
               
Declare Function GlobalFree Lib "kernel32" _
               (ByVal hmem As Long) As Long

Type AUXCAPS
       wMid As Integer
       wPid As Integer
       vDriverVersion As Long
       szPname As String * MAXPNAMELEN
       wTechnology As Integer
       dwSupport As Long
End Type


Type MIXERCAPS
   wMid As Integer                   '  manufacturer id
   wPid As Integer                   '  product id
   vDriverVersion As Long            '  version of the driver
   szPname As String * MAXPNAMELEN   '  product name
   fdwSupport As Long                '  misc. support bits
   cDestinations As Long             '  count of destinations
End Type

Type MIXERCONTROL
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
   cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
   dwControlID As Long    '  control id to get/set details on
   cChannels As Long      '  number of channels in paDetails array
   item As Long           '  hwndOwner or cMultipleItems
   cbDetails As Long      '  size of _one_ details_XX struct
   paDetails As Long      '  pointer to array of details_XX structs
End Type

Type MIXERCONTROLDETAILS_UNSIGNED
   dwValue As Long        '  value of the control (volume level)
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
   dwType As Long
   dwDeviceID As Long
   wMid  As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
End Type

Type MIXERLINECONTROLS
   cbStruct As Long       '  size in Byte of MIXERLINECONTROLS
   dwLineID As Long       '  line id (from MIXERLINE.dwLineID)
                          '  MIXER_GETLINECONTROLSF_ONEBYID or
   dwControl As Long      '  MIXER_GETLINECONTROLSF_ONEBYTYPE
   cControls As Long      '  count of controls pmxctrl points to
   cbmxctrl As Long       '  size in Byte of _one_ MIXERCONTROL
   pamxctrl As Long       '  pointer to first MIXERCONTROL array
End Type

Function GetMixerControl(ByVal hmixer As Long, _
                        ByVal componentType As Long, _
                        ByVal ctrlType As Long, _
                        ByRef mxc As MIXERCONTROL) As Boolean
                        
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
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       ' Get the control
       rc = mixerGetLineControls(hmixer, _
                                 mxlc, _
                                 MIXER_GETLINECONTROLSF_ONEBYTYPE)
       If (MMSYSERR_NOERROR = rc) Then
           GetMixerControl = True
           ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       Else
           GetMixerControl = False
       End If
       GlobalFree (hmem)
       Exit Function
   End If
   GetMixerControl = False
End Function

Function SetVolumeControl(ByVal hmixer As Long, _
                        mxc As MIXERCONTROL, _
                        ByVal Volume As Long) As Boolean
'This function sets the value for a volume control. Returns True if successful
   Dim mxcd As MIXERCONTROLDETAILS
   Dim Vol As MIXERCONTROLDETAILS_UNSIGNED
                       
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.item = 0
   mxcd.cbDetails = Len(Vol)
   ' Allocate memory for the control value buffer
   ' Len(Vol) = Number of Bytes
   ' &H40     = GMEM_ZEROINIT = Initialize to zero
   ' hmem     = Handle to the newly created memory object (Buffer)
   hmem = GlobalAlloc(&H40, Len(Vol))
   ' Lock the memory object (buffer) and return a pointer to the first byte.
   mxcd.paDetails = GlobalLock(hmem)
   Vol.dwValue = Volume
   ' Copy the data into the control value buffer
   CopyPtrFromStruct mxcd.paDetails, Vol, Len(Vol)
   ' Set the control value
   rc = mixerSetControlDetails(hmixer, _
                              mxcd, _
                              MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetVolumeControl = True
   Else
       SetVolumeControl = False
   End If
End Function

Function GetVolumeControlValue(ByVal hmixer As Long, mxc As MIXERCONTROL) As Long
'This function Gets the value for a volume control. Returns True if successful
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Vol As MIXERCONTROLDETAILS_UNSIGNED

    mxcd.cbStruct = Len(mxcd)
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cChannels = 1
    mxcd.item = 0
    mxcd.cbDetails = Len(Vol)
    mxcd.paDetails = 0
    ' Allocate memory for the control value buffer
    ' Len(Vol) = Number of Bytes
    ' &H40     = GMEM_ZEROINIT = Initialize to zero
    ' hmem     = Handle to the newly created memory object (Buffer)
    hmem = GlobalAlloc(&H40, Len(Vol))
    ' Lock the memory object (buffer) and return a pointer to the first byte.
    mxcd.paDetails = GlobalLock(hmem)
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    ' Copy the data from the control value buffer
    CopyStructFromPtr Vol, mxcd.paDetails, Len(Vol)
    GlobalFree (hmem)
    If (MMSYSERR_NOERROR = rc) Then
       GetVolumeControlValue = Vol.dwValue
    Else
        GetVolumeControlValue = -1
    End If
End Function
Public Function lSetVolume(ByRef lLeftVol As Long, ByRef lRightVol As Long, lDeviceID As Long) As Long
'****************************************************************************
'* This function sets the current Windows volume settings to the specified  *
'* device using two Custom numbers from 0 to HIGHEST_VOLUME_SETTING for the *
'* right and left volume settings.                                          *
'*                                                                          *
'* The return value of this function is the Return value of the auxGetVolume*
'* Windows API call.                                                        *
'****************************************************************************

    Dim bReturnValue As Boolean                     ' Return Value from Function
    Dim Volume As VolumeSetting                     ' Type structure used to convert a long to/from
                                                    ' two Integers.
                                                    
    Dim lAPIReturnVal As Long                       ' Return value from API Call
    Dim lBothVolumes As Long                        ' The API passed value of the Combined Volumes
    
    
'****************************************************************************
'* Calculate the Integers                                                   *
'****************************************************************************
    Volume.LeftVol = nSigned(lLeftVol * 65535 / HIGHEST_VOLUME_SETTING)
    Volume.RightVol = nSigned(lRightVol * 65535 / HIGHEST_VOLUME_SETTING)
    
'****************************************************************************
'* Combine the Integers into a Long to be Passed to the API                 *
'****************************************************************************
    lDataLen = Len(Volume)
    CopyMemory lBothVolumes, Volume.LeftVol, lDataLen

'****************************************************************************
'* Set the Value to the API                                               *
'****************************************************************************
    lAPIReturnVal = auxSetVolume(lDeviceID, lBothVolumes)
    lSetVolume = lAPIReturnVal

End Function


Public Function lGetVolume(ByRef lLeftVol As Long, ByRef lRightVol As Long, lDeviceID As Long) As Long
'****************************************************************************
'* This function reads the current Windows volume settings from the         *
'* specified device, and returns two numbers from 0 to                      *
'* HIGHEST_VOLUME_SETTING for the right and left volume settings.           *
'*                                                                          *
'* The return value of this function is the Return value of the auxGetVolume*
'* Windows API call.                                                        *
'****************************************************************************

    Dim bReturnValue As Boolean                     ' Return Value from Function
    Dim Volume As VolumeSetting                     ' Type structure used to convert a long to/from
                                                    ' two Integers.
    Dim lAPIReturnVal As Long                       ' Return value from API Call
    Dim lBothVolumes As Long                        ' The API Return of the Combined Volumes
    
'****************************************************************************
'* Get the Value from the API                                               *
'****************************************************************************
    lAPIReturnVal = auxGetVolume(lDeviceID, lBothVolumes)
    
'****************************************************************************
'* Split the Long value returned from the API into to Integers              *
'****************************************************************************
    lDataLen = Len(Volume)
    CopyMemory Volume.LeftVol, lBothVolumes, lDataLen
    
'****************************************************************************
'* Calculate the Return Values.                                             *
'****************************************************************************
    lLeftVol = HIGHEST_VOLUME_SETTING * lUnsigned(Volume.LeftVol) / 65535
    lRightVol = HIGHEST_VOLUME_SETTING * lUnsigned(Volume.RightVol) / 65535

    lGetVolume = lAPIReturnVal
End Function

Public Function nSigned(ByVal lUnsignedInt As Long) As Integer
    Dim nReturnVal As Integer                          ' Return value from Function
    
    If lUnsignedInt > 65535 Or lUnsignedInt < 0 Then
        MsgBox "Error in conversion from Unsigned to nSigned Integer"
        nSignedInt = 0
        Exit Function
    End If

    If lUnsignedInt > 32767 Then
        nReturnVal = lUnsignedInt - 65536
    Else
        nReturnVal = lUnsignedInt
    End If
    
    nSigned = nReturnVal

End Function

Public Function lUnsigned(ByVal nSignedInt As Integer) As Long
    Dim lReturnVal As Long                          ' Return value from Function
    
    If nSignedInt < 0 Then
        lReturnVal = nSignedInt + 65536
    Else
        lReturnVal = nSignedInt
    End If
    
    If lReturnVal > 65535 Or lReturnVal < 0 Then
        MsgBox "Error in conversion from nSigned to Unsigned Integer"
        lReturnVal = 0
    End If
    
    lUnsigned = lReturnVal
End Function



