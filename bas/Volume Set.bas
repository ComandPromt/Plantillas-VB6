Attribute VB_Name = "MIXER"
'****************************************************************************
'* This constant holds the value of the Highest Custom volume setting.  The *
'* lowest value will always be zero.                                        *
'****************************************************************************
Public Const HIGHEST_VOLUME_SETTING = 12

'Put these into a module
'  device ID for aux device mapper
Public Const AUX_MAPPER = -1&
Public Const MAXPNAMELEN = 32

Type AUXCAPS
       wMid As Integer
       wPid As Integer
       vDriverVersion As Long
       szPname As String * MAXPNAMELEN
       wTechnology As Integer
       dwSupport As Long
End Type

'  flags for wTechnology field in AUXCAPS structure
Public Const AUXCAPS_CDAUDIO = 1  '  audio from internal CD-ROM drive
Public Const AUXCAPS_AUXIN = 2  '  audio from auxiliary input jacks

'  flags for dwSupport field in AUXCAPS structure
Public Const AUXCAPS_VOLUME = &H1         '  supports volume control
Public Const AUXCAPS_LRVOLUME = &H2         '  separate left-right volume control

Declare Function auxGetNumDevs Lib "winmm.dll" () As Long
Declare Function auxGetDevCaps Lib "winmm.dll" Alias "auxGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As AUXCAPS, ByVal uSize As Long) As Long

Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByRef lpdwVolume As Long) As Long
Declare Function auxOutMessage Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

'****************************************************************************
'* Possible Return values from auxGetVolume, auxSetVolume                   *
'****************************************************************************
Public Const MMSYSERR_NOERROR = 0
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)

'****************************************************************************
'* Use the CopyMemory function from the Windows API                         *
'****************************************************************************
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'****************************************************************************
'* Use this structure to break the Long into two Integers                   *
'****************************************************************************
Public Type VolumeSetting
    LeftVol As Integer
    RightVol As Integer
End Type

Sub lCrossFader()
'Vol1 = 100 - Slider1.Value ' Left
'Vol2 = 100 - Slider5.Value ' Right
'E = CrossFader.Value
'F = 100 - E
'If Check4.Value = 1 Then ' Half Fader Check
'    LVol = (F * Val(Vol1) / 100) * 2
'    RVol = (E * Val(Vol2) / 100) * 2
'    If LVol > (50 * Val(Vol1) / 100) * 2 Then
'        LVol = (50 * Val(Vol1) / 100) * 2
'    End If
'    If RVol > (50 * Val(Vol2) / 100) * 2 Then
'        RVol = (50 * Val(Vol2) / 100) * 2
'    End If
'Else
'    LVol = (F * Val(Vol1) / 100)
'    RVol = (E * Val(Vol2) / 100)
'End If
'Label1.Caption = "Fader: " + LTrim$(Str$(LVol)) + " x " + LTrim$(Str$(RVol))
'
End Sub


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


