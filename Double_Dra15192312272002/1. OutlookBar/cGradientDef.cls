VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGradientDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
'   You are free to use this source as long as this copyright message
'     appears on your program's "About" dialog:
'
'   Outlook Bar Project
'   Copyright (c) 2002 Vlad Vissoultchev (wqweto@myrealbox.com)
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "cGradientDef"

'=========================================================================
' Events
'=========================================================================

'Purpose: Occurs when the contents of the current <b>cGradientDef</b> have changed.
Event Changed()
Attribute Changed.VB_HelpID = 5002

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsGradientType
    ucsGrd_Undefined = -1       ' Undefined value. Usually value is inherited from parent definition.
    ucsGrdSolid = 0             ' A solid rectange is used as a background. Color can be customized.
    ucsGrdHorizontal = 1        ' A gradient from left to right is used as a background. Source and destination colors can be customized.
    ucsGrdVertical = 2          ' A gradient from top to bottom is used as a background. Source and destination colors can be customized.
    ucsGrdBlend = 3             ' A rectangle filled with an alternating blend of two colors is used as a background. Both colors can be customized.
    ucsGrdTransparent = 4       ' Background is left untouched and shows whatever is drawn underneath.
    ucsGrdColorOffset = 5       ' A solid rectange is used as a background. Color can be customized and a hue offset and saturation/bright percents can be specified.
    ucsGrdAlphaBlend = 6        ' A solid rectange is used as a background. Color can be customized as a percent blend of two colors.
    ucsGrdStretchBitmap = 7     ' A bitmap is stretch over the background.
    ucsGrdTileBitmap = 8        ' A bitmap is tiled over the background.
End Enum

'=========================================================================
' Private constants and variables
'=========================================================================

Private Const DEF_GRADIENT_TYPE     As Long = ucsGrdSolid
Private Const DEF_COLOR             As Long = vbButtonFace
Private Const DEF_SECOND_COLOR      As Long = vbButtonFace
Private Const DEF_ALPHA             As Long = 0

Private m_eGradientType         As UcsGradientType
Private m_clrColor              As OLE_COLOR
Private m_clrSecondColor        As OLE_COLOR
Private m_lAlpha                As Long
Private m_oPicture              As StdPicture
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunc As String)
    PushError sFunc, MODULE_NAME
    PopRaiseError
End Sub

'=========================================================================
' Propeties
'=========================================================================

'Purpose: Returns or sets the display style of the background in a <b>cButton</b> object.
Property Get GradientType() As UcsGradientType
Attribute GradientType.VB_HelpID = 5005
    GradientType = m_eGradientType
End Property

Property Let GradientType(ByVal eValue As UcsGradientType)
    m_eGradientType = eValue
    RaiseEvent Changed
End Property

'Purpose: Returns or sets first background color to be used be current GradientType.
Property Get Color() As OLE_COLOR
Attribute Color.VB_HelpID = 5003
    Color = m_clrColor
End Property

Property Let Color(ByVal clrValue As OLE_COLOR)
    m_clrColor = clrValue
    RaiseEvent Changed
End Property

'Purpose: Returns or sets second background color to be used be current GradientType.
Property Get SecondColor() As OLE_COLOR
Attribute SecondColor.VB_HelpID = 5010
    SecondColor = m_clrSecondColor
End Property

Property Let SecondColor(ByVal clrValue As OLE_COLOR)
    m_clrSecondColor = clrValue
    RaiseEvent Changed
End Property

'Purpose: Returns or sets <b>Color</b> property saturation offset as a percent.
Property Get PercentSaturation() As Long
Attribute PercentSaturation.VB_HelpID = 5008
    PercentSaturation = ((m_lAlpha) And &HFF&) - 100
End Property

Property Let PercentSaturation(ByVal lValue As Long)
    m_lAlpha = (m_lAlpha And (Not &HFF&)) _
            Or pvLimitLong(lValue + 100, 0, 200&)
    RaiseEvent Changed
End Property

'Purpose: Returns or sets <b>Color</b> property brightness offset as a percent.
Property Get PercentBrightness() As Long
Attribute PercentBrightness.VB_HelpID = 5007
    PercentBrightness = ((m_lAlpha And &HFF00&) \ &H100&) - 100
End Property

Property Let PercentBrightness(ByVal lValue As Long)
    m_lAlpha = (m_lAlpha And (Not &HFF00&)) _
            Or (pvLimitLong(lValue + 100, 0, 200) * &H100&)
    RaiseEvent Changed
End Property

'Purpose: Returns or sets <b>Color</b> property hue offset.
Property Get OffsetHue() As Long
Attribute OffsetHue.VB_HelpID = 5006
    OffsetHue = (m_lAlpha And &HFFFF0000) \ &H10000
End Property

Property Let OffsetHue(ByVal lValue As Long)
    m_lAlpha = (m_lAlpha And (Not &HFFFF0000)) _
            Or (pvLimitLong(lValue, 0, 360) * &H10000)
    RaiseEvent Changed
End Property

'Purpose: Returns or sets alpha blend value between <b>Color</b> and <b>SecondColor</b> color values.
Property Get Alpha() As Long
Attribute Alpha.VB_HelpID = 5001
    Alpha = m_lAlpha
End Property

Property Let Alpha(ByVal lValue As Long)
    m_lAlpha = pvLimitLong(lValue, 0, 255)
    RaiseEvent Changed
End Property

'Purpose: Returns or sets a picture to be used by <i>ucsGrdStretchBitmap</i> and <i>ucsGrdTileBitmap</i> gradient type as a background image.
Property Get Picture() As StdPicture
Attribute Picture.VB_HelpID = 5009
    Set Picture = m_oPicture
End Property

Property Set Picture(ByVal oValue As StdPicture)
    Set m_oPicture = oValue
    RaiseEvent Changed
End Property

'Purpose: Returns or sets whether tiled bitmap is positioned absolutely or relatively
Property Get TileAbsolutePosition() As Boolean
    TileAbsolutePosition = (m_lAlpha <> 0)
End Property

Property Let TileAbsolutePosition(ByVal bValue As Boolean)
    m_lAlpha = Abs(bValue)
End Property

'Purpose: Returns or sets the contents of the <b>CGradientDef</b> object as a serialized byte array.
Public Property Get Contents() As Byte()
Attribute Contents.VB_HelpID = 5004
    Const FUNC_NAME     As String = "Contents [get]"
    Dim oBag            As PropertyBag
    
    On Error GoTo EH
    Set oBag = New PropertyBag
    With oBag
        Call .WriteProperty("T", GradientType, DEF_GRADIENT_TYPE)
        Call .WriteProperty("C", Color, DEF_COLOR)
        Call .WriteProperty("S", SecondColor, DEF_SECOND_COLOR)
        Call .WriteProperty("A", Alpha, DEF_ALPHA)
        WritePictureProperty oBag, "P", Picture
        Contents = .Contents
    End With
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Public Property Let Contents(aValue() As Byte)
    Const FUNC_NAME     As String = "Contents [let]"
    Dim oBag            As PropertyBag

    On Error GoTo EH
    Set oBag = New PropertyBag
    With oBag
        .Contents = aValue
        GradientType = .ReadProperty("T", DEF_GRADIENT_TYPE)
        Color = .ReadProperty("C", DEF_COLOR)
        SecondColor = .ReadProperty("S", DEF_SECOND_COLOR)
        Alpha = .ReadProperty("A", DEF_ALPHA)
        Set Picture = ReadPictureProperty(oBag, "P", Nothing)
    End With
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Function pvLimitLong(ByVal lValue As Long, ByVal lMin As Long, ByVal lMax As Long)
    pvLimitLong = IIf(lValue < lMin, lMin, IIf(lValue > lMax, lMax, lValue))
End Function

'=========================================================================
' Class events
'=========================================================================

Private Sub Class_Initialize()
    On Error Resume Next
    GradientType = DEF_GRADIENT_TYPE
    Color = DEF_COLOR
    SecondColor = DEF_SECOND_COLOR
    Alpha = DEF_ALPHA
    #If DebugMode Then
        DebugInit m_sDebugID, MODULE_NAME
    #End If
End Sub

#If DebugMode Then
    Private Sub Class_Terminate()
        DebugTerm m_sDebugID
    End Sub
#End If

