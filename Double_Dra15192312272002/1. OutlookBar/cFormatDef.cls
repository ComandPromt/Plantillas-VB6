VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFormatDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
Private Const MODULE_NAME As String = "cFormatDef"

'=========================================================================
' Events
'=========================================================================

'Purpose: Occurs when the contents of the current <b>cFormatDef</b> have changed.
Event Changed()
Attribute Changed.VB_HelpID = 4006
'Purpose: Occurs when the font of the <b>Outlook Bar</b> control is needed. Implemented internally.
Event ControlFont(oFont As StdFont)
Attribute ControlFont.VB_HelpID = 4008

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsFormatHorAlignmentStyle
    ucsFha_Undefined = -1       ' Undefined value. Usually value is inherited from parent definition.
    ucsFhaLeft = 0              ' Horizontally align to the left of the bounding rectangle.
    ucsFhaCenter = 1            ' Horizontally align to the center of the bounding rectangle.
    ucsFhaRight = 2             ' Horizontally align to the right of the bounding rectangle.
End Enum

Public Enum UcsFormatVertAlignmentStyle
    ucsFva_Undefined = -1       ' Undefined value. Usually value is inherited from parent definition.
    ucsFvaTop = 0               ' Vertically align to the top of the bounding rectangle.
    ucsFvaMiddle = 1            ' Vertically align to the middle of the bounding rectangle.
    ucsFvaBottom = 2            ' Vertically align to the bottom of the bounding rectangle.
End Enum

Public Enum UcsFormatBorderStyle
    ucsFbd_Undefined = -1       ' Undefined value. Usually value is inherited from parent definition.
    ucsFbdNone = 0              ' No border is drawn around the item.
    ucsFbdFixed = 1             ' A single point wide flat border is drawn around the item. Color can be customized.
    ucsFbdSingle3D = 2          ' A single point wide 3d border is drawn around the item. Shadow color can be customized. Light color is always vb3DHighlight.
    ucsFbdDouble3D = 3          ' A double point wide 3d border is drawn around the item. Colors can not be customized.
End Enum

'=========================================================================
' Private constants and variables
'=========================================================================

Private Const DEF_BORDER                As Long = ucsFbd_Undefined
Private Const DEF_BORDERCOLOR           As Long = -1
Private Const DEF_FORECOLOR             As Long = -1
Private Const DEF_HOR_ALIGNMENT         As Long = ucsFha_Undefined
Private Const DEF_VERT_ALIGNMENT        As Long = ucsFva_Undefined
Private Const DEF_OFFSET_X              As Long = LNG_UNDEFINED
Private Const DEF_OFFSET_Y              As Long = LNG_UNDEFINED
Private Const DEF_PADDING               As Long = LNG_UNDEFINED
Private Const DEF_BORDER_SUNKEN         As Long = ucsTri_Undefined
Private Const DEF_NAME                  As String = ""

Private WithEvents m_oFontDef       As cFontDef
Attribute m_oFontDef.VB_VarHelpID = -1
Private m_clrFore                   As OLE_COLOR
Private WithEvents m_oBack          As cGradientDef
Attribute m_oBack.VB_VarHelpID = -1
Private m_eHorAlignment             As UcsFormatHorAlignmentStyle
Private m_eVertAlignment            As UcsFormatVertAlignmentStyle
Private m_eBorder                   As UcsFormatBorderStyle
Private m_clrBorder                 As OLE_COLOR
Private m_lOffsetX                  As Long
Private m_lOffsetY                  As Long
Private m_lPadding                  As Long
Private m_oParentFmt                As cFormatDef
Private m_sName                     As String
Private m_oParentFont               As StdFont
Private m_bInSet                    As Boolean
Private m_eBorderSunken             As UcsTriState
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

'Purpose: Returns or sets a <b>StdFont</b> object that is the font to be used by <b>cFormatDef</b>.
Property Get Font() As StdFont
Attribute Font.VB_HelpID = 4009
    Set Font = m_oFontDef.ApplyTo(ParentFont)
End Property

Property Set Font(ByVal oValue As StdFont)
    m_oFontDef.ExtractDelta oValue, ParentFont
End Property

'Purpose: Returns a <b>cFontDef</b> object that is the font to be used by <b>cFormatDef</b>.
Property Get FontDef() As cFontDef
Attribute FontDef.VB_HelpID = 4010
    Set FontDef = m_oFontDef
End Property

'Purpose: Returns or sets the foreground color for the text in to be formated.
Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_HelpID = 4011
    ForeColor = m_clrFore
End Property

Property Let ForeColor(oValue As OLE_COLOR)
    m_clrFore = oValue
    If Not m_bInSet Then
        RaiseEvent Changed
    End If
End Property

'Purpose: Returns a <b>cGradientDef</b> object representing a definition of the background to be formated.
Property Get BackGradient() As cGradientDef
Attribute BackGradient.VB_HelpID = 4001
    Set BackGradient = m_oBack
End Property

'Purpose: Returns or sets a value that determines the horizontal alignment of text or images to be formated.
Property Get HorAlignment() As UcsFormatHorAlignmentStyle
Attribute HorAlignment.VB_HelpID = 4013
    HorAlignment = m_eHorAlignment
End Property

Property Let HorAlignment(ByVal eValue As UcsFormatHorAlignmentStyle)
    '--- limit values
    If eValue <> LNG_UNDEFINED Then
        eValue = pvLimitLong(eValue, -1, 2)
    End If
    '--- touch only if different
    If m_eHorAlignment <> eValue Then
        m_eHorAlignment = eValue
        If Not m_bInSet Then
            RaiseEvent Changed
        End If
    End If
End Property

'Purpose: Returns or sets a value that determines the vertical alignment of text or images to be formated.
Property Get VertAlignment() As UcsFormatVertAlignmentStyle
Attribute VertAlignment.VB_HelpID = 4021
    VertAlignment = m_eVertAlignment
End Property

Property Let VertAlignment(ByVal eValue As UcsFormatVertAlignmentStyle)
    '--- limit values
    If eValue <> LNG_UNDEFINED Then
        eValue = pvLimitLong(eValue, -1, 2)
    End If
    '--- touch only if different
    If m_eVertAlignment <> eValue Then
        m_eVertAlignment = eValue
        If Not m_bInSet Then
            RaiseEvent Changed
        End If
    End If
End Property

'Purpose: Returns or sets the border style of an <b>cFormatDef</b> object.
Property Get Border() As UcsFormatBorderStyle
Attribute Border.VB_HelpID = 4002
    Border = m_eBorder
End Property

Property Let Border(ByVal eValue As UcsFormatBorderStyle)
    '--- limit values
    If eValue <> LNG_UNDEFINED Then
        eValue = pvLimitLong(eValue, -1, 3)
    End If
    '--- touch only if different
    If m_eBorder <> eValue Then
        m_eBorder = eValue
        If Not m_bInSet Then
            RaiseEvent Changed
        End If
    End If
End Property

'Purpose: Returns the size of the current border style in pixels.
Property Get BorderSize() As Long
Attribute BorderSize.VB_HelpID = 4004
    Select Case Border
    Case ucsFbdFixed, ucsFbdSingle3D
        BorderSize = 1
    Case ucsFbdDouble3D
        BorderSize = 2
    End Select
End Property

'Purpose: Returns or sets the color to be used for current border style.
Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_HelpID = 4003
    BorderColor = m_clrBorder
End Property

Property Let BorderColor(ByVal clrValue As OLE_COLOR)
    m_clrBorder = clrValue
    If Not m_bInSet Then
        RaiseEvent Changed
    End If
End Property

Property Get BorderSunken() As UcsTriState
Attribute BorderSunken.VB_HelpID = 4005
    BorderSunken = m_eBorderSunken
End Property

Property Let BorderSunken(ByVal eValue As UcsTriState)
    m_eBorderSunken = eValue
    If Not m_bInSet Then
        RaiseEvent Changed
    End If
End Property

'Purpose: Returns or sets a long value indicating horizontal offset of text or image in pixels.
Property Get OffsetX() As Long
Attribute OffsetX.VB_HelpID = 4015
    OffsetX = m_lOffsetX
End Property

Property Let OffsetX(ByVal lValue As Long)
    '--- limit values
    If lValue <> LNG_UNDEFINED Then
        lValue = pvLimitLong(lValue, -128, 128)
    End If
    '--- touch only if different
    If m_lOffsetX <> lValue Then
        m_lOffsetX = lValue
        If Not m_bInSet Then
            RaiseEvent Changed
        End If
    End If
End Property

'Purpose: Returns or sets a long value indicating vertical offset of text or image in pixels.
Property Get OffsetY() As Long
Attribute OffsetY.VB_HelpID = 4016
    OffsetY = m_lOffsetY
End Property

Property Let OffsetY(ByVal lValue As Long)
    '--- limit values
    If lValue <> LNG_UNDEFINED Then
        lValue = pvLimitLong(lValue, -128, 128)
    End If
    '--- touch only if different
    If m_lOffsetY <> lValue Then
        m_lOffsetY = lValue
        If Not m_bInSet Then
            RaiseEvent Changed
        End If
    End If
End Property

'Purpose: Returns or sets a long value indicating number of pixels left as a whitespace on every border.
Property Get Padding() As Long
Attribute Padding.VB_HelpID = 4017
    Padding = m_lPadding
End Property

Property Let Padding(ByVal lValue As Long)
    '--- limit values
    If lValue <> LNG_UNDEFINED Then
        lValue = pvLimitLong(lValue, 0, 128)
    End If
    '--- touch only if different
    If m_lPadding <> lValue Then
        m_lPadding = lValue
        If Not m_bInSet Then
            RaiseEvent Changed
        End If
    End If
End Property

'Purpose: Returns or sets a references to a parent <b>cFormatDef</b> object.
Property Get ParentFmt() As cFormatDef
Attribute ParentFmt.VB_HelpID = 4018
    Set ParentFmt = m_oParentFmt
End Property

Property Set ParentFmt(ByVal oValue As cFormatDef)
    Set m_oParentFmt = oValue
    If Not m_bInSet Then
        RaiseEvent Changed
    End If
End Property

'Purpose: Returns or sets the font used by parent <b>cFormatDef</b> as a <b>StdFont</b> object.
Property Get ParentFont() As StdFont
Attribute ParentFont.VB_HelpID = 4019
    Const FUNC_NAME     As String = "ParentFont [get]"
    
    On Error GoTo EH
    If ParentFmt Is Nothing Then
        Dim oFont As StdFont
        RaiseEvent ControlFont(oFont)
        If Not oFont Is Nothing Then
            Set ParentFont = oFont
        Else
            Set ParentFont = m_oParentFont
        End If
    Else
        Set ParentFont = ParentFmt.Font
    End If
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Property Set ParentFont(oValue As StdFont)
    Set m_oParentFont = oValue
End Property

'Purpose: Returns or sets a string that uniquely identifies a <b>cFormatDef</b> object in an <b>Outlook Bar</b> control.
Property Get Name() As String
Attribute Name.VB_HelpID = 4014
    Name = m_sName
End Property

Property Let Name(sValue As String)
    m_sName = sValue
    If Not m_bInSet Then
        RaiseEvent Changed
    End If
End Property

'Purpose: Returns or sets a string that includes the names of parent <b>cFormatDef</b> objects, fully traversing parent hierarchy.
Property Get FullName() As String
Attribute FullName.VB_HelpID = 4012
    If Not ParentFmt Is Nothing Then
        FullName = ParentFmt.FullName
    End If
    FullName = FullName & "_" & Name
End Property

'Purpose: Returns or sets the contents of the <b>cFormatDef</b> object as a serialized byte array.
Public Property Get Contents() As Byte()
Attribute Contents.VB_HelpID = 4007
    Const FUNC_NAME As String = "Contents [get]"
    
    On Error GoTo EH
    With New PropertyBag
        Call .WriteProperty("B", BackGradient.Contents, DEF_GRADIENT.Contents)
        Call .WriteProperty("O", Border, DEF_BORDER)
        Call .WriteProperty("R", BorderColor, DEF_BORDERCOLOR)
        Call .WriteProperty("BS", BorderSunken, DEF_BORDER_SUNKEN)
        Call .WriteProperty("FD", FontDef.Contents, DEF_FONTDEF.Contents)
        Call .WriteProperty("C", ForeColor, DEF_FORECOLOR)
        Call .WriteProperty("H", HorAlignment, DEF_HOR_ALIGNMENT)
        Call .WriteProperty("V", VertAlignment, DEF_VERT_ALIGNMENT)
        Call .WriteProperty("X", OffsetX, DEF_OFFSET_X)
        Call .WriteProperty("Y", OffsetY, DEF_OFFSET_Y)
        Call .WriteProperty("P", Padding, DEF_PADDING)
        Contents = .Contents
    End With
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Public Property Let Contents(aValue() As Byte)
    Const FUNC_NAME As String = "Contents [let]"
    
    With New PropertyBag
        .Contents = aValue
        m_bInSet = True
        BackGradient.Contents = .ReadProperty("B", DEF_GRADIENT.Contents)
        Border = .ReadProperty("O", DEF_BORDER)
        BorderColor = .ReadProperty("R", DEF_BORDERCOLOR)
        BorderSunken = .ReadProperty("BS", DEF_BORDER_SUNKEN)
        FontDef.Contents = .ReadProperty("FD", DEF_FONTDEF.Contents)
        ForeColor = .ReadProperty("C", DEF_FORECOLOR)
        HorAlignment = .ReadProperty("H", DEF_HOR_ALIGNMENT)
        VertAlignment = .ReadProperty("V", DEF_VERT_ALIGNMENT)
        OffsetX = .ReadProperty("X", DEF_OFFSET_X)
        OffsetY = .ReadProperty("Y", DEF_OFFSET_Y)
        Padding = .ReadProperty("P", DEF_PADDING)
'        Name = .ReadProperty("N", DEF_NAME)
        m_bInSet = False
    End With
    RaiseEvent Changed
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get DEF_GRADIENT() As cGradientDef
    Set DEF_GRADIENT = New cGradientDef
End Property

Private Property Get DEF_FONT() As StdFont
    Set DEF_FONT = New StdFont
    DEF_FONT.Name = "Tahoma"
    DEF_FONT.Size = 8
End Property

Private Property Get DEF_FONTDEF() As cFontDef
    Set DEF_FONTDEF = New cFontDef
End Property

'=========================================================================
' Methods
'=========================================================================

'Purpose: Replaces undefined properties of the current definition with values from a <b>cFormatDef</b> object and returns the result.
Public Function Render(Optional ByVal oFmt As cFormatDef) As cFormatDef
Attribute Render.VB_HelpID = 4020
    Const FUNC_NAME     As String = "Render"
    
    On Error GoTo EH
    '--- try to render parent format
    If oFmt Is Nothing And Not ParentFmt Is Nothing Then
        Set oFmt = ParentFmt.Render
    End If
    Set Render = New cFormatDef
    If Not oFmt Is Nothing Then
        '--- parent format present
        With Render
            .Border = IIf(Border <> ucsFbd_Undefined, Border, oFmt.Border)
            .BorderColor = IIf(BorderColor <> -1, BorderColor, oFmt.BorderColor)
            .BorderSunken = IIf(BorderSunken <> ucsTri_Undefined, BorderSunken, oFmt.BorderSunken)
            .ForeColor = IIf(ForeColor <> -1, ForeColor, oFmt.ForeColor)
            .HorAlignment = IIf(HorAlignment <> ucsFha_Undefined, HorAlignment, oFmt.HorAlignment)
            .VertAlignment = IIf(VertAlignment <> ucsFva_Undefined, VertAlignment, oFmt.VertAlignment)
            .OffsetX = IIf(OffsetX <> LNG_UNDEFINED, OffsetX, oFmt.OffsetX)
            .OffsetY = IIf(OffsetY <> LNG_UNDEFINED, OffsetY, oFmt.OffsetY)
            .Padding = IIf(Padding <> LNG_UNDEFINED, Padding, oFmt.Padding)
            Set .ParentFont = Font
            Set .Font = Font
            If BackGradient.GradientType <> ucsGrd_Undefined Then
                .BackGradient.Contents = BackGradient.Contents
            Else
                .BackGradient.Contents = oFmt.BackGradient.Contents
            End If
        End With
    Else
        '--- no parent format found (supplied)
        With Render
            .Border = IIf(Border <> ucsFbd_Undefined, Border, ucsFbdSingle3D)
            .BorderColor = IIf(BorderColor <> LNG_UNDEFINED, BorderColor, vbButtonShadow)
            .BorderSunken = IIf(BorderSunken <> ucsTri_Undefined, BorderSunken, ucsTriFalse)
            .ForeColor = IIf(ForeColor <> LNG_UNDEFINED, ForeColor, vbButtonText)
            .HorAlignment = IIf(HorAlignment <> ucsFha_Undefined, HorAlignment, ucsFhaLeft)
            .VertAlignment = IIf(VertAlignment <> ucsFva_Undefined, VertAlignment, ucsFvaTop)
            .OffsetX = IIf(OffsetX <> LNG_UNDEFINED, OffsetX, 0)
            .OffsetY = IIf(OffsetY <> LNG_UNDEFINED, OffsetY, 0)
            .Padding = IIf(Padding <> LNG_UNDEFINED, Padding, 0)
            Set .ParentFont = DEF_FONT
            Set .Font = DEF_FONT
            .BackGradient.Contents = BackGradient.Contents
        End With
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvLimitLong(ByVal lValue As Long, ByVal lMin As Long, ByVal lMax As Long)
    pvLimitLong = IIf(lValue < lMin, lMin, IIf(lValue > lMax, lMax, lValue))
End Function

'=========================================================================
' Class events
'=========================================================================

Private Sub Class_Initialize()
    On Error Resume Next
    m_bInSet = True
    Set m_oFontDef = DEF_FONTDEF
    Set m_oBack = DEF_GRADIENT
    ForeColor = DEF_FORECOLOR
    Border = DEF_BORDER
    BorderColor = DEF_BORDERCOLOR
    BorderSunken = DEF_BORDER_SUNKEN
    HorAlignment = DEF_HOR_ALIGNMENT
    VertAlignment = DEF_VERT_ALIGNMENT
    OffsetX = DEF_OFFSET_X
    OffsetY = DEF_OFFSET_Y
    Padding = DEF_PADDING
    m_bInSet = False
    RaiseEvent Changed
    #If DebugMode Then
        DebugInit m_sDebugID, MODULE_NAME
    #End If
End Sub

#If DebugMode Then
    Private Sub Class_Terminate()
        DebugTerm m_sDebugID
    End Sub
#End If

Private Sub m_oBack_Changed()
    If Not m_bInSet Then
        RaiseEvent Changed
    End If
End Sub

Private Sub m_oFontDef_Changed()
    If Not m_bInSet Then
        RaiseEvent Changed
    End If
End Sub
